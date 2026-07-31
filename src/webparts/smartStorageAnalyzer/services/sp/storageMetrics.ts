import { SpApiClient, folderApi, odata, SYSTEM_FOLDER_NAMES, TaskQueue } from './spCore';
import { FolderStorageNode } from '../../models/models';
import { getCachedFolderChildren, setCachedFolderChildren } from './folderSizeCache';

// Sums SP.FileVersion.Size across a file's retained old versions, and counts
// them — the /Versions collection excludes the current version, so the count
// is free from the same call that already fetches size. Same shape as
// fileWalk.ts's fetchVersionInfo, kept local here since ExplorerView only
// ever needs it for one folder's immediate files at a time (no shared
// walk-level queue to plug into).
async function fetchVersionInfo(
  client: SpApiClient,
  siteUrl: string,
  fileServerRelativeUrl: string,
): Promise<{ sizeBytes: number; count: number }> {
  const url = `${siteUrl}/_api/web/GetFileByServerRelativePath(decodedUrl='${encodeURIComponent(
    odata(fileServerRelativeUrl),
  )}')/Versions?$select=Size`;
  const versions = await client.getJsonPaged(url);
  return {
    sizeBytes: versions.reduce((sum, v) => sum + Number(v.Size ?? 0), 0),
    count: versions.length,
  };
}

interface RawMetrics {
  totalSizeBytes: number;
  fileCount: number;
  lastModified?: string;
  // Set when a live walk (getShallowEstimate) couldn't fully resolve this
  // folder's own subtree (throttling exhausted, a transient error, a
  // permissions issue) — the totals above are then whatever partial amount
  // it had when it gave up, not a confirmed result. Never set for a
  // StorageMetrics-sourced result (that path either succeeds or is treated
  // as unavailable, with no partial/uncertain state).
  errored?: boolean;
  // The first failure encountered in this folder's subtree. Preserved rather
  // than discarded so callers can report the real cause instead of guessing.
  errorMessage?: string;
  // Set when the walk hit its folder budget and stopped descending, so the
  // total is a genuine floor ("at least this much") rather than exact.
  truncated?: boolean;
}

// Bounds how many folder listings ONE getFolderChildren/getLibraryRollups
// call may issue across all of its fallback walks combined.
//
// Without this, the Explorer would fully recurse a subtree per immediate child
// whose StorageMetrics rollup reads 0 — and on a real archive that is
// catastrophic: a single folder-open on this tenant walked three subtrees more
// than seven levels deep, tens of thousands of listings, until SharePoint
// blocked the account outright and every folder came back unmeasurable. An
// interactive view only needs its immediate children's sizes; spending
// unbounded requests to make those exact is the wrong trade every time.
//
// The count alone used to BE the bound, fixed at 400 regardless of tenant
// speed. That's really a wall-clock limit wearing a request-count costume:
// the walk advances roughly `scanConcurrency` folder-visits per network
// round-trip, so the same 400 costs ~10s on a fast tenant but multiple
// minutes on one that's actively throttling (each throttle event can gate
// every request for up to 60s — see SpApiClient's shared circuit breaker in
// spCore.ts). So the deadline below is now the PRIMARY bound — it's what
// actually keeps one user action's wait time predictable — and the count is
// a backstop tuned to how hard the user has told the app it can push this
// tenant (scanConcurrency), not a fixed guess.
//
// Both are shared across the whole call (not per child/library) because
// that's what actually bounds the cost of one user action.
export const FALLBACK_FOLDERS_PER_CONCURRENCY = 150;
export const MIN_FALLBACK_FOLDERS = 300;
export const MAX_FALLBACK_FOLDERS = 2400;
export const FALLBACK_WALK_DEADLINE_MS = 45_000;
export const LIBRARY_ROLLUP_WALK_DEADLINE_MS = 20_000;
// The retry pass (see getFolderChildren) runs serially, one folder at a time,
// only after the user has already waited out the entire first pass — it must
// not be allowed to run long by itself on top of that.
export const RETRY_WALK_DEADLINE_MS = 15_000;

export interface WalkBudget {
  remaining: number;
  deadlineAtMs: number;
  truncated: boolean;
}

// Single source of truth for creating a budget, so the several creation
// sites (getShallowEstimate's own default, getFolderChildren,
// getLibraryRollups, the retry pass) can't drift out of sync with each
// other or with client.scanConcurrency.
export function newWalkBudget(client: SpApiClient, deadlineMs: number): WalkBudget {
  const perConcurrency = client.scanConcurrency * FALLBACK_FOLDERS_PER_CONCURRENCY;
  return {
    remaining: Math.min(MAX_FALLBACK_FOLDERS, Math.max(MIN_FALLBACK_FOLDERS, perConcurrency)),
    deadlineAtMs: Date.now() + deadlineMs,
    truncated: false,
  };
}

// GetFolderByServerRelativePath(...)/StorageMetrics is the classic, still-
// functional (but UNDOCUMENTED) endpoint behind Site Settings → Storage
// Metrics. It returns a folder's RECURSIVE rollup size/file count in a
// single call — no need to walk every file underneath. It is not in
// Microsoft's official REST reference, so every call here is defensive:
// any unexpected shape or failure returns undefined and the caller falls
// back to a shallow (one-level, non-recursive) estimate instead of breaking.
export async function getStorageMetrics(
  client: SpApiClient,
  siteUrl: string,
  serverRelativeUrl: string,
): Promise<RawMetrics | undefined> {
  try {
    const data = await client.getJson(`${folderApi(siteUrl, serverRelativeUrl)}/StorageMetrics`);
    // Legacy verbose OData nests the payload under `d` — unwrapped here so a
    // verbose-mode response doesn't miss both fields and force an
    // unnecessary full walk as if the endpoint were unavailable.
    const m = data?.d ?? data;
    const totalSizeBytes = Number(m?.TotalSize ?? m?.TotalFileStreamSize);
    if (!isFinite(totalSizeBytes)) return undefined;
    return {
      totalSizeBytes,
      fileCount: Number(m?.TotalFileCount ?? 0),
      lastModified: m?.LastModified as string | undefined,
    };
  } catch {
    return undefined;
  }
}

// One folder's slot in a live walk. Each node gets its OWN accumulator
// (not one shared total for the whole walk) so that once a node's entire
// subtree finishes, its immediate children's individual totals — not just
// the grand total — are known and can be cached for reuse.
interface WalkNode {
  url: string;
  name: string;
  timeLastModified?: string;
  parent?: WalkNode;
  acc: RawMetrics;
  pendingChildren: number;
  childNodes: FolderStorageNode[];
  // Set when this node's own listing failed, or any descendant's did — see
  // finishNode. An errored node's immediate-children breakdown is
  // incomplete, so it must never be written to folderSizeCache: caching it
  // would let a transient throttle/403 poison the cache for its full TTL,
  // with the Explorer confidently rendering the folder as empty rather than
  // unreadable.
  errored?: boolean;
  errorMessage?: string;
  truncated?: boolean;
}

// Recursively walks one folder's subtree on a caller-supplied TaskQueue.
// The walk already reads every subfolder's own Files/Folders listing along
// the way to compute the root's total — this used to throw that per-folder
// detail away and keep only the grand total, so drilling into any folder
// touched during the walk repeated the exact same work moments later. Now
// every node's own immediate-children breakdown is cached (folderSizeCache)
// as soon as it's known, so a subsequent Explorer visit to any folder this
// walk passed through is an instant cache hit instead of a re-walk.
//
// Recursion never awaits a child directly — that would tie up a queue slot
// per ancestor for the whole subtree's duration and can deadlock a deep
// tree (N ancestors permanently holding N slots, with nothing left to run
// the (N+1)th level). Each node's task fetches its own listing, fires its
// children off as new independent queue tasks, and returns immediately;
// completion is instead signaled by `pendingChildren` reaching 0 and
// propagated up via `parent`, so a slot is always free for whatever's next
// regardless of how deep the tree goes.
function walkFolderSubtree(
  client: SpApiClient,
  siteUrl: string,
  node: WalkNode,
  queue: TaskQueue,
  onRootDone: (acc: RawMetrics) => void,
  budget: WalkBudget,
): void {
  queue.add(async () => {
    // Budget is consumed on ENTRY (one listing pair per folder visited), and
    // checked before descending further — so an exhausted budget stops the
    // walk cleanly with a partial-but-honest total rather than aborting it.
    // The deadline is the primary bound (see FALLBACK_WALK_DEADLINE_MS) —
    // it's what actually keeps one user action's wait time predictable
    // regardless of tenant speed; the folder count is a backstop.
    if (budget.remaining <= 0 || Date.now() >= budget.deadlineAtMs) {
      budget.truncated = true;
      node.truncated = true;
      finishNode(siteUrl, node, onRootDone);
      return;
    }
    budget.remaining--;
    try {
      const [filesData, foldersData] = await Promise.all([
        client.getJsonPaged(`${folderApi(siteUrl, node.url)}/Files?$select=Length,TimeLastModified&$top=5000`),
        client.getJsonPaged(`${folderApi(siteUrl, node.url)}/Folders?$select=Name,ServerRelativeUrl,TimeLastModified&$top=5000`),
      ]);
      for (const f of filesData) {
        node.acc.totalSizeBytes += Number(f.Length ?? 0);
        node.acc.fileCount++;
        const m = f.TimeLastModified as string;
        if (m && (!node.acc.lastModified || m > node.acc.lastModified)) node.acc.lastModified = m;
      }

      const subFolders = foldersData.filter(
        (f: any) => !SYSTEM_FOLDER_NAMES.has(String(f.Name).toLowerCase()),
      );
      if (subFolders.length === 0) {
        finishNode(siteUrl, node, onRootDone);
        return;
      }
      node.pendingChildren = subFolders.length;
      for (const f of subFolders) {
        const child: WalkNode = {
          url: f.ServerRelativeUrl,
          name: f.Name,
          timeLastModified: f.TimeLastModified,
          parent: node,
          acc: { totalSizeBytes: 0, fileCount: 0 },
          pendingChildren: 0,
          childNodes: [],
        };
        walkFolderSubtree(client, siteUrl, child, queue, onRootDone, budget);
      }
    } catch (err: any) {
      // Inaccessible subtree — this node keeps whatever partial total it
      // had (usually 0) and still finishes/caches normally so its parent
      // isn't left waiting forever on a branch that will never resolve.
      //
      // The error itself used to be dropped on the floor (a bare `catch {}`),
      // which meant the UI could only guess at the cause and told everyone to
      // lower their concurrency regardless of whether throttling was involved.
      node.errored = true;
      node.errorMessage = err?.message ?? String(err);
      // eslint-disable-next-line no-console
      console.warn(`[SmartStorageAnalyzer] Could not measure ${node.url}: ${node.errorMessage}`);
      finishNode(siteUrl, node, onRootDone);
    }
  });
}

// Called once a node's own subtree is fully resolved (no children, or every
// child has already called this). Caches the node's immediate-children
// breakdown, folds its total into its parent's accumulator, and propagates
// the same completion check up the tree — or, at the root, hands the final
// total back to the caller.
function finishNode(siteUrl: string, node: WalkNode, onRootDone: (acc: RawMetrics) => void): void {
  if (!node.errored) {
    setCachedFolderChildren(siteUrl, node.url, node.childNodes.sort((a, b) => b.totalSizeBytes - a.totalSizeBytes));
  }

  const parent = node.parent;
  if (!parent) {
    // node.errored lives on the WalkNode, not the RawMetrics acc handed back
    // to the caller — fold it in here so getShallowEstimate's caller can
    // tell "genuinely empty" apart from "the walk gave up partway through."
    onRootDone({
      ...node.acc,
      errored: node.errored,
      errorMessage: node.errorMessage,
      truncated: node.truncated,
    });
    return;
  }
  parent.acc.totalSizeBytes += node.acc.totalSizeBytes;
  parent.acc.fileCount += node.acc.fileCount;
  if (node.acc.lastModified && (!parent.acc.lastModified || node.acc.lastModified > parent.acc.lastModified)) {
    parent.acc.lastModified = node.acc.lastModified;
  }
  // Propagate so an ancestor of a failed branch is not cached either — its
  // own childNodes list would otherwise look complete while one entry's
  // total silently reflects only a partial subtree.
  if (node.errored) {
    parent.errored = true;
    // Keep the first cause seen rather than the last — the earliest failure in
    // a subtree is usually the root cause, later ones its knock-on effects.
    parent.errorMessage = parent.errorMessage ?? node.errorMessage;
  }
  // An ancestor of a truncated branch is itself only a floor, and must not be
  // cached as an exact figure either (see the !node.errored cache guard above,
  // which this deliberately mirrors).
  if (node.truncated) parent.truncated = true;
  parent.childNodes.push({
    name: node.name,
    serverRelativeUrl: node.url,
    totalSizeBytes: node.acc.totalSizeBytes,
    fileCount: node.acc.fileCount,
    lastModified: node.acc.lastModified ?? node.timeLastModified,
    sizeSource: 'estimate',
    sizeApproximate: node.truncated,
    children: [],
    hasChildren: true,
  });
  parent.pendingChildren--;
  if (parent.pendingChildren === 0) finishNode(siteUrl, parent, onRootDone);
}

// Live recursive rollup used when StorageMetrics is unavailable or reports
// zero for a folder that may have content. Accepts an optional shared
// TaskQueue so a caller resolving many folders at once (getFolderChildren,
// getLibraryRollups) can run every fallback walk through ONE bounded
// pool instead of each call opening its own — nesting bounded pools inside
// an already-bounded outer loop multiplies concurrency (outer limit ×
// inner limit) instead of capping it, which is what was driving requests
// into SharePoint throttling and turning "load a folder" into a ~minute-long
// wait. Per-folder errors skip just that subtree, so an inaccessible branch
// yields partial totals, not zeros.
export async function getShallowEstimate(
  client: SpApiClient,
  siteUrl: string,
  serverRelativeUrl: string,
  sharedQueue?: TaskQueue,
  sharedBudget?: WalkBudget,
): Promise<RawMetrics> {
  const queue = sharedQueue ?? new TaskQueue(client.scanConcurrency);
  // A caller that resolves several folders at once passes ONE budget so the
  // ceiling applies to the whole user action rather than per folder. No
  // caller currently omits sharedBudget for a multi-folder resolution, but
  // this default keeps a lone standalone call self-bounded regardless.
  const budget = sharedBudget ?? newWalkBudget(client, FALLBACK_WALK_DEADLINE_MS);
  const root: WalkNode = {
    url: serverRelativeUrl,
    name: '',
    acc: { totalSizeBytes: 0, fileCount: 0 },
    pendingChildren: 0,
    childNodes: [],
  };
  return new Promise<RawMetrics>((resolve) => {
    walkFolderSubtree(client, siteUrl, root, queue, resolve, budget);
  });
}

// Fired as each immediate child's size resolves: `done`/`total` for a
// progress readout, plus the resolved `node` itself so a caller can render
// folders as they arrive instead of waiting for the whole batch. Also fired
// once with done=0 right after the (fast) folder listing resolves, so a
// caller can show "0 of N" immediately instead of a blank gap before the
// first child finishes.
export type FolderChildProgress = (done: number, total: number, node?: FolderStorageNode) => void;

// Lists immediate child folders and resolves each one's rolled-up size,
// preferring StorageMetrics (recursive, fast) and falling back to a shallow
// per-folder estimate when it's unavailable. Concurrency is capped via the
// client's runConcurrent so expanding a folder with many children doesn't
// fire an unbounded burst of requests. Results are cached for a short TTL
// (see folderSizeCache.ts) so revisiting the same folder — including after a
// page refresh, which wipes the in-memory cache the Explorer keeps — doesn't
// redo this work; a cache hit returns immediately with no progress callbacks.
export async function getFolderChildren(
  client: SpApiClient,
  siteUrl: string,
  parentServerRelativeUrl: string,
  onProgress?: FolderChildProgress,
): Promise<FolderStorageNode[]> {
  const cached = getCachedFolderChildren(siteUrl, parentServerRelativeUrl);
  if (cached) return cached;

  const listData = await client.getJsonPaged(
    `${folderApi(siteUrl, parentServerRelativeUrl)}/Folders?$select=Name,ServerRelativeUrl,ItemCount,TimeLastModified&$top=5000`,
  );
  const rawFolders = listData.filter(
    (f: any) => !SYSTEM_FOLDER_NAMES.has(String(f.Name).toLowerCase()),
  );
  const total = rawFolders.length;
  onProgress?.(0, total);

  // One shared queue for every fallback walk triggered below. Without it,
  // each folder that needs a live walk would open its own bounded pool on
  // top of the runConcurrent pool already bounding this loop, multiplying
  // concurrency instead of capping it (see getShallowEstimate).
  const fallbackQueue = new TaskQueue(client.scanConcurrency);
  // One budget for every fallback walk this call triggers — see
  // FALLBACK_WALK_DEADLINE_MS for why an unbounded walk per child is untenable.
  const fallbackBudget = newWalkBudget(client, FALLBACK_WALK_DEADLINE_MS);

  let done = 0;
  const tasks = rawFolders.map((f: any) => async () => {
    // ItemCount from the folder listing already answers "is this folder
    // completely empty?" (it counts files AND subfolders), so a genuinely
    // empty folder needs no probe and no walk at all. This used to fall
    // through to a full walk to establish a result we already had — pure
    // request volume spent confirming zero, once per empty folder, and every
    // one of those requests was another chance to get throttled.
    if (Number(f.ItemCount ?? -1) === 0) {
      const emptyNode: FolderStorageNode = {
        name: f.Name,
        serverRelativeUrl: f.ServerRelativeUrl,
        totalSizeBytes: 0,
        fileCount: 0,
        lastModified: f.TimeLastModified,
        sizeSource: 'estimate',
        children: [],
        hasChildren: false,
      };
      done++;
      onProgress?.(done, total, emptyNode);
      return emptyNode;
    }

    const metrics = await getStorageMetrics(client, siteUrl, f.ServerRelativeUrl);
    // Trust StorageMetrics only when it reports actual content. Its rollup
    // is computed by a periodic background job that lags behind recent
    // activity, so "0 bytes" for a folder that really has files is common —
    // and indistinguishable from a genuinely empty folder without checking.
    // When it says 0 (or is unavailable), settle it with a live walk using
    // the same Files/Folders calls the Storage Report scan relies on.
    const useMetrics = !!metrics && metrics.totalSizeBytes > 0;
    const resolved = useMetrics
      ? metrics!
      : await getShallowEstimate(client, siteUrl, f.ServerRelativeUrl, fallbackQueue, fallbackBudget);
    // resolved.errored means the live walk gave up partway through (throttling
    // exhausted, a transient error) — its totals are whatever partial amount
    // it had at that point, NOT a confirmed result, and must be presented as
    // such rather than silently rendered as a genuinely empty folder.
    //
    // A truncated walk with a ZERO total is a distinct third case: the shared
    // budget/deadline (see fallbackBudget above) was already exhausted by
    // OTHER children in this same batch before this folder's own walk got to
    // list anything at all — not "measured and found empty", but "never
    // looked at". Left as an approximate 0, this folder would render as
    // "≥ 0 B", sort to the bottom of the list, and paint as a near-invisible
    // treemap square — presenting a folder nobody looked at as confirmed
    // near-empty, exactly the failure mode sizeUnknown exists to prevent.
    const neverMeasured = !useMetrics && !!resolved.truncated
      && resolved.totalSizeBytes === 0 && resolved.fileCount === 0;
    const sizeUnknown = !useMetrics && (!!resolved.errored || neverMeasured);
    if (!useMetrics && !sizeUnknown && resolved.totalSizeBytes === 0 && resolved.fileCount === 0) {
      // eslint-disable-next-line no-console
      console.debug(
        `[SmartStorageAnalyzer] ${f.ServerRelativeUrl}: StorageMetrics ${metrics ? 'reported 0 bytes' : 'unavailable'}` +
        ' and live walk found no files — treating as empty.',
      );
    }
    const node: FolderStorageNode = {
      name: f.Name,
      serverRelativeUrl: f.ServerRelativeUrl,
      totalSizeBytes: resolved.totalSizeBytes,
      fileCount: resolved.fileCount,
      lastModified: resolved.lastModified ?? f.TimeLastModified,
      sizeSource: useMetrics ? 'storageMetrics' : (sizeUnknown ? 'error' : 'estimate'), // 'estimate' here is a live walk, exact unless truncated below
      sizeErrorMessage: sizeUnknown
        ? (resolved.errorMessage
          ?? "Not measured — this view's measurement budget was used up by other folders. Open this folder directly to measure it.")
        : undefined,
      // sizeUnknown already covers the neverMeasured case (a real 0, not an
      // approximate one) — only mark approximate when there's an actual
      // nonzero partial total to be a floor for.
      sizeApproximate: !useMetrics && !sizeUnknown && !!resolved.truncated,
      children: [],
      hasChildren: true, // resolved lazily on next expand; harmless if it turns out empty
    };
    done++;
    onProgress?.(done, total, node);
    return node;
  });

  // No longer doubled. This used to run at min(scanConcurrency * 2, 10) on the
  // reasoning that the probe batch is fixed-size and therefore safe to run
  // faster than the open-ended fallback walk. That held in isolation, but the
  // probe pool and the fallback walks it spawns are alive at the SAME time, so
  // the real peak was the sum (up to 10 + scanConcurrency) — on a large tenant,
  // enough to sit in sustained throttling. SpApiClient now enforces one global
  // in-flight ceiling across every pool and shrinks it while throttled, so
  // asking for extra here buys nothing and only deepens the queue.
  const results = await client.runConcurrent(tasks, client.scanConcurrency);
  const nodes = results.filter((n): n is FolderStorageNode => !!n);

  // Second pass for anything that couldn't be measured, walked ONE AT A TIME.
  // Most first-pass failures are throttling or transient, and by the time the
  // first pass has finished the shared throttle gate has usually cleared and
  // the tenant is no longer shedding load — so a single serial retry recovers
  // the majority of them. Serial specifically: retrying these concurrently is
  // what caused them to fail in the first place, and there are only ever a
  // handful, so the cost is small and bounded.
  const failed = nodes.filter((n) => n.sizeSource === 'error');
  if (failed.length > 0) {
    // eslint-disable-next-line no-console
    console.warn(`[SmartStorageAnalyzer] Retrying ${failed.length} unmeasured folder(s) one at a time…`);
    const retryQueue = new TaskQueue(1);
    // One shared budget for the WHOLE retry pass, not a fresh one per folder —
    // this only runs after the user already waited out the entire first pass,
    // so it gets a short, single window (RETRY_WALK_DEADLINE_MS) rather than
    // potentially failed.length independent 45s windows stacked serially.
    const retryBudget = newWalkBudget(client, RETRY_WALK_DEADLINE_MS);
    for (const node of failed) {
      const retry = await getShallowEstimate(client, siteUrl, node.serverRelativeUrl, retryQueue, retryBudget);
      if (retry.errored) {
        node.sizeErrorMessage = retry.errorMessage ?? node.sizeErrorMessage;
        continue;
      }
      node.totalSizeBytes = retry.totalSizeBytes;
      node.fileCount = retry.fileCount;
      node.lastModified = retry.lastModified ?? node.lastModified;
      node.sizeSource = 'estimate';
      node.sizeErrorMessage = undefined;
      // The shared retry budget/deadline can itself run out partway through
      // the retry pass (new in this pass — previously the retry had no budget
      // at all) — a truncated retry result is a real floor, not confirmed
      // exact, and must carry that forward the same as the first pass does.
      node.sizeApproximate = !!retry.truncated;
    }
    const stillFailed = nodes.filter((n) => n.sizeSource === 'error').length;
    // eslint-disable-next-line no-console
    console.warn(
      `[SmartStorageAnalyzer] Retry recovered ${failed.length - stillFailed} of ${failed.length} folder(s).`,
    );
  }

  nodes.sort((a, b) => b.totalSizeBytes - a.totalSizeBytes);
  // Same reasoning as the internal walk's own per-node cache guard: caching a
  // batch containing an 'error' (uncertain-size) child would keep showing it
  // as unknown for the cache's full TTL even once throttling has cleared up.
  // Leaving the whole batch uncached costs one re-probe on the next visit —
  // cheap next to that folder staying wrong for minutes.
  if (!nodes.some((n) => n.sizeSource === 'error')) {
    setCachedFolderChildren(siteUrl, parentServerRelativeUrl, nodes);
  }
  return nodes;
}

// Immediate files under a folder (used for the file-list panel and the
// treemap's leaf cells) — a single shallow call, not part of any recursive
// walk, so it's cheap to call whenever the user drills into a folder.
export async function getFolderFiles(
  client: SpApiClient,
  siteUrl: string,
  serverRelativeUrl: string,
  includeVersionHistory = false,
): Promise<{ name: string; serverRelativeUrl: string; sizeBytes: number; timeCreated: string; timeLastModified: string; authorLoginName?: string; authorDisplayName?: string; versionSizeBytes?: number; versionCount?: number }[]> {
  const files = await client.getJsonPaged(
    `${folderApi(siteUrl, serverRelativeUrl)}/Files?$select=Name,ServerRelativeUrl,Length,TimeCreated,TimeLastModified,Author/Title,Author/LoginName&$expand=Author&$top=5000`,
  );
  const rows = files.map((f: any) => ({
    name: f.Name,
    serverRelativeUrl: f.ServerRelativeUrl,
    sizeBytes: Number(f.Length ?? 0),
    timeCreated: f.TimeCreated,
    timeLastModified: f.TimeLastModified,
    authorLoginName: f.Author?.LoginName,
    authorDisplayName: f.Author?.Title,
    versionSizeBytes: undefined as number | undefined,
    versionCount: undefined as number | undefined,
  }));

  if (includeVersionHistory && rows.length > 0) {
    // Bounded to this one folder's (typically small) file list — no
    // shared walk-level queue needed since callers only ever load one
    // folder's immediate files at a time.
    await client.runConcurrent(
      rows.map((row) => async () => {
        try {
          const info = await fetchVersionInfo(client, siteUrl, row.serverRelativeUrl);
          row.versionSizeBytes = info.sizeBytes;
          row.versionCount = info.count;
        } catch {
          row.versionSizeBytes = undefined;
          row.versionCount = undefined;
        }
      }),
      Math.max(1, Math.floor(client.scanConcurrency / 2)),
    );
  }

  return rows;
}
