import { SpApiClient, folderApi, SYSTEM_FOLDER_NAMES, TaskQueue } from './spCore';
import { FolderStorageNode } from '../../models/models';
import { getCachedFolderChildren, setCachedFolderChildren } from './folderSizeCache';

interface RawMetrics {
  totalSizeBytes: number;
  fileCount: number;
  lastModified?: string;
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
): void {
  queue.add(async () => {
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
        walkFolderSubtree(client, siteUrl, child, queue, onRootDone);
      }
    } catch {
      // Inaccessible subtree — this node keeps whatever partial total it
      // had (usually 0) and still finishes/caches normally so its parent
      // isn't left waiting forever on a branch that will never resolve.
      node.errored = true;
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
    onRootDone(node.acc);
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
  if (node.errored) parent.errored = true;
  parent.childNodes.push({
    name: node.name,
    serverRelativeUrl: node.url,
    totalSizeBytes: node.acc.totalSizeBytes,
    fileCount: node.acc.fileCount,
    lastModified: node.acc.lastModified ?? node.timeLastModified,
    sizeSource: 'estimate',
    children: [],
    hasChildren: true,
  });
  parent.pendingChildren--;
  if (parent.pendingChildren === 0) finishNode(siteUrl, parent, onRootDone);
}

// Live recursive rollup used when StorageMetrics is unavailable or reports
// zero for a folder that may have content. Accepts an optional shared
// TaskQueue so a caller resolving many folders at once (getFolderChildren,
// getLibrariesWithStats) can run every fallback walk through ONE bounded
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
): Promise<RawMetrics> {
  const queue = sharedQueue ?? new TaskQueue(client.scanConcurrency);
  const root: WalkNode = {
    url: serverRelativeUrl,
    name: '',
    acc: { totalSizeBytes: 0, fileCount: 0 },
    pendingChildren: 0,
    childNodes: [],
  };
  return new Promise<RawMetrics>((resolve) => {
    walkFolderSubtree(client, siteUrl, root, queue, resolve);
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

  let done = 0;
  const tasks = rawFolders.map((f: any) => async () => {
    const metrics = await getStorageMetrics(client, siteUrl, f.ServerRelativeUrl);
    // Trust StorageMetrics only when it reports actual content. Its rollup
    // is computed by a periodic background job that lags behind recent
    // activity, so "0 bytes" for a folder that really has files is common —
    // and indistinguishable from a genuinely empty folder without checking.
    // When it says 0 (or is unavailable), settle it with a live walk using
    // the same Files/Folders calls the Storage Report scan relies on; for
    // truly empty folders that walk is only a couple of cheap requests.
    const useMetrics = !!metrics && metrics.totalSizeBytes > 0;
    const resolved = useMetrics ? metrics! : await getShallowEstimate(client, siteUrl, f.ServerRelativeUrl, fallbackQueue);
    if (!useMetrics && resolved.totalSizeBytes === 0 && resolved.fileCount === 0) {
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
      sizeSource: useMetrics ? 'storageMetrics' : 'estimate', // 'estimate' here is a live walk, still exact for readable subtrees
      children: [],
      hasChildren: true, // resolved lazily on next expand; harmless if it turns out empty
    };
    done++;
    onProgress?.(done, total, node);
    return node;
  });

  // The StorageMetrics probe above is a single, fixed-size batch — exactly
  // one lightweight call per immediate folder, however many that is — unlike
  // the fallback walk, which can fan out into arbitrarily many requests as
  // it recurses into subfolders. That open-ended fan-out is what caused
  // throttling before, so the walk stays capped at the user's configured
  // concurrency; the fixed, bounded probe batch is safe to run faster. Scaled
  // relative to the user's own setting (not a fixed floor) so someone who has
  // deliberately turned concurrency down to avoid throttling on their tenant
  // isn't overridden — capped at 10 so a high setting doesn't get doubled
  // again on top of an already-generous value.
  const probeConcurrency = Math.min(client.scanConcurrency * 2, 10);
  const results = await client.runConcurrent(tasks, probeConcurrency);
  const nodes = results.filter((n): n is FolderStorageNode => !!n).sort((a, b) => b.totalSizeBytes - a.totalSizeBytes);
  setCachedFolderChildren(siteUrl, parentServerRelativeUrl, nodes);
  return nodes;
}

// Immediate files under a folder (used for the file-list panel and the
// treemap's leaf cells) — a single shallow call, not part of any recursive
// walk, so it's cheap to call whenever the user drills into a folder.
export async function getFolderFiles(
  client: SpApiClient,
  siteUrl: string,
  serverRelativeUrl: string,
): Promise<{ name: string; serverRelativeUrl: string; sizeBytes: number; timeCreated: string; timeLastModified: string; authorLoginName?: string; authorDisplayName?: string }[]> {
  const files = await client.getJsonPaged(
    `${folderApi(siteUrl, serverRelativeUrl)}/Files?$select=Name,ServerRelativeUrl,Length,TimeCreated,TimeLastModified,Author/Title,Author/LoginName&$expand=Author&$top=5000`,
  );
  return files.map((f: any) => ({
    name: f.Name,
    serverRelativeUrl: f.ServerRelativeUrl,
    sizeBytes: Number(f.Length ?? 0),
    timeCreated: f.TimeCreated,
    timeLastModified: f.TimeLastModified,
    authorLoginName: f.Author?.LoginName,
    authorDisplayName: f.Author?.Title,
  }));
}
