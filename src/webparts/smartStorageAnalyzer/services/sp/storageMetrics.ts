import { SpApiClient, folderApi } from './spCore';
import { FolderStorageNode, LibraryInfo } from '../../models/models';
import { fetchLibraryItems, fetchSiteUsers, FlatItem, LibraryFetchError } from './listItems';
import { fetchRecycleBinItems } from './recycleBin';
import { aggregateLibrary, LibraryAggregate } from './folderAggregate';

// Data layer behind the Explorer's Tree View / List View.
//
// This was a lazy, per-folder recursive walk: for each folder the user
// opened, probe StorageMetrics per child and fall back to walking that
// child's whole subtree (two requests per folder in it) whenever the probe
// came back stale — which it commonly does. Cost scaled with the shape of
// the tree, so on a large archive the biggest folders were exactly the ones
// that never finished.
//
// It is now one flat sweep per LIBRARY (listItems.ts), aggregated in memory
// (folderAggregate.ts). That inverts the tradeoff: the first folder opened in
// a library pays for the whole library (~1 request per 5,000 items), and
// every drill-down after that — sizes, file lists, child folders, at any
// depth — is a map lookup with no network at all. Sizes are exact rather
// than "≥ a floor", because they are summed from the items themselves.

interface RawMetrics {
  totalSizeBytes: number;
  fileCount: number;
  lastModified?: string;
}

// Fired while a library sweep is in progress, with the running item count.
// Pages are 5,000 items, so this fires about once per request rather than
// continuously — no throttling needed on top.
export type WalkProgress = (itemsFetched: number) => void;

export interface WalkOptions {
  // Cooperative cancellation. getJsonPaged checks it between pages, so a
  // cancel takes effect within one in-flight request; whatever pages already
  // arrived are kept and reported as partial.
  signal?: AbortSignal;
  onWalkProgress?: WalkProgress;
}

// ── Per-library aggregate cache ───────────────────────────────────────────
// In memory only, deliberately. The previous sessionStorage cache stored a
// per-folder children breakdown; an aggregate for a large library is far too
// big for sessionStorage's few-megabyte budget, and re-sweeping on a page
// reload now costs ~25 requests rather than a full traversal, so persisting
// it buys much less than it used to.
interface CachedAggregate {
  aggregate: LibraryAggregate;
  // A sweep stopped by the user. Kept so the view can render what arrived,
  // but never treated as authoritative: sizes derived from it are marked
  // approximate, and it is dropped rather than reused for a later request.
  partial: boolean;
}

const aggregateCache = new Map<string, CachedAggregate>();
const userCache = new Map<string, Map<number, { title: string; loginName: string }>>();

function cacheKey(siteUrl: string, libraryUrl: string): string {
  return `${siteUrl}::${libraryUrl}`;
}

export function clearAggregateCache(siteUrl?: string): void {
  if (!siteUrl) {
    aggregateCache.clear();
    userCache.clear();
    return;
  }
  const prefix = `${siteUrl}::`;
  for (const key of Array.from(aggregateCache.keys())) {
    if (key.startsWith(prefix)) aggregateCache.delete(key);
  }
  userCache.delete(siteUrl);
}

/**
 * The library's aggregate, sweeping it if it isn't cached. A partial
 * (canceled) result is returned to the caller but not reused for the next
 * call, so retrying after a cancel re-sweeps rather than serving a stale
 * half-answer forever.
 */
export async function getLibraryAggregate(
  client: SpApiClient,
  siteUrl: string,
  library: LibraryInfo,
  options?: WalkOptions,
): Promise<CachedAggregate> {
  const key = cacheKey(siteUrl, library.serverRelativeUrl);
  const cached = aggregateCache.get(key);
  if (cached && !cached.partial) return cached;

  // The Recycle Bin has no list id and isn't addressed via `_api/web/lists`
  // at all — route it to its own fetch (recycleBin.ts) rather than
  // listItems.ts's `_api/web/lists(guid'...')/items` sweep.
  const items = library.isRecycleBin
    ? await fetchRecycleBinItems(client, siteUrl, {
      signal: options?.signal,
      onProgress: options?.onWalkProgress,
    })
    : await fetchLibraryItems(client, siteUrl, library, {
      signal: options?.signal,
      onProgress: options?.onWalkProgress,
    });
  const partial = !!options?.signal?.aborted;
  const entry: CachedAggregate = {
    aggregate: aggregateLibrary(library.serverRelativeUrl, items),
    partial,
  };
  // Only a complete sweep is worth keeping — caching a canceled one would
  // make the cancel sticky, so the user would have to Refresh to undo it.
  if (!partial) aggregateCache.set(key, entry);
  return entry;
}

async function getUsers(
  client: SpApiClient,
  siteUrl: string,
  signal?: AbortSignal,
): Promise<Map<number, { title: string; loginName: string }>> {
  const cached = userCache.get(siteUrl);
  if (cached) return cached;
  const users = await fetchSiteUsers(client, siteUrl, signal);
  userCache.set(siteUrl, users);
  return users;
}

// GetFolderByServerRelativePath(...)/StorageMetrics is the classic, still-
// functional (but UNDOCUMENTED) endpoint behind Site Settings → Storage
// Metrics. It returns a folder's RECURSIVE rollup size/file count in a
// single call. It is not in Microsoft's official REST reference, so every
// call here is defensive: any unexpected shape or failure returns undefined
// and the caller falls back to a real measurement instead of breaking.
//
// Still worth trying FIRST at library level: one request per library beats
// sweeping a library the user may never open. It is not used below folder
// level any more — once a library has been swept, the aggregate has an exact
// figure for every folder in it, which is strictly better than a probe.
export async function getStorageMetrics(
  client: SpApiClient,
  siteUrl: string,
  serverRelativeUrl: string,
): Promise<RawMetrics | undefined> {
  try {
    const data = await client.getJson(`${folderApi(siteUrl, serverRelativeUrl)}/StorageMetrics`);
    // Legacy verbose OData nests the payload under `d` — unwrapped here so a
    // verbose-mode response doesn't miss both fields and force an
    // unnecessary full sweep as if the endpoint were unavailable.
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

// Fired once the sweep resolves, with the child count, so the caller can
// swap a spinner for content. Retained (rather than removed) because the
// Explorer's loading gate is built around it; there is no longer a
// meaningful per-child progression to report, since all children resolve
// together the moment the aggregate is built.
export type FolderChildProgress = (done: number, total: number, node?: FolderStorageNode) => void;

/**
 * Immediate child folders of `parentServerRelativeUrl`, each with an exact
 * recursive size. Sweeps the library on first use, then answers from memory.
 */
export async function getFolderChildren(
  client: SpApiClient,
  siteUrl: string,
  library: LibraryInfo,
  parentServerRelativeUrl: string,
  onProgress?: FolderChildProgress,
  options?: WalkOptions,
): Promise<FolderStorageNode[]> {
  onProgress?.(0, 0);
  const { aggregate, partial } = await getLibraryAggregate(client, siteUrl, library, options);
  const parent = aggregate.folders.get(parentServerRelativeUrl.replace(/\/+$/, ''));
  if (!parent) return [];

  const nodes: FolderStorageNode[] = parent.childFolders.map((path) => {
    const rollup = aggregate.folders.get(path)!;
    return {
      name: rollup.name,
      serverRelativeUrl: rollup.serverRelativeUrl,
      totalSizeBytes: rollup.totalSizeBytes,
      fileCount: rollup.fileCount,
      lastModified: rollup.lastModified,
      // Summed from the items themselves rather than probed, so this is an
      // exact figure — unless the sweep that produced it was canceled.
      sizeSource: 'estimate',
      sizeApproximate: partial,
      versionSizeBytes: rollup.versionSizeBytes,
      children: [],
      hasChildren: rollup.childFolders.length > 0,
    };
  });
  onProgress?.(nodes.length, nodes.length);
  return nodes;
}

/**
 * Immediate files under a folder. Comes straight out of the aggregate — the
 * sweep already read every file in the library, so this costs nothing at
 * all, including version history: both size (SMTotalFileStreamSize) and an
 * approximate retained-version count (OData__UIVersionString's major version
 * number, minus 1 — see listItems.ts's VERSION_LABEL_FIELD comment for what
 * this does and doesn't capture exactly) arrive in the same bulk sweep. This
 * used to also fire one request PER FILE just for an exact count; that's
 * gone, confirmed as the right trade for a storage audit rather than a
 * version-history audit.
 */
export async function getFolderFiles(
  client: SpApiClient,
  siteUrl: string,
  library: LibraryInfo,
  serverRelativeUrl: string,
  options?: WalkOptions,
): Promise<{
  name: string; serverRelativeUrl: string; sizeBytes: number; timeCreated: string;
  timeLastModified: string; authorLoginName?: string; authorDisplayName?: string;
  versionSizeBytes?: number; versionCount?: number;
}[]> {
  const { aggregate } = await getLibraryAggregate(client, siteUrl, library, options);
  const files: FlatItem[] = aggregate.filesByParent.get(serverRelativeUrl.replace(/\/+$/, '')) ?? [];
  const users = await getUsers(client, siteUrl, options?.signal);

  return files.map((f) => {
    // Recycle Bin rows arrive with the name/login already resolved
    // (DeletedByTitle/Email are plain strings, not lookup ids) — prefer
    // those directly rather than trying (and failing) an authorId lookup
    // that was never set for them.
    const author = f.authorDisplayName != null || f.authorLoginName != null
      ? { title: f.authorDisplayName, loginName: f.authorLoginName }
      : (f.authorId != null ? users.get(f.authorId) : undefined);
    return {
      name: f.name,
      serverRelativeUrl: f.fileRef,
      sizeBytes: f.sizeBytes,
      timeCreated: f.created,
      timeLastModified: f.modified,
      authorLoginName: author?.loginName,
      authorDisplayName: author?.title,
      // Both free from the sweep — see the function comment above.
      versionSizeBytes: f.versionSizeBytes,
      versionCount: f.versionCountApprox,
    };
  });
}

export { LibraryFetchError };
