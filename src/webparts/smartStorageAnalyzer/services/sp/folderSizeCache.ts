import { FolderStorageNode } from '../../models/models';

// Folder sizes are expensive to (re)compute (a StorageMetrics call per
// folder, sometimes a full recursive walk) but don't change every minute, so
// a short-lived cache means a page refresh doesn't throw all that work away.
// sessionStorage rather than localStorage: it should not survive to a new
// browser session where the data is more likely stale, and it's naturally
// scoped per-tab so switching sites in another tab can't cross-contaminate.
const CACHE_PREFIX = 'sp-smart-storage-analyzer-folder-cache::';
const TTL_MS = 10 * 60 * 1000; // 10 minutes

interface CacheEntry {
  timestamp: number;
  nodes: FolderStorageNode[];
}

function cacheKey(siteUrl: string, folderUrl: string): string {
  return `${CACHE_PREFIX}${siteUrl}::${folderUrl}`;
}

export function getCachedFolderChildren(siteUrl: string, folderUrl: string): FolderStorageNode[] | undefined {
  try {
    const raw = sessionStorage.getItem(cacheKey(siteUrl, folderUrl));
    if (!raw) return undefined;
    const entry = JSON.parse(raw) as CacheEntry;
    if (Date.now() - entry.timestamp > TTL_MS) {
      sessionStorage.removeItem(cacheKey(siteUrl, folderUrl));
      return undefined;
    }
    return entry.nodes;
  } catch {
    return undefined;
  }
}

// Drops every cached folder size for a site (or all sites when omitted).
//
// Needed because the TTL alone leaves a user stuck: a folder that came back
// unmeasurable (throttling) or truncated at the walk budget is otherwise
// unretryable for up to 10 minutes, and the advice to "try again in a moment"
// has nothing to act on. Iterating the keys rather than clearing
// sessionStorage wholesale so unrelated app state on the page survives.
export function clearCachedFolderChildren(siteUrl?: string): void {
  try {
    const prefix = siteUrl ? `${CACHE_PREFIX}${siteUrl}::` : CACHE_PREFIX;
    // Snapshot the keys first — removing while iterating sessionStorage by
    // index shifts subsequent indices and silently skips entries.
    const doomed: string[] = [];
    for (let i = 0; i < sessionStorage.length; i++) {
      const key = sessionStorage.key(i);
      if (key && key.startsWith(prefix)) doomed.push(key);
    }
    doomed.forEach((key) => sessionStorage.removeItem(key));
  } catch {
    // sessionStorage unavailable (private browsing) — nothing cached to clear.
  }
}

export function setCachedFolderChildren(siteUrl: string, folderUrl: string, nodes: FolderStorageNode[]): void {
  try {
    const entry: CacheEntry = { timestamp: Date.now(), nodes };
    sessionStorage.setItem(cacheKey(siteUrl, folderUrl), JSON.stringify(entry));
  } catch {
    // sessionStorage full (rare — entries are small) or unavailable (private
    // browsing in some browsers) — caching is a pure optimization, so a
    // write failure just means this folder won't be fast next time, not a
    // functional problem.
  }
}
