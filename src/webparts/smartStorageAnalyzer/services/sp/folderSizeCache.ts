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
