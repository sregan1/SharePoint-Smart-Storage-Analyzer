import { FlatItem } from './listItems';

// Turns one library's flat item list (listItems.ts) into recursive per-folder
// rollups, entirely in memory.
//
// This is the half of the rewrite that makes folder sizes cheap. Previously
// every folder's total came from either a StorageMetrics probe (one request
// per folder, and frequently stale) or a recursive subtree walk (two requests
// per folder in the subtree). Here, a single pass over rows the app already
// fetched produces an EXACT recursive total for every folder in the library
// at once — no further requests, and nothing to be stale relative to the
// items it was computed from.

export interface FolderRollup {
  // Server-relative path, no trailing slash.
  serverRelativeUrl: string;
  name: string;
  // Recursive: this folder and everything beneath it.
  totalSizeBytes: number;
  // Retained-version bytes beneath this folder, recursive. undefined when
  // the library didn't report version sizes (see listItems.ts field
  // negotiation) — deliberately distinct from 0, which means "reported, and
  // there is no retained version content here".
  versionSizeBytes?: number;
  fileCount: number;
  // Most recent Modified anywhere beneath this folder.
  lastModified?: string;
  // Immediate children only, largest first.
  childFolders: string[];
}

export interface LibraryAggregate {
  // Keyed by server-relative path (no trailing slash).
  folders: Map<string, FolderRollup>;
  root: FolderRollup;
  // Files keyed by their PARENT folder path, so the file list for a folder
  // is a lookup rather than a scan of every item in the library.
  filesByParent: Map<string, FlatItem[]>;
}

function parentPath(path: string): string {
  const i = path.lastIndexOf('/');
  return i <= 0 ? '' : path.substring(0, i);
}

function leafName(path: string): string {
  const i = path.lastIndexOf('/');
  return i < 0 ? path : path.substring(i + 1);
}

function newRollup(path: string): FolderRollup {
  return {
    serverRelativeUrl: path,
    name: leafName(path),
    totalSizeBytes: 0,
    versionSizeBytes: undefined,
    fileCount: 0,
    lastModified: undefined,
    childFolders: [],
  };
}

/**
 * Builds every folder's recursive rollup for one library.
 *
 * `rootUrl` is the library's own server-relative root; paths above it are
 * ignored so a library nested under a managed path doesn't accumulate
 * totals onto folders that aren't part of this library.
 */
export function aggregateLibrary(rootUrl: string, items: FlatItem[]): LibraryAggregate {
  const root = rootUrl.replace(/\/+$/, '');
  const folders = new Map<string, FolderRollup>();
  const filesByParent = new Map<string, FlatItem[]>();
  folders.set(root, newRollup(root));

  // Pass 1: make sure every folder that EXISTS has an entry, even if it is
  // completely empty. A folder with no files anywhere beneath it still needs
  // to render (as 0 B) rather than vanish from the treemap — the old walk
  // got this right and it would be a regression to lose it.
  for (const item of items) {
    if (!item.isFolder) continue;
    const path = item.fileRef.replace(/\/+$/, '');
    if (path === root || !path.startsWith(`${root}/`)) continue;
    if (!folders.has(path)) folders.set(path, newRollup(path));
  }

  // Pass 2: fold every file into each of its ancestors, up to (and
  // including) the library root. Walking the ancestor chain per file is what
  // makes these totals RECURSIVE — the cost is O(items x depth), which for
  // 100k items at depth 10 is a million cheap map lookups: microseconds
  // against the seconds the network sweep itself took.
  for (const item of items) {
    if (item.isFolder) continue;
    const path = item.fileRef;
    if (!path.startsWith(`${root}/`)) continue;

    const parent = parentPath(path);
    const siblings = filesByParent.get(parent);
    if (siblings) siblings.push(item);
    else filesByParent.set(parent, [item]);

    let cursor = parent;
    // Guard against a pathologically deep or malformed path looping forever.
    for (let depth = 0; depth < 200; depth++) {
      let rollup = folders.get(cursor);
      if (!rollup) {
        // A file whose parent folder never appeared as its own item — rare,
        // but real (a folder can be filtered out of the item set while its
        // contents are not). Create it so the size is still attributed
        // somewhere visible instead of being silently dropped.
        rollup = newRollup(cursor);
        folders.set(cursor, rollup);
      }
      rollup.totalSizeBytes += item.sizeBytes;
      rollup.fileCount++;
      if (item.versionSizeBytes != null) {
        rollup.versionSizeBytes = (rollup.versionSizeBytes ?? 0) + item.versionSizeBytes;
      }
      if (item.modified && (!rollup.lastModified || item.modified > rollup.lastModified)) {
        rollup.lastModified = item.modified;
      }
      if (cursor === root) break;
      const next = parentPath(cursor);
      // Stop if we've walked out of the library (or can't go further).
      if (!next || next === cursor || !(next === root || next.startsWith(`${root}/`))) break;
      cursor = next;
    }
  }

  // Pass 3: wire up immediate-child links, largest first, so the Explorer
  // can render one level without re-scanning the whole map per drill-down.
  for (const [path, rollup] of folders) {
    if (path === root) continue;
    const parent = folders.get(parentPath(path));
    if (parent) parent.childFolders.push(path);
  }
  for (const rollup of folders.values()) {
    rollup.childFolders.sort(
      (a, b) => (folders.get(b)?.totalSizeBytes ?? 0) - (folders.get(a)?.totalSizeBytes ?? 0),
    );
  }

  return { folders, root: folders.get(root)!, filesByParent };
}
