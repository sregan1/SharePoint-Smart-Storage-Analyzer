import { SpApiClient, valueArray } from './spCore';
import { LibraryInfo } from '../../models/models';
import { FlatItem, LibraryFetchError, ITEMS_PAGE_SIZE } from './listItems';

// Surfaces a site's Recycle Bin as an additional pseudo-library, so it shows
// up everywhere a real library does: the site-root treemap, the library
// switcher, Tree View / List View drill-down, and Storage Report scans.
//
// This is `_api/web/RecycleBin` — the FIRST-STAGE recycle bin only, scoped to
// the current web, which is exactly what a Site Owner can already see in
// classic Site Settings → Recycle bin. The SECOND-STAGE ("site collection")
// recycle bin lives at `_api/site/RecycleBin` and requires Site Collection
// Administrator — a materially higher right than the ManageWeb/
// ManagePermissions this app already gates everything else behind — so it is
// deliberately out of scope rather than silently under- or over-reaching on
// permissions.
//
// ASSUMPTION FLAGGED FOR VALIDATION: deleting a folder creates one recycle
// bin row per descendant file (each carrying its own real Size), plus
// separate zero-weight rows for the folder(s) themselves — this is the
// documented CSOM shape (RecycleBinItemType has both File/ListItem leaf types
// and Folder/FolderWithLists container types, plus a CascadeParent type that
// only makes sense if children are itemized separately). If a tenant instead
// rolls a deleted folder's total size onto the folder's own row, sizes here
// would double count. To make that failure mode impossible rather than just
// unlikely, only File/ListItem rows ever contribute bytes — Folder/
// FolderWithLists rows are always treated as zero-weight containers, exactly
// like the folder rows in listItems.ts.

export const RECYCLE_BIN_TITLE = 'Recycle Bin';

// Subset of RecycleBinItemType (Microsoft.SharePoint.Client) relevant here.
// The rest (FileVersion=2, List=4, Attachment=7, ListItemVersion=8,
// CascadeParent=9, Web=10, App=11) are deliberately excluded below rather
// than guessed at — see the module comment.
const ITEM_TYPE_FILE = 1;
const ITEM_TYPE_LIST_ITEM = 3;
const ITEM_TYPE_FOLDER = 5;
const ITEM_TYPE_FOLDER_WITH_LISTS = 6;

function sitePath(siteUrl: string): string {
  return new URL(siteUrl).pathname.replace(/\/+$/, '');
}

export function recycleBinRoot(siteUrl: string): string {
  return `${sitePath(siteUrl)}/${RECYCLE_BIN_TITLE}`;
}

export function makeRecycleBinLibrary(siteUrl: string): LibraryInfo {
  return {
    title: RECYCLE_BIN_TITLE,
    serverRelativeUrl: recycleBinRoot(siteUrl),
    // Not a real list template — LIBRARY_TEMPLATES (spCore.ts) is used only
    // to build the `_api/web/lists` $filter, which this pseudo-library never
    // goes through, so there is no real value that belongs here.
    baseTemplate: -1,
    isRecycleBin: true,
  };
}

// Re-roots an item's ORIGINAL location (DirName, e.g. "/sites/x/Shared
// Documents/Old") under the Recycle Bin pseudo-path, so opening "Recycle
// Bin" shows top-level entries matching the site's real libraries, and
// drilling in shows the folder structure items used to live in — reusing
// the Explorer's existing folder UI with no special-casing.
function reparentUnderRecycleBin(siteUrl: string, dirName: string): string {
  const root = recycleBinRoot(siteUrl);
  const site = sitePath(siteUrl);
  const normalized = `/${dirName.replace(/^\/+/, '')}`;
  if (normalized === site) return root; // deleted directly from the site root
  if (normalized.startsWith(`${site}/`)) {
    return `${root}${normalized.substring(site.length)}`;
  }
  // Unexpected shape (e.g. a different site's path, or an already-relative
  // DirName on a tenant that formats it differently) — nest it under a
  // catch-all rather than dropping the item or throwing.
  return `${root}/Other${normalized}`;
}

function toFlatItem(siteUrl: string, raw: any): FlatItem | undefined {
  const itemType = Number(raw.ItemType ?? -1);
  const isFile = itemType === ITEM_TYPE_FILE || itemType === ITEM_TYPE_LIST_ITEM;
  const isFolder = itemType === ITEM_TYPE_FOLDER || itemType === ITEM_TYPE_FOLDER_WITH_LISTS;
  if (!isFile && !isFolder) return undefined; // see module comment

  const leafName = String(raw.LeafName ?? raw.Title ?? '');
  const parent = reparentUnderRecycleBin(siteUrl, String(raw.DirName ?? ''));
  const deletedDate = raw.DeletedDate as string | undefined;

  return {
    fileRef: `${parent}/${leafName}`,
    name: leafName,
    isFolder,
    // Folder/FolderWithLists rows never carry bytes — see module comment.
    sizeBytes: isFolder ? 0 : Math.max(0, Number(raw.Size ?? 0)),
    // The Recycle Bin's REST shape has no separate created/modified — DeletedDate
    // is the only timestamp available, so archival tiering for Recycle Bin
    // content reflects "time since deleted" rather than "time since last
    // edited". For content that is, by definition, already marked for
    // removal, that is arguably the more useful of the two anyway.
    created: deletedDate ?? '',
    modified: deletedDate ?? '',
    authorDisplayName: raw.DeletedByTitle as string | undefined,
    authorLoginName: raw.DeletedByEmail as string | undefined,
  };
}

export interface FetchRecycleBinOptions {
  signal?: AbortSignal;
  onProgress?: (fetchedSoFar: number) => void;
}

function pageUrl(siteUrl: string, skip: number): string {
  return `${siteUrl}/_api/web/RecycleBin`
    + '?$select=LeafName,Title,DirName,ItemType,Size,DeletedDate,DeletedByTitle,DeletedByEmail'
    + `&$orderby=Id asc&$top=${ITEMS_PAGE_SIZE}&$skip=${skip}`;
}

// Sweeps the (first-stage) Recycle Bin, paged. Unlike listItems.ts's use of
// getJsonPagedMeta, this loop drives its own $skip — the RecycleBin resource
// is not documented to return a server-side paging token the way
// `_api/web/lists(...)/items` does, so skip-based paging (checking for a
// nextLink first, in case a tenant does provide one, and falling back to
// incrementing $skip when it doesn't) is the safe default here.
export async function fetchRecycleBinItems(
  client: SpApiClient,
  siteUrl: string,
  options?: FetchRecycleBinOptions,
): Promise<FlatItem[]> {
  const items: FlatItem[] = [];
  let fetched = 0;
  let next: string | undefined = pageUrl(siteUrl, 0);
  let skip = 0;
  // Same order-of-magnitude safety valve as listItems.ts — a recycle bin
  // needing more pages than this holds over 2,000,000 items.
  for (let page = 0; page < 400 && next && !options?.signal?.aborted; page++) {
    let data: any;
    try {
      // skipBatch: a 5,000-row page is the wrong shape to coalesce with
      // small requests — see SpApiClient.getJson.
      data = await client.getJson(next, true);
    } catch (err: any) {
      throw new LibraryFetchError(RECYCLE_BIN_TITLE, err?.message ?? String(err));
    }
    const rows = valueArray(data);
    for (const row of rows) {
      const item = toFlatItem(siteUrl, row);
      if (item) items.push(item);
    }
    fetched += rows.length;
    options?.onProgress?.(fetched);

    // Prefer a server-provided token if a tenant happens to return one;
    // otherwise fall back to advancing our own $skip. Either way, a short
    // page (or an empty one) means this was the last one.
    const nextLink = data?.['odata.nextLink'] ?? data?.['@odata.nextLink'];
    if (nextLink) {
      next = nextLink;
    } else if (rows.length < ITEMS_PAGE_SIZE) {
      next = undefined;
    } else {
      skip += ITEMS_PAGE_SIZE;
      next = pageUrl(siteUrl, skip);
    }
  }
  return items;
}
