import { SpApiClient } from './spCore';
import { FileEntry, LibraryInfo, ScanOptions } from '../../models/models';
import { ageInDays, classify } from '../../utils/archivalClassification';
import { fetchLibraryItems, FlatItem } from './listItems';
import { fetchRecycleBinItems } from './recycleBin';

export interface WalkLibraryResult {
  // Set when the library's own bulk query failed and nothing could be read
  // from it. There is deliberately no recursive-walk fallback any more, so
  // this is terminal for the library and the caller reports it as skipped.
  failed?: { url: string; error: string };
  // Per-file version fetches that failed during the per-file fallback (see
  // listItems.ts's fillMissingVersionSizes) — only ever nonzero when
  // options.includeVersionHistory triggered that fallback for this library.
  skippedVersions: number;
}

export type SiteUsers = Map<number, { title: string; loginName: string }>;

// Enumerates one library for the Storage Report.
//
// This used to be a recursive folder walk: two requests per folder (Files +
// Folders), so cost tracked the SHAPE of the tree rather than the amount of
// content. On a real 2.3TB archive that managed about four files a second —
// hours per library — because a deep tree full of small folders spends every
// round trip discovering folders rather than reading content.
//
// It is now one flat, paged sweep of the library's items (see listItems.ts):
// roughly one request per 5,000 items, independent of how the folders are
// arranged. Version history — both size (SMTotalFileStreamSize) and an
// approximate retained-version COUNT (OData__UIVersionString's major version
// number, minus 1) — normally arrives in that same sweep at no extra cost.
// This used to also do a separate request PER FILE just to get an exact
// count; on a real 192,978-item library that pass alone ran for 3+ hours and
// was still incomplete. The count is now an estimate rather than exact (it
// can overstate on a library with a configured version-retention limit —
// see listItems.ts's VERSION_LABEL_FIELD comment for the full explanation)
// in exchange for costing nothing at all — EXCEPT on a list where
// SMTotalFileStreamSize isn't selectable at all (Storage Metrics not active
// for that list/site), where listItems.ts falls back to the old per-file
// request when includeVersionHistory is on, paying that same per-file cost
// again but only for libraries that actually need it.
export async function walkLibrary(
  client: SpApiClient,
  siteUrl: string,
  library: LibraryInfo,
  options: ScanOptions,
  users: SiteUsers,
  onEntry: (entry: FileEntry) => void,
  onProgress?: (fetchedSoFar: number) => void,
): Promise<WalkLibraryResult> {
  let items: FlatItem[];
  let skippedVersions = 0;
  try {
    items = library.isRecycleBin
      ? await fetchRecycleBinItems(client, siteUrl, { signal: options.signal, onProgress })
      : await fetchLibraryItems(client, siteUrl, library, {
        signal: options.signal,
        onProgress,
        includeVersionHistory: options.includeVersionHistory,
        onVersionSkipped: () => { skippedVersions++; },
      });
  } catch (err: any) {
    // Terminal for this library. Reported rather than silently under-counted
    // — a scan that quietly omits a whole library is worse than one that
    // says which library it could not read.
    const message = err?.message ?? String(err);
    // eslint-disable-next-line no-console
    console.warn(`[SmartStorageAnalyzer] Could not enumerate ${library.title}: ${message}`);
    return { failed: { url: library.serverRelativeUrl, error: message }, skippedVersions };
  }

  if (options.signal?.aborted) return { skippedVersions };

  const files = items.filter((i) => !i.isFolder);
  const entries: FileEntry[] = files.map((f) => {
    const ageDays = ageInDays(f.modified);
    // Recycle Bin rows arrive with the name/login already resolved
    // (DeletedByTitle/Email are plain strings, not lookup ids) — prefer
    // those directly over an authorId lookup that was never set for them.
    const author = f.authorDisplayName != null || f.authorLoginName != null
      ? { title: f.authorDisplayName, loginName: f.authorLoginName }
      : (f.authorId != null ? users.get(f.authorId) : undefined);
    return {
      name: f.name,
      serverRelativeUrl: f.fileRef,
      libraryTitle: library.title,
      sizeBytes: f.sizeBytes,
      timeCreated: f.created,
      timeLastModified: f.modified,
      authorLoginName: author?.loginName,
      authorDisplayName: author?.title,
      ageDays,
      tier: classify(ageDays, options.staleDays, options.veryStaleDays),
      // Both free from the bulk sweep — see the module comment above.
      versionSizeBytes: f.versionSizeBytes,
      versionCount: f.versionCountApprox,
    };
  });

  for (const entry of entries) onEntry(entry);
  return { skippedVersions };
}
