import { SpApiClient } from './spCore';
import { FileEntry, LibraryInfo, ScanOptions } from '../../models/models';
import { ageInDays, classify } from '../../utils/archivalClassification';
import { fetchLibraryItems, FlatItem } from './listItems';
import { fetchRecycleBinItems } from './recycleBin';
import { applyVersionSizes, probeVersionSizeStrategy } from './versionSizes';

export interface WalkLibraryResult {
  // Set when the library's own bulk query failed and nothing could be read
  // from it. There is deliberately no recursive-walk fallback any more, so
  // this is terminal for the library and the caller reports it as skipped.
  failed?: { url: string; error: string };
  // Version measurements ATTEMPTED AND FAILED — only ever nonzero when the
  // escalation reached its per-file path for this library.
  skippedVersions: number;
  // Version measurements NEVER ATTEMPTED (Quick-mode budget, cancellation, or
  // no working mechanism at all). See StorageReportSummary for why this is
  // deliberately not folded into skippedVersions.
  unmeasuredVersions: number;
  // Which mechanism produced this library's version numbers, when it needed
  // an escalation at all. Undefined means the inline bulk field worked.
  versionStrategy?: string;
}

export type SiteUsers = Map<number, { title: string; loginName: string }>;

// Progress for the version-measuring phase of one library, if it reaches the
// per-file path. Already throttled at the source (versionSizes.ts).
export type VersionProgress = (done: number, total: number) => void;

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
// in exchange for costing nothing at all.
//
// When that inline field is rejected by the list — which happens, and on a
// real tenant happens for a 193,915-item "Documents" library — version size
// is escalated through versionSizes.ts instead. That escalation is owned HERE
// rather than inside listItems.ts purely to keep the two modules from
// depending on each other's values (versionSizes.ts needs FlatItem).
export async function walkLibrary(
  client: SpApiClient,
  siteUrl: string,
  library: LibraryInfo,
  options: ScanOptions,
  users: SiteUsers,
  onEntry: (entry: FileEntry) => void,
  onProgress?: (fetchedSoFar: number) => void,
  onVersionProgress?: VersionProgress,
): Promise<WalkLibraryResult> {
  let items: FlatItem[];
  let skippedVersions = 0;
  let unmeasuredVersions = 0;
  let versionStrategy: string | undefined;
  // Set by listItems.ts when the reduced field set won, i.e. these rows carry
  // no inline version size.
  let versionFieldUnavailable = false;
  try {
    items = library.isRecycleBin
      ? await fetchRecycleBinItems(client, siteUrl, { signal: options.signal, onProgress })
      : await fetchLibraryItems(client, siteUrl, library, {
        signal: options.signal,
        onProgress,
        onVersionFieldUnavailable: () => { versionFieldUnavailable = true; },
      });
  } catch (err: any) {
    // Terminal for this library. Reported rather than silently under-counted
    // — a scan that quietly omits a whole library is worse than one that
    // says which library it could not read.
    const message = err?.message ?? String(err);
    // eslint-disable-next-line no-console
    console.warn(`[SmartStorageAnalyzer] Could not enumerate ${library.title}: ${message}`);
    return {
      failed: { url: library.serverRelativeUrl, error: message },
      skippedVersions,
      unmeasuredVersions,
    };
  }

  if (options.signal?.aborted) return { skippedVersions, unmeasuredVersions };

  // ── Version-history escalation ──────────────────────────────────────────
  // Only when the inline field was missing AND the user asked for version
  // history. The Recycle Bin is excluded: its rows are not list items and
  // deleted files have no addressable /Versions collection.
  if (versionFieldUnavailable && !library.isRecycleBin) {
    if (!options.includeVersionHistory) {
      // Same message this has always emitted — the data genuinely isn't in
      // the report, and saying so beats a silent 0 B.
      // eslint-disable-next-line no-console
      console.warn(
        `[SmartStorageAnalyzer] ${library.title}: version-history size unavailable on this list ` +
        '— sizes exclude retained versions for this library.',
      );
    } else {
      const strategy = await probeVersionSizeStrategy(client, siteUrl, library, options.signal);
      const fill = await applyVersionSizes(client, siteUrl, library, items, {
        signal: options.signal,
        mode: options.versionScanMode ?? 'quick',
        onProgress: onVersionProgress,
        onSkipped: () => { skippedVersions++; },
      }, strategy);
      versionStrategy = fill.strategy;
      // skippedVersions is already accumulated via onSkipped, so only the
      // never-attempted count is taken from the result here.
      unmeasuredVersions += fill.unmeasured;
    }
  }

  // Entries are built AFTER the escalation on purpose: applyVersionSizes
  // mutates the FlatItems in place, and onEntry must never publish a
  // FileEntry whose versionSizeBytes is about to change underneath it.

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
      // Usually free from the bulk sweep; otherwise filled by the escalation
      // above. Still undefined when nothing could measure it — which the
      // report renders as "—" rather than 0 B.
      versionSizeBytes: f.versionSizeBytes,
      versionCount: f.versionCountApprox,
    };
  });

  for (const entry of entries) onEntry(entry);
  return { skippedVersions, unmeasuredVersions, versionStrategy };
}
