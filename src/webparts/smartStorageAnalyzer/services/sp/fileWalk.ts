import { SpApiClient } from './spCore';
import { FileEntry, LibraryInfo, ScanOptions, ScanStage } from '../../models/models';
import { ageInDays, classify } from '../../utils/archivalClassification';
import { fetchLibraryItems, FlatItem } from './listItems';
import { fetchRecycleBinItems } from './recycleBin';
import { applyVersionSizes, probeVersionSizeStrategy, VersionProgressEvent } from './versionSizes';
import { getStorageMetrics } from './storageMetrics';

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

// An INDEPENDENT check on the version-history total, for one extra request.
//
// The folder-level StorageMetrics endpoint reports TotalSize (version-inclusive)
// and TotalFileStreamSize (current content only) for a library's root folder, so
// their difference is a library-level version total derived from a completely
// different SharePoint subsystem than the per-file /Versions sums. If the two
// disagree wildly, one of them is wrong and the report shouldn't be trusted
// without a look — which is worth knowing when the figure is large enough to
// drive a cleanup decision.
//
// Diagnostic only, deliberately: it logs and never alters the reported numbers.
// The per-file measurement is the authoritative source (it is what SharePoint
// itself charges for), and this has no per-file granularity, so it can
// corroborate but not correct.
async function crossCheckVersionTotal(
  client: SpApiClient,
  siteUrl: string,
  library: LibraryInfo,
  items: FlatItem[],
  signal: AbortSignal | undefined,
): Promise<void> {
  if (signal?.aborted) return;
  try {
    const metrics = await getStorageMetrics(client, siteUrl, library.serverRelativeUrl);
    if (!metrics || metrics.totalFileStreamSizeBytes == null) return;
    // A stale zero is a documented, common state for this endpoint (see
    // libraryStats.ts) — it means "not computed yet", not "empty".
    if (metrics.totalSizeBytes <= 0) return;

    const fromMetrics = metrics.totalSizeBytes - metrics.totalFileStreamSizeBytes;
    const measuredVersions = items.reduce((sum, i) => sum + (i.isFolder ? 0 : (i.versionSizeBytes ?? 0)), 0);
    const measuredCurrent = items.reduce((sum, i) => sum + (i.isFolder ? 0 : i.sizeBytes), 0);
    if (fromMetrics <= 0 && measuredVersions <= 0) return;

    const bigger = Math.max(fromMetrics, measuredVersions);
    const driftPct = bigger > 0 ? Math.abs(fromMetrics - measuredVersions) / bigger * 100 : 0;
    const fmt = (n: number): string => `${(n / 1024 / 1024 / 1024).toFixed(1)} GB`;
    // Every raw input, not just the derived gap.
    //
    // The first run of this check reported "27% apart" and that was not enough to
    // act on: it could not distinguish "our per-file sum is wrong" from
    // "TotalSize is lagging", because the two operands were already collapsed
    // into one number. Printing all four lets the arithmetic be checked by hand:
    // TotalFileStreamSize should land on our own current-content sum (which
    // corroborates the field's meaning), and TotalSize should land on
    // current + versions. Whichever of those two fails is the one to distrust.
    const detail =
      `StorageMetrics TotalSize ${fmt(metrics.totalSizeBytes)}, `
      + `TotalFileStreamSize ${fmt(metrics.totalFileStreamSizeBytes)} `
      + `⇒ versions ${fmt(fromMetrics)}. `
      + `Measured: current ${fmt(measuredCurrent)}, versions ${fmt(measuredVersions)} `
      + `⇒ expected TotalSize ${fmt(measuredCurrent + measuredVersions)}.`;

    // 25%: loose on purpose. Both sides are legitimately imprecise — the metrics
    // job lags, and TotalSize may include metadata overhead and possibly the
    // first-stage recycle bin. This is looking for an order-of-magnitude
    // disagreement, not an accounting reconciliation.
    if (driftPct > 25) {
      // eslint-disable-next-line no-console
      console.warn(
        `[SmartStorageAnalyzer] ${library.title}: version-history cross-check DISAGREES ` +
        `(${driftPct.toFixed(0)}% apart). ${detail} The per-file figure is the authoritative one ` +
        '(/Versions is what SharePoint charges for) and StorageMetrics is maintained by a lagging ' +
        'background job, so a low TotalSize is the more likely explanation — but verify against ' +
        'Site Settings → Storage Metrics before acting on the number.',
      );
    } else {
      // eslint-disable-next-line no-console
      console.info(
        `[SmartStorageAnalyzer] ${library.title}: version-history cross-check agrees ` +
        `(${driftPct.toFixed(0)}% apart). ${detail}`,
      );
    }
  } catch {
    // Best-effort corroboration. Never let it affect a scan that otherwise
    // succeeded.
  }
}

// One progress update from anywhere inside a library's scan.
//
// Replaces two separate callbacks with two different units (items read, files
// measured). That split is why the version stages had nowhere to report from:
// there was no callback carrying a stage, so probing, bulk sweeping and
// validating — up to ~37 heavyweight requests — were structurally invisible.
export interface WalkStageProgress {
  stage: Exclude<ScanStage, 'discovering'>;
  stageLabel: string;
  detail?: string;
  done?: number;
  total?: number;
  unit?: 'items' | 'pages' | 'files';
  // Appended to the scan-level stage key so the view restarts its rate window
  // when this library moves to a genuinely different kind of work.
  stageKeySuffix: string;
  // Cumulative WITHIN THIS LIBRARY, carried on every emit from every stage so a
  // stage transition never momentarily zeroes a counter on screen.
  itemsRead: number;
  filesSeen: number;
  skippedVersions: number;
  unmeasuredVersions: number;
}
export type WalkProgress = (p: WalkStageProgress) => void;

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
  onProgress?: WalkProgress,
): Promise<WalkLibraryResult> {
  let items: FlatItem[];
  let skippedVersions = 0;
  let unmeasuredVersions = 0;
  let versionStrategy: string | undefined;
  // Set by listItems.ts when the reduced field set won, i.e. these rows carry
  // no inline version size.
  let versionFieldUnavailable = false;
  // Running library-level counters, so every emit from every stage carries them.
  let itemsRead = 0;
  let filesSeen = 0;
  const emit = (p: Omit<WalkStageProgress, 'itemsRead' | 'filesSeen' | 'skippedVersions' | 'unmeasuredVersions'>): void =>
    onProgress?.({ ...p, itemsRead, filesSeen, skippedVersions, unmeasuredVersions });

  try {
    items = library.isRecycleBin
      // Recycle Bin rows are all files and have no page denominator of their
      // own, so its sweep reports items only.
      ? await fetchRecycleBinItems(client, siteUrl, {
        signal: options.signal,
        onProgress: (fetchedSoFar) => {
          itemsRead = fetchedSoFar;
          filesSeen = fetchedSoFar;
          emit({
            stage: 'items',
            stageLabel: 'Reading deleted items',
            detail: `${fetchedSoFar.toLocaleString()} items read`,
            stageKeySuffix: 'items',
          });
        },
      })
      : await fetchLibraryItems(client, siteUrl, library, {
        signal: options.signal,
        onProgress: (p) => {
          itemsRead = p.items;
          filesSeen = p.files;
          emit({
            stage: 'items',
            stageLabel: 'Reading items',
            detail: p.pagesTotal
              ? `page ${p.pagesDone} of ${p.pagesTotal} · ${p.items.toLocaleString()} items`
              : `${p.items.toLocaleString()} items read`,
            done: p.pagesDone,
            total: p.pagesTotal,
            unit: 'pages',
            stageKeySuffix: 'items',
          });
        },
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
      // Every version stage funnels its own progress through here, translated
      // into the scan's vocabulary. The stage names map 1:1, so a new version
      // stage cannot appear without a home in the status line.
      const versionOptions = {
        signal: options.signal,
        onProgress: (e: VersionProgressEvent) => {
          const stage: Exclude<ScanStage, 'discovering'> = e.stage === 'probe'
            ? 'version-probe'
            : e.stage === 'bulk-sweep'
              ? 'version-bulk'
              : e.stage === 'validate' ? 'version-validate' : 'version-per-file';
          emit({
            stage,
            stageLabel: e.stage === 'per-file'
              ? 'Measuring version history'
              : e.stage === 'bulk-sweep'
                ? 'Reading version history'
                : e.stage === 'validate' ? 'Checking version history' : 'Checking how version history can be read',
            detail: e.detail,
            done: e.done,
            total: e.total,
            unit: e.unit,
            // The target count is part of the key: a per-file pass restarting
            // for a different library is a different rate window, and so is a
            // bulk sweep giving way to a per-file fallback.
            stageKeySuffix: `${e.stage}:${e.total ?? 0}`,
          });
        },
        onSkipped: () => { skippedVersions++; },
      };
      const strategy = await probeVersionSizeStrategy(client, siteUrl, library, items, versionOptions);
      const fill = await applyVersionSizes(client, siteUrl, library, items, versionOptions, strategy);
      versionStrategy = fill.strategy;
      // skippedVersions is already accumulated via onSkipped, so only the
      // never-attempted count is taken from the result here.
      unmeasuredVersions += fill.unmeasured;
      await crossCheckVersionTotal(client, siteUrl, library, items, options.signal);
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
