import { SpApiClient } from './spCore';
import { getLibraries, getSubwebsRecursive } from './siteDiscovery';
import { walkLibrary } from './fileWalk';
import { fetchSiteUsers } from './listItems';
import { CandidateTier, FileEntry, ScanOptions, ScanProgress, StorageReportSummary } from '../../models/models';

export interface ScanResult {
  entries: FileEntry[];
  summary: StorageReportSummary;
  canceled: boolean;
}

// Caps how many skipped-folder error details get stored in the report
// (IndexedDB history) — skippedFolders count itself stays uncapped/exact,
// this only limits the detail list attached for display.
const MAX_SKIPPED_DETAILS = 200;

// Orchestrates a full Storage Report scan: resolves the site + (optionally)
// every subsite beneath it, walks every document library on each, and
// classifies every file. Libraries are processed one at a time — each
// library's own walkLibrary() already runs with bounded concurrency
// internally (options.scanConcurrency), so nesting another concurrent pool
// here would multiply total in-flight requests instead of bounding them.
export async function scanSite(
  client: SpApiClient,
  options: ScanOptions,
  onProgress: (progress: ScanProgress) => void,
  onEntry?: (entry: FileEntry) => void,
): Promise<ScanResult> {
  const start = Date.now();

  const entries: FileEntry[] = [];
  let libsDone = 0;
  let libsTotal = 0;
  let skippedFolders = 0;
  let skippedSites = 0;
  let skippedVersions = 0;
  let unmeasuredVersions = 0;
  // Cumulative across COMPLETED libraries; the in-progress library's own
  // counters are added on top when emitting — see `emit`.
  let itemsDoneAcrossLibraries = 0;
  let filesDoneAcrossLibraries = 0;
  let libraryItemsRead = 0;
  let libraryFilesSeen = 0;
  let currentSiteUrl = options.siteUrl;
  let currentLibraryTitle: string | undefined;
  let multiSite = false;
  let totalItemsHint = 0;

  // The ONE place a ScanProgress is built.
  //
  // There used to be four object literals across this function, and they had
  // already drifted: the pre-library emit omitted `phase` entirely (so the view
  // read a library's first update as an items-phase update), and every one of
  // them sent `scanned: entries.length` — which is 0 for the whole of the first
  // library, because entries are published only after the version escalation.
  // A single builder means a new stage cannot forget a field.
  const emit = (
    p: Pick<ScanProgress, 'stage' | 'stageLabel'>
      & Partial<ScanProgress>
      & { stageKeySuffix?: string },
  ): void => {
    const { stageKeySuffix, ...rest } = p;
    onProgress({
      libsDone,
      libsTotal,
      itemsFetched: itemsDoneAcrossLibraries + libraryItemsRead,
      totalItemsHint: totalItemsHint > 0 ? totalItemsHint : undefined,
      filesSeen: filesDoneAcrossLibraries + libraryFilesSeen,
      scanned: entries.length,
      skippedVersions,
      unmeasuredVersions,
      skippedLibraries: skippedFolders,
      libraryTitle: currentLibraryTitle,
      // Only when the scan really spans several sites, so a single-site scan
      // doesn't carry a redundant segment on screen.
      siteLabel: multiSite ? currentSiteUrl : undefined,
      // Site included so two libraries with the same title on different subsites
      // don't share a rate window.
      stageKey: `${currentSiteUrl}|${currentLibraryTitle ?? ''}|${stageKeySuffix ?? p.stage}`,
      ...rest,
    });
  };

  // Discovery is a multi-request window (recursive subweb enumeration, then one
  // getLibraries per site) that previously reported nothing at all — the UI sat
  // on "Starting scan…" with libsTotal 0 for its whole duration.
  emit({ stage: 'discovering', stageLabel: 'Finding libraries', detail: 'reading site structure' });

  const sites = [options.siteUrl];
  if (options.includeSubsites) {
    const subwebs = await getSubwebsRecursive(client, options.siteUrl);
    sites.push(...subwebs.map((w) => w.url));
  }
  multiSite = sites.length > 1;
  // Which escalation mechanism was used for version size. Recorded from the
  // libraries that needed one; if several disagree (rare — it depends on the
  // list, not the tenant) the least capable wins, since that is what actually
  // bounds how complete the report's version numbers are.
  const versionStrategies = new Set<string>();
  const skippedFolderDetails: { url: string; error: string }[] = [];
  const librariesPerSite: { siteUrl: string; libraries: Awaited<ReturnType<typeof getLibraries>> }[] = [];

  for (let i = 0; i < sites.length; i++) {
    const siteUrl = sites[i];
    currentSiteUrl = siteUrl;
    emit({
      stage: 'discovering',
      stageLabel: 'Finding libraries',
      detail: `site ${i + 1} of ${sites.length}`,
      stageDone: i,
      stageTotal: sites.length,
    });
    try {
      const libs = await getLibraries(client, siteUrl, options.includeHidden);
      const filtered = options.libraryUrls?.length
        ? libs.filter((l) => options.libraryUrls!.indexOf(l.serverRelativeUrl) !== -1)
        : libs;
      librariesPerSite.push({ siteUrl, libraries: filtered });
      libsTotal += filtered.length;
    } catch (err) {
      // The requested site itself (index 0) failing means there is nothing
      // to report — propagate so the caller shows "Scan failed" instead of
      // a deceptively "successful" empty report. Subsites are best-effort:
      // no access to a subsite's lists is expected and common (mirrors how
      // subsite discovery itself skips inaccessible branches), so that
      // stays a silent, counted skip rather than aborting the whole scan.
      if (i === 0) throw err;
      skippedSites++;
    }
  }

  // Sum of every library's own ItemCount (files + folders), known once discovery
  // finishes. A COUNTER ONLY — ItemCount is routinely stale, and using it as a
  // denominator is what let the bar peg at 100% and the ETA read "~0s remaining"
  // while a third of the work remained. Per-library page counts drive the bar
  // now (see listItems.ts).
  totalItemsHint = librariesPerSite.reduce(
    (sum, { libraries }) => sum + libraries.reduce((s, l) => s + (l.itemCount ?? 0), 0),
    0,
  );

  for (const { siteUrl, libraries } of librariesPerSite) {
    if (options.signal?.aborted) break;
    currentSiteUrl = siteUrl;
    // Author/Editor names for every library on this site, resolved once. The
    // bulk item sweep returns numeric lookup ids rather than expanded user
    // objects, because $expand=Author on a per-file query was one of the
    // biggest per-row costs in the old walk.
    const users = await fetchSiteUsers(client, siteUrl, options.signal);
    for (const library of libraries) {
      if (options.signal?.aborted) break;
      currentLibraryTitle = library.title;
      libraryItemsRead = 0;
      libraryFilesSeen = 0;
      emit({ stage: 'items', stageLabel: 'Reading items', detail: 'starting', stageKeySuffix: 'items' });
      const {
        failed,
        skippedVersions: libSkippedVersions,
        unmeasuredVersions: libUnmeasuredVersions,
        versionStrategy,
      } = await walkLibrary(
        client, siteUrl, library, options, users,
        (entry) => {
          entries.push(entry);
          onEntry?.(entry);
        },
        // Every stage of this library's scan arrives here — items sweep, version
        // probe, bulk version sweep, validation, per-file measurement — already
        // carrying its own unit and denominator. It only needs the library's
        // running counters lifted into the scan-wide ones.
        (p) => {
          libraryItemsRead = p.itemsRead;
          libraryFilesSeen = p.filesSeen;
          emit({
            stage: p.stage,
            stageLabel: p.stageLabel,
            detail: p.detail,
            stageDone: p.done,
            stageTotal: p.total,
            stageUnit: p.unit,
            stageKeySuffix: p.stageKeySuffix,
          });
        },
      );
      // A library that could not be enumerated at all. Counted with the same
      // skipped/details reporting the UI already surfaces, so a partial scan
      // still says exactly what is missing.
      if (failed) {
        skippedFolders++;
        if (skippedFolderDetails.length < MAX_SKIPPED_DETAILS) {
          skippedFolderDetails.push(failed);
        }
      }
      skippedVersions += libSkippedVersions;
      unmeasuredVersions += libUnmeasuredVersions;
      if (versionStrategy) versionStrategies.add(versionStrategy);
      itemsDoneAcrossLibraries += libraryItemsRead;
      filesDoneAcrossLibraries += libraryFilesSeen;
      libraryItemsRead = 0;
      libraryFilesSeen = 0;
      libsDone++;
      emit({ stage: 'items', stageLabel: 'Finished library', stageKeySuffix: 'done' });
    }
  }

  const summary: StorageReportSummary = entries.reduce(
    (acc, e) => {
      acc.totalSizeBytes += e.sizeBytes;
      acc.totalVersionSizeBytes! += e.versionSizeBytes ?? 0;
      acc.totalVersionCount! += e.versionCount ?? 0;
      acc.totalFiles += 1;
      if (e.tier === CandidateTier.Stale) {
        acc.staleCount += 1;
        acc.staleSizeBytes += e.sizeBytes;
      } else if (e.tier === CandidateTier.VeryStale) {
        acc.veryStaleCount += 1;
        acc.veryStaleSizeBytes += e.sizeBytes;
      }
      return acc;
    },
    {
      totalSizeBytes: 0,
      totalFiles: 0,
      staleCount: 0,
      veryStaleCount: 0,
      staleSizeBytes: 0,
      veryStaleSizeBytes: 0,
      durationSeconds: 0,
      skippedFolders: 0,
      skippedSites: 0,
      totalVersionSizeBytes: 0,
      totalVersionCount: 0,
    } as StorageReportSummary,
  );
  summary.durationSeconds = (Date.now() - start) / 1000;
  summary.skippedFolders = skippedFolders;
  summary.skippedSites = skippedSites;
  summary.skippedFolderDetails = skippedFolderDetails;
  summary.skippedVersions = skippedVersions;
  summary.unmeasuredVersions = unmeasuredVersions;
  summary.versionHistoryIncluded = options.includeVersionHistory;
  // Undefined when every library got version size inline — the common, fast
  // case, where naming a "strategy" would imply something unusual happened.
  if (versionStrategies.size > 0) {
    summary.versionSizeStrategy = Array.from(versionStrategies).sort().join(', ');
  }

  return { entries, summary, canceled: !!options.signal?.aborted };
}
