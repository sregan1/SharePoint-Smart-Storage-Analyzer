import { SpApiClient } from './spCore';
import { getLibraries, getSubwebsRecursive } from './siteDiscovery';
import { walkLibrary } from './fileWalk';
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
  const sites = [options.siteUrl];
  if (options.includeSubsites) {
    const subwebs = await getSubwebsRecursive(client, options.siteUrl);
    sites.push(...subwebs.map((w) => w.url));
  }

  const entries: FileEntry[] = [];
  let libsDone = 0;
  let libsTotal = 0;
  let skippedFolders = 0;
  let skippedSites = 0;
  let skippedVersions = 0;
  const skippedFolderDetails: { url: string; error: string }[] = [];
  const librariesPerSite: { siteUrl: string; libraries: Awaited<ReturnType<typeof getLibraries>> }[] = [];

  for (let i = 0; i < sites.length; i++) {
    const siteUrl = sites[i];
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

  for (const { siteUrl, libraries } of librariesPerSite) {
    if (options.signal?.aborted) break;
    for (const library of libraries) {
      if (options.signal?.aborted) break;
      onProgress({ message: `Scanning ${library.title}…`, scanned: entries.length, libsDone, libsTotal });
      const { skippedFolders: libSkipped, skippedVersions: libSkippedVersions, skippedFolderDetails: libDetails } = await walkLibrary(
        client, siteUrl, library, options, (entry) => {
          entries.push(entry);
          onEntry?.(entry);
        },
      );
      skippedFolders += libSkipped;
      skippedVersions += libSkippedVersions;
      if (skippedFolderDetails.length < MAX_SKIPPED_DETAILS) {
        skippedFolderDetails.push(...libDetails.slice(0, MAX_SKIPPED_DETAILS - skippedFolderDetails.length));
      }
      libsDone++;
      onProgress({ message: `Scanning ${library.title}…`, scanned: entries.length, libsDone, libsTotal });
    }
  }

  const summary: StorageReportSummary = entries.reduce(
    (acc, e) => {
      acc.totalSizeBytes += e.sizeBytes;
      acc.totalVersionSizeBytes! += e.versionSizeBytes ?? 0;
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
    } as StorageReportSummary,
  );
  summary.durationSeconds = (Date.now() - start) / 1000;
  summary.skippedFolders = skippedFolders;
  summary.skippedSites = skippedSites;
  summary.skippedVersions = skippedVersions;
  summary.skippedFolderDetails = skippedFolderDetails;
  summary.versionHistoryIncluded = options.includeVersionHistory;

  return { entries, summary, canceled: !!options.signal?.aborted };
}
