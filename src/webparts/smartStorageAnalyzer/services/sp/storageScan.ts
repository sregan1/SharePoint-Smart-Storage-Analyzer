import { SpApiClient } from './spCore';
import { getLibraries, getSubwebsRecursive } from './siteDiscovery';
import { walkLibrary } from './fileWalk';
import { CandidateTier, FileEntry, ScanOptions, ScanProgress, StorageReportSummary } from '../../models/models';

export interface ScanResult {
  entries: FileEntry[];
  summary: StorageReportSummary;
}

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
  const librariesPerSite: { siteUrl: string; libraries: Awaited<ReturnType<typeof getLibraries>> }[] = [];

  for (const siteUrl of sites) {
    try {
      const libs = await getLibraries(client, siteUrl, options.includeHidden);
      const filtered = options.libraryUrls?.length
        ? libs.filter((l) => options.libraryUrls!.indexOf(l.serverRelativeUrl) !== -1)
        : libs;
      librariesPerSite.push({ siteUrl, libraries: filtered });
      libsTotal += filtered.length;
    } catch {
      // no access to this site's lists — skip silently, same as subsite discovery
    }
  }

  for (const { siteUrl, libraries } of librariesPerSite) {
    for (const library of libraries) {
      onProgress({ message: `Scanning ${library.title}…`, scanned: entries.length, libsDone, libsTotal });
      await walkLibrary(client, siteUrl, library, options, (entry) => {
        entries.push(entry);
        onEntry?.(entry);
      });
      libsDone++;
      onProgress({ message: `Scanning ${library.title}…`, scanned: entries.length, libsDone, libsTotal });
    }
  }

  const summary: StorageReportSummary = entries.reduce(
    (acc, e) => {
      acc.totalSizeBytes += e.sizeBytes;
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
    } as StorageReportSummary,
  );
  summary.durationSeconds = (Date.now() - start) / 1000;

  return { entries, summary };
}
