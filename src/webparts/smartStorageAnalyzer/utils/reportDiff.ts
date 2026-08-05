import { CandidateTier, FileEntry, ReportDiff, StoredReportMeta } from '../models/models';

// Compares two saved Storage Report scans (same site, different points in
// time) so an owner can see whether stale storage is growing or being
// cleaned up between scans.
//
// Split into two tiers of answer, deliberately:
//   - Size and file-count deltas come from the summaries, which are always
//     stored, so those always work.
//   - "New" and "resolved" archival candidates need to compare individual FILE
//     PATHS, which means the file listings. Those are stored separately and can
//     be evicted when browser storage runs short, so they are passed in as
//     nullable and the counts are omitted rather than guessed when either side
//     is missing.
//
// Reporting 0 new / 0 resolved for a report whose listing was evicted would be
// indistinguishable from "nothing changed", which is why these are optional
// rather than defaulted.
export function diffReports(
  older: StoredReportMeta,
  newer: StoredReportMeta,
  olderEntries: FileEntry[] | null,
  newerEntries: FileEntry[] | null,
): ReportDiff {
  const base: ReportDiff = {
    olderTimestamp: older.timestamp,
    newerTimestamp: newer.timestamp,
    sizeDeltaBytes: newer.summary.totalSizeBytes - older.summary.totalSizeBytes,
    totalFilesDelta: newer.summary.totalFiles - older.summary.totalFiles,
  };
  if (!olderEntries || !newerEntries) return base;

  const stalePaths = (entries: FileEntry[]): Set<string> => new Set(
    entries
      .filter((e) => e.tier === CandidateTier.Stale || e.tier === CandidateTier.VeryStale)
      .map((e) => e.serverRelativeUrl),
  );
  const olderStalePaths = stalePaths(olderEntries);
  const newerStalePaths = stalePaths(newerEntries);

  let newStaleCount = 0;
  newerStalePaths.forEach((p) => {
    if (!olderStalePaths.has(p)) newStaleCount++;
  });
  let resolvedStaleCount = 0;
  olderStalePaths.forEach((p) => {
    if (!newerStalePaths.has(p)) resolvedStaleCount++;
  });

  return { ...base, newStaleCount, resolvedStaleCount };
}
