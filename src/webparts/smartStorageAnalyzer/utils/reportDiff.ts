import { CandidateTier, ReportDiff, StoredReport } from '../models/models';

// Compares two saved Storage Report scans (same site, different points in
// time) so an owner can see whether stale storage is growing or being
// cleaned up between scans.
export function diffReports(older: StoredReport, newer: StoredReport): ReportDiff {
  const olderStalePaths = new Set(
    older.entries
      .filter((e) => e.tier === CandidateTier.Stale || e.tier === CandidateTier.VeryStale)
      .map((e) => e.serverRelativeUrl),
  );
  const newerStalePaths = new Set(
    newer.entries
      .filter((e) => e.tier === CandidateTier.Stale || e.tier === CandidateTier.VeryStale)
      .map((e) => e.serverRelativeUrl),
  );

  let newStaleCount = 0;
  newerStalePaths.forEach((p) => {
    if (!olderStalePaths.has(p)) newStaleCount++;
  });
  let resolvedStaleCount = 0;
  olderStalePaths.forEach((p) => {
    if (!newerStalePaths.has(p)) resolvedStaleCount++;
  });

  return {
    olderTimestamp: older.timestamp,
    newerTimestamp: newer.timestamp,
    sizeDeltaBytes: newer.summary.totalSizeBytes - older.summary.totalSizeBytes,
    newStaleCount,
    resolvedStaleCount,
    totalFilesDelta: newer.summary.totalFiles - older.summary.totalFiles,
  };
}
