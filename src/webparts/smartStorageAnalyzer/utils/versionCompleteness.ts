import { StorageReportSummary } from '../models/models';

// True when a scan's Version History Size (and therefore Total Storage Size)
// is a floor rather than an exact figure: some lookups failed, some files were
// never measured, or at least one library had no usable version source.
export function isVersionHistoryIncomplete(summary: StorageReportSummary): boolean {
  if (!summary.versionHistoryIncluded) return false;
  return (summary.skippedVersions ?? 0) > 0
    || (summary.unmeasuredVersions ?? 0) > 0
    || usedVersionStrategy(summary, 'none');
}

// versionSizeStrategy is every library's strategy joined with ", " (e.g.
// "inline, none"), so an exact comparison misses any scan that mixed them.
export function usedVersionStrategy(summary: StorageReportSummary, kind: string): boolean {
  return (summary.versionSizeStrategy ?? '').split(', ').indexOf(kind) !== -1;
}
