// Shared data model for the Smart Storage Analyzer web part.

export enum CandidateTier {
  Active = 'Active',
  Stale = 'Stale',
  VeryStale = 'VeryStale',
}

export interface LibraryInfo {
  title: string;
  serverRelativeUrl: string;
  itemCount?: number;
  totalSizeBytes?: number;
  sizeSource?: 'storageMetrics' | 'estimate' | 'unknown';
  lastModified?: string;
  noCrawl?: boolean;
  baseTemplate: number;
  versionSizeBytes?: number;
}

export interface FolderStorageNode {
  name: string;
  serverRelativeUrl: string;
  totalSizeBytes: number;
  fileCount: number;
  lastModified?: string;
  // 'error' means the recursive fallback walk couldn't complete for this
  // folder (throttling exhausted, a transient error, a permissions issue) —
  // totalSizeBytes/fileCount are whatever partial total it had when it gave
  // up, usually 0, and must NOT be presented as a confirmed empty folder.
  sizeSource: 'storageMetrics' | 'estimate' | 'error';
  children: FolderStorageNode[];
  hasChildren: boolean;
  isLoading?: boolean;
  versionSizeBytes?: number;
}

export interface FileEntry {
  name: string;
  serverRelativeUrl: string;
  libraryTitle: string;
  sizeBytes: number;
  timeCreated: string;
  timeLastModified: string;
  authorLoginName?: string;
  authorDisplayName?: string;
  ageDays: number;
  tier: CandidateTier;
  // Sum of SP.FileVersion.Size across a file's retained old versions. Only
  // populated when ScanOptions.includeVersionHistory is true (opt-in — an
  // extra REST call per file). Undefined means "not measured", not zero.
  versionSizeBytes?: number;
}

export interface ScanOptions {
  siteUrl: string;
  includeSubsites: boolean;
  includeHidden: boolean;
  libraryUrls?: string[];
  staleDays: number;
  veryStaleDays: number;
  scanConcurrency: number;
  // Opt-in: fetches each file's version history (SP.FileVersion.Size) and
  // sums it separately from the current file size. Requires an extra REST
  // call per file, so it roughly doubles request volume and scan time.
  includeVersionHistory: boolean;
  // Cooperative cancellation only: checked between queued tasks/libraries/
  // sites, not passed into in-flight HTTP requests. A scan can be many
  // minutes long with no other way to stop it short of leaving the page.
  signal?: AbortSignal;
}

export interface ScanProgress {
  message: string;
  scanned: number;
  libsDone: number;
  libsTotal: number;
}

export interface StorageReportSummary {
  totalSizeBytes: number;
  totalFiles: number;
  staleCount: number;
  veryStaleCount: number;
  staleSizeBytes: number;
  veryStaleSizeBytes: number;
  durationSeconds: number;
  // Folders/subsites that could not be read (403, throttling exhausted,
  // list view threshold, etc.) and were skipped rather than aborting the
  // whole scan — so totals above may be partial. Optional so reports saved
  // by older versions (IndexedDB history) still deserialize; treat a
  // missing value as 0.
  skippedFolders?: number;
  skippedSites?: number;
  // The actual failure behind each skipped folder (e.g. list view threshold,
  // 403, throttling exhausted after retries) — capped at
  // MAX_SKIPPED_DETAILS entries (see storageScan.ts) so a pathological scan
  // with thousands of failures doesn't bloat the stored report; skippedFolders
  // above is always the true, uncapped count.
  skippedFolderDetails?: { url: string; error: string }[];
  // Per-file version-history fetch failures (kept in-scope rather than
  // aborting the scan) — see skippedFolders/skippedSites above.
  skippedVersions?: number;
  // Sum of FileEntry.versionSizeBytes across all entries. Only meaningful
  // when versionHistoryIncluded is true — see that field's comment.
  totalVersionSizeBytes?: number;
  // Whether this particular scan collected version-history data. Needed
  // because totalVersionSizeBytes === 0 is ambiguous between "no old
  // versions exist" and "the toggle was off" (or this report predates the
  // feature and the field is simply missing).
  versionHistoryIncluded?: boolean;
}

export interface StoredReport {
  id: string;
  timestamp: number;
  siteUrl: string;
  options: Pick<ScanOptions, 'includeSubsites' | 'staleDays' | 'veryStaleDays' | 'includeVersionHistory'>;
  summary: StorageReportSummary;
  entries: FileEntry[];
  // Set when `entries` holds only stale/very-stale rows because the full
  // scan exceeded the stored-history row cap (see StorageReportView) — kept
  // small enough that 10 large scans don't strain IndexedDB quota. The
  // summary and diffing (diffReports only reads stale-tier paths + summary
  // fields) stay fully accurate either way; only the "View" file listing is
  // incomplete for a truncated report.
  entriesTruncated?: boolean;
}

export interface ReportDiff {
  olderTimestamp: number;
  newerTimestamp: number;
  sizeDeltaBytes: number;
  newStaleCount: number;
  resolvedStaleCount: number;
  totalFilesDelta: number;
}

// A single rectangle-ready node consumed by the Treemap renderer — the
// common shape both folder rollups and leaf files get mapped into.
export interface TreemapItem {
  id: string;
  label: string;
  sizeBytes: number;
  kind: 'folder' | 'file' | 'other';
  tier?: CandidateTier;
  lastModified?: string;
  count?: number; // populated only for the aggregated "Other" bucket
  itemCount?: number; // folders only — rolled-up file count (FolderStorageNode.fileCount)
  versionSizeBytes?: number;
  // Folders only — true when FolderStorageNode.sizeSource === 'error': the
  // size/count shown is not a confirmed result, just whatever partial total
  // a throttled/failed live walk had when it gave up.
  sizeUnknown?: boolean;
}

export interface TreemapRect extends TreemapItem {
  x: number;
  y: number;
  width: number;
  height: number;
}

// Unified row shape for the Explorer's List View — one immediate child of the
// current folder, whether it's a subfolder rollup or a leaf file.
export interface FolderListRow {
  kind: 'folder' | 'file';
  name: string;
  serverRelativeUrl: string;
  sizeBytes: number;
  lastModified?: string;      // folder: FolderStorageNode.lastModified; file: timeLastModified
  itemCount?: number;         // folders only — FolderStorageNode.fileCount (rolled-up count)
  ageDays?: number;           // files only
  tier?: CandidateTier;       // files only
  authorDisplayName?: string; // files only
  versionSizeBytes?: number;  // files only — folders show '—' (no recursive rollup)
  sizeUnknown?: boolean;      // folders only — see TreemapItem.sizeUnknown
}
