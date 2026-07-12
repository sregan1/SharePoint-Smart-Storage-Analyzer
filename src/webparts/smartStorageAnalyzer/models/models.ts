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
}

export interface FolderStorageNode {
  name: string;
  serverRelativeUrl: string;
  totalSizeBytes: number;
  fileCount: number;
  lastModified?: string;
  sizeSource: 'storageMetrics' | 'estimate';
  children: FolderStorageNode[];
  hasChildren: boolean;
  isLoading?: boolean;
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
}

export interface ScanOptions {
  siteUrl: string;
  includeSubsites: boolean;
  includeHidden: boolean;
  libraryUrls?: string[];
  staleDays: number;
  veryStaleDays: number;
  scanConcurrency: number;
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
}

export interface StoredReport {
  id: string;
  timestamp: number;
  siteUrl: string;
  options: Pick<ScanOptions, 'includeSubsites' | 'staleDays' | 'veryStaleDays'>;
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
}
