// Shared data model for the Smart Storage Analyzer web part.

export enum CandidateTier {
  Active = 'Active',
  Stale = 'Stale',
  VeryStale = 'VeryStale',
}

// How hard to work for version-history size on a library where no BULK
// mechanism for it exists and the only option left is one request per file
// (see services/sp/versionSizes.ts). Both modes behave identically whenever a
// bulk mechanism does work — which is the common case — so this only bites on
// lists where SharePoint won't project SMTotalFileStreamSize at all.
//
// Declared here rather than in versionSizes.ts because listItems.ts,
// fileWalk.ts and the UI all need it, and routing the type through the module
// that also owns FlatItem would make those files circularly dependent.
export type VersionScanMode = 'quick' | 'full';

export interface LibraryInfo {
  title: string;
  // List GUID. Addressing the list by id rather than title for the bulk
  // item query (listItems.ts) avoids escaping problems with titles
  // containing quotes/ampersands, and survives a library being renamed
  // mid-scan.
  id?: string;
  serverRelativeUrl: string;
  // True for the synthetic "Recycle Bin" pseudo-library (recycleBin.ts) —
  // not a real SharePoint list, so it has no `id` and is never scanned via
  // listItems.ts. Consumers branch on this to route to the Recycle Bin's own
  // fetch path instead, and to exclude it from "pick a default library"
  // logic (pickDefaultLibrary/findNamedDefault in libraryStats.ts).
  isRecycleBin?: boolean;
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
  // The actual failure behind sizeSource === 'error'. Kept because the cause
  // is genuinely ambiguous — throttling, a 403, a 406 on an odd folder name,
  // the list view threshold — and guessing at it in the UI ("likely
  // throttling") sends people to fix the wrong thing.
  sizeErrorMessage?: string;
  // The walk was canceled partway (or, rarely, a folder listing hit its own
  // page-count safety valve), so totalSizeBytes/fileCount are a floor ("at
  // least this much") rather than exact. Distinct from sizeSource ===
  // 'error': this is a real, usable measurement — just deliberately
  // incomplete because the user (or a listing limit) stopped it early. See
  // WalkOptions.signal in storageMetrics.ts.
  sizeApproximate?: boolean;
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
  // Count of those same retained old versions (current version excluded —
  // that's what the /Versions endpoint itself returns). Comes from the same
  // call as versionSizeBytes, so it shares its populated-only-when-measured
  // semantics: undefined means "not measured", not zero.
  versionCount?: number;
}

export interface ScanOptions {
  siteUrl: string;
  includeSubsites: boolean;
  includeHidden: boolean;
  libraryUrls?: string[];
  staleDays: number;
  veryStaleDays: number;
  scanConcurrency: number;
  // Opt-in: collect version-history size alongside current file size. Usually
  // free (it rides along in the bulk item sweep), but on a list where
  // SharePoint won't project the metrics field it escalates — see
  // versionScanMode and services/sp/versionSizes.ts.
  includeVersionHistory: boolean;
  // Only consulted when includeVersionHistory is true AND the escalation
  // reaches its per-file last resort. Defaults to 'quick' at the call site.
  versionScanMode?: VersionScanMode;
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
  // Running count of raw items (files AND folders) read across every
  // library so far this scan — the same unit totalItemsHint below is in, so
  // the two divide cleanly into a completion fraction/ETA. Distinct from
  // `scanned`, which counts FILES only (what the UI reports as progress),
  // since a library's raw read includes folder rows too.
  itemsFetched?: number;
  // Sum of LibraryInfo.itemCount across every library this scan will visit,
  // known once site + library discovery finishes. A hint, not a guarantee —
  // ItemCount can be stale — so it's undefined (not 0) when it couldn't be
  // computed, and callers must treat that as "no estimate available" rather
  // than "0 items expected".
  totalItemsHint?: number;
  // Which phase of a library's scan this update is about. The two phases
  // measure DIFFERENT units (items read vs files measured), so the UI must
  // scope its progress bar and rate estimate to one phase at a time rather
  // than averaging across them — see StorageReportView's ETA.
  //
  // 'versions' only ever appears for the per-file escalation; the bulk
  // strategies finish in a handful of requests and never report a phase.
  phase?: 'items' | 'versions';
  // Progress within a 'versions' phase. Both undefined outside it.
  versionsDone?: number;
  versionsTotal?: number;
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
  // Version-history measurements that were ATTEMPTED AND FAILED (throttling
  // exhausted, transient error) during the per-file escalation — kept
  // in-scope rather than aborting, same as skippedFolders/skippedSites above.
  // Retrying, or lowering concurrency, can genuinely help these.
  skippedVersions?: number;
  // Version-history measurements that were NEVER ATTEMPTED: Quick mode's
  // per-library budget ran out, the scan was canceled mid-pass, or no
  // mechanism for version size worked on that list at all.
  //
  // Deliberately separate from skippedVersions. Conflating the two would tell
  // the user to retry or lower concurrency for a condition that retrying
  // cannot change — the fix here is switching to a Full scan, or nothing.
  unmeasuredVersions?: number;
  // Which mechanism actually produced the version numbers, for the libraries
  // that needed an escalation ('items-side-channel', 'per-file', 'none', …).
  // Diagnostic: it is the difference between a ~40-request answer and a
  // ~194,000-request one, and the user cannot otherwise tell which they got.
  versionSizeStrategy?: string;
  // Which mode this scan ran in, so a report loaded from history explains its
  // own numbers rather than being read against the current UI selection.
  versionScanMode?: VersionScanMode;
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
  options: Pick<ScanOptions,
    'includeSubsites' | 'staleDays' | 'veryStaleDays' | 'includeVersionHistory' | 'versionScanMode'>;
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
  sizeErrorMessage?: string;
  sizeApproximate?: boolean; // folders only — see FolderStorageNode.sizeApproximate
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
  versionCount?: number;      // files only — see FileEntry.versionCount
  sizeUnknown?: boolean;      // folders only — see TreemapItem.sizeUnknown
  sizeErrorMessage?: string;  // folders only — the actual failure, for hover detail
  sizeApproximate?: boolean;  // folders only — see FolderStorageNode.sizeApproximate
}
