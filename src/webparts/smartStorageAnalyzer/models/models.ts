// Shared data model for the Smart Storage Analyzer web part.

export enum CandidateTier {
  Active = 'Active',
  Stale = 'Stale',
  VeryStale = 'VeryStale',
}

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
  // free (it rides along in the bulk item sweep); on a list that won't project
  // the metrics field it escalates through services/sp/versionSizes.ts, which
  // measures EVERY file that could have retained versions.
  //
  // There is deliberately no "how thorough" option here any more. There was a
  // Quick/Full choice, capping the per-file pass at the largest 5,000 files per
  // library; measurement retired it. On a real 184,859-file library only 8,168
  // files are candidates at all (version labels prove the rest have nothing),
  // the capped run and the complete run agreed to 0.06%, and the complete run
  // was fast. A user-facing decision that changes almost nothing is worse than
  // no decision.
  includeVersionHistory: boolean;
  // Cooperative cancellation only: checked between queued tasks/libraries/
  // sites, not passed into in-flight HTTP requests. A scan can be many
  // minutes long with no other way to stop it short of leaving the page.
  signal?: AbortSignal;
}

// Every distinct thing a scan can be doing. Each one either has an honest
// denominator or admits it doesn't — see stageTotal.
//
// This exists because the previous two-value `phase` could not name most of
// what a scan actually does. A library whose list won't project the version
// field goes: items sweep -> strategy probe -> bulk version sweep -> validate
// -> per-file measurement. Only the first and last had a phase, so the middle
// three (up to ~37 heavyweight requests) had nowhere to report from and the UI
// sat frozen through them.
export type ScanStage =
  | 'discovering'        // subsites + library lists; no denominator
  | 'items'              // flat item sweep of one library
  | 'version-probe'      // deciding how version size can be read for one list
  | 'version-bulk'       // the bulk version sweep the probe chose
  | 'version-validate'   // cross-checking a bulk result before trusting it
  | 'version-per-file';  // one request per file — the last resort

export interface ScanProgress {
  stage: ScanStage;
  // A verb phrase with no target and no punctuation ('Reading items',
  // 'Measuring version history'). Producers must NOT build whole sentences:
  // the view composes the status line, and two prose styles fighting on screen
  // is how this display became a wall of text once already.
  stageLabel: string;
  // What is happening right now, short enough for one line
  // ('page 12 of ~37', '1,204 of 8,168 files').
  detail?: string;
  libraryTitle?: string;
  // Set only when the scan actually spans more than one site, so a single-site
  // scan doesn't carry a redundant segment.
  siteLabel?: string;

  libsDone: number;
  libsTotal: number;

  // ── The current stage's own progress, in its own unit ───────────────────
  // stageTotal is undefined whenever there is no denominator worth drawing a
  // bar against — INCLUDING when an estimate has been overrun. The producer
  // owns that judgement, because it is the only thing that knows whether its
  // own estimate still holds; the view's rule is simply "no total, no bar, no
  // ETA". That inversion is deliberate: the old code clamped instead, which
  // pegged the bar at 100% and printed "~0s remaining" for eleven minutes
  // while a third of the work was still to come.
  stageDone?: number;
  stageTotal?: number;
  stageUnit?: 'items' | 'pages' | 'files';
  // Identity of the current determinate run, e.g. 'site|Documents|items' or
  // 'site|Documents|perfile:8168'. The view restarts its rate window when this
  // changes. Composed by the producer rather than inferred by the view because
  // only the producer knows when a run genuinely restarts — the previous
  // version keyed on the phase name, which was the constant 'items' for every
  // library in the scan, so one library inherited the whole scan's average
  // rate.
  stageKey: string;

  // ── Cumulative across the whole scan ───────────────────────────────────
  itemsFetched?: number;   // raw rows (files AND folders) read so far
  // Sum of LibraryInfo.itemCount across every library this scan will visit.
  // A COUNTER ONLY — never the bar and never the ETA. ItemCount is routinely
  // stale, so this gets reached and exceeded mid-scan; using it as a
  // denominator is what produced the false "finished" state above.
  totalItemsHint?: number;
  // Files identified during sweeps (FSObjType 0), known the instant each page
  // lands. Monotonic, and an upper bound on what the report will contain.
  // This is the number that can honestly move early — see `scanned`.
  filesSeen: number;
  // Files COMMITTED to the report. Necessarily lags filesSeen: an entry may
  // not be published while its versionSizeBytes can still change, so a
  // library's entries all appear at once after its version work finishes.
  // Converges with filesSeen per library.
  scanned: number;
  skippedVersions: number;
  unmeasuredVersions: number;
  skippedLibraries: number;
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
  // Version-history measurements NEVER ATTEMPTED, as opposed to attempted and
  // failed (skippedVersions above). Three causes now, none of which retrying at
  // a lower concurrency fixes:
  //   - the scan was canceled while the per-file pass was still running
  //   - no version-size mechanism worked on that list at all (strategy 'none')
  //   - the row was in the main item sweep but not in the bulk version sweep
  //     (created or deleted between the two)
  // Deliberately separate from skippedVersions: THAT one is worth retrying, and
  // conflating them tells the user to change a setting that cannot help.
  unmeasuredVersions?: number;
  // Which mechanism actually produced the version numbers, for the libraries
  // that needed an escalation ('items-side-channel', 'render-list-data',
  // 'per-file', 'none'). Diagnostic, and load-bearing for the UI: it is the
  // difference between a ~40-request answer and a per-file one, and nothing
  // else on screen would tell the user which they got.
  versionSizeStrategy?: string;
  // LEGACY, READ-ONLY. Written only by builds that had a Quick/Full choice;
  // nothing assigns it any more. Kept because a report loaded from history must
  // explain ITS OWN numbers — an old Quick report's unmeasuredVersions means
  // "the budget ran out, re-run to measure them", a cause no current scan can
  // have, and without this field the UI would describe it using the current
  // build's wording and name the wrong reason.
  //
  // Note the asymmetry with ScanOptions, which no longer has this field at all:
  // ScanOptions describes how a FUTURE scan runs (no longer a choice), this
  // describes what a PAST scan did (a historical fact). Don't "tidy" it away.
  versionScanMode?: 'quick' | 'full';
  // Sum of FileEntry.versionSizeBytes across all entries. Only meaningful
  // when versionHistoryIncluded is true — see that field's comment.
  totalVersionSizeBytes?: number;
  // Sum of FileEntry.versionCount across all entries — the retained-version
  // COUNT analogue of totalVersionSizeBytes above. Free: versionCount is
  // already on every entry, this just adds it up the same way size is summed.
  totalVersionCount?: number;
  // Whether this particular scan collected version-history data. Needed
  // because totalVersionSizeBytes === 0 is ambiguous between "no old
  // versions exist" and "the toggle was off" (or this report predates the
  // feature and the field is simply missing).
  versionHistoryIncluded?: boolean;
}

// A saved report WITHOUT its file listing.
//
// The listing lives in a separate IndexedDB store and is loaded only when
// something actually needs it (see ReportHistoryService). That split is what
// makes complete reports possible at all: the history list, its badges, and
// report diffing all read nothing but the fields below, and loading every
// report's full entries just to render a row of dates put ~1.8M objects in
// memory and forced a row cap that silently dropped 12% of a real report.
export interface StoredReportMeta {
  id: string;
  timestamp: number;
  siteUrl: string;
  options: Pick<ScanOptions,
    'includeSubsites' | 'staleDays' | 'veryStaleDays' | 'includeVersionHistory'>;
  summary: StorageReportSummary;
  // How many rows the listing holds. Shown in the UI and used to size the
  // "loading listing" state, so neither has to fetch the listing to find out.
  entryCount: number;
  // True when the file listing is NOT retrievable — it was evicted to make room
  // for a newer report, or it never fit. Deliberately NOT called "truncated":
  // nothing about the scan or the summary is partial, and the previous wording
  // ("Partial") read as though the numbers themselves were incomplete.
  listingEvicted?: boolean;
}

// Meta plus the listing. Only materialised when a listing is actually loaded.
export interface StoredReport extends StoredReportMeta {
  entries: FileEntry[];
}

export interface ReportDiff {
  olderTimestamp: number;
  newerTimestamp: number;
  // From the summaries, which are always stored — so always available.
  sizeDeltaBytes: number;
  totalFilesDelta: number;
  // Need both reports' FILE LISTINGS to compare individual paths, and a listing
  // can be evicted when browser storage runs short. Undefined means "could not
  // be determined", which is deliberately distinct from 0 ("nothing changed").
  newStaleCount?: number;
  resolvedStaleCount?: number;
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
