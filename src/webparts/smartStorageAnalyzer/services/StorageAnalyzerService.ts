import { WebPartContext } from '@microsoft/sp-webpart-base';
import {
  FileEntry, FolderStorageNode, LibraryInfo, ScanOptions, ScanProgress,
} from '../models/models';
import { SpApiClient } from './sp/spCore';
import * as siteDiscovery from './sp/siteDiscovery';
import * as storageMetrics from './sp/storageMetrics';
import * as libraryStats from './sp/libraryStats';
import * as storageScan from './sp/storageScan';
import { ScanResult } from './sp/storageScan';

export type { WalkOptions, WalkProgress } from './sp/storageMetrics';

// Facade over the sp/ modules so views keep a single dependency with a
// stable API. Implementation lives in:
//   sp/spCore.ts        — API client, throttling/paging, shared helpers
//   sp/siteDiscovery.ts — libraries, subwebs, owner/access checks
//   sp/storageMetrics.ts— StorageMetrics-backed folder rollups (Folder Explorer)
//   sp/libraryStats.ts  — per-library rollups (Library Overview)
//   sp/fileWalk.ts       — recursive per-file walk of one library
//   sp/storageScan.ts    — Storage Report scan orchestration (multi-site/library)
export class StorageAnalyzerService {
  private readonly client: SpApiClient;

  constructor(context: WebPartContext) {
    this.client = new SpApiClient(context);
  }

  /** Max concurrent API requests during scans. Settable from Settings. */
  get scanConcurrency(): number { return this.client.scanConcurrency; }
  set scanConcurrency(value: number) { this.client.scanConcurrency = value; }

  /**
   * True while requests are paused waiting out SharePoint throttling. Views
   * poll this so a long stall can say *why* it's stalling — a 60s+ throttle
   * wait is otherwise indistinguishable from the app having hung.
   */
  get isThrottled(): boolean { return this.client.isThrottled; }

  // ── Site / permission checks ──────────────────────────────────────────────

  checkCanManageWeb(siteUrl: string): Promise<boolean> {
    return siteDiscovery.checkCanManageWeb(this.client, siteUrl);
  }

  // ── Library Overview ──────────────────────────────────────────────────────

  getLibrariesWithStats(siteUrl: string, includeHidden: boolean): Promise<LibraryInfo[]> {
    return libraryStats.getLibrariesWithStats(this.client, siteUrl, includeHidden);
  }

  /**
   * The library list with no size computation at all — one request.
   *
   * Preferred over getLibrariesWithStats by any caller that resolves sizes
   * itself: that function only skips size computation when a library *named*
   * "Documents"/"Shared Documents" exists, and otherwise falls back to a full
   * recursive walk per library to pick a default by size. For the Explorer,
   * which sizes libraries via getLibraryRollups instead, that hidden path
   * would be a whole-site walk before the cheap path even ran.
   */
  getLibraries(siteUrl: string, includeHidden: boolean): Promise<LibraryInfo[]> {
    return siteDiscovery.getLibraries(this.client, siteUrl, includeHidden);
  }

  /**
   * Per-library sizes for the Explorer's site-root treemap. Tries a
   * StorageMetrics probe per library first; falls back to a live recursive
   * walk when that's stale/unavailable (see getLibraryRollups). The walk has
   * no automatic time/count limit — pass `options.signal` to cancel it, and
   * `options.onWalkProgress` for a live folder-visited counter while it runs.
   */
  getLibraryRollups(
    siteUrl: string,
    libraries: LibraryInfo[],
    options?: storageMetrics.WalkOptions,
  ): Promise<libraryStats.LibraryRollup[]> {
    return libraryStats.getLibraryRollups(this.client, siteUrl, libraries, options);
  }

  // ── Tree View / List View (one flat sweep per library, then in memory) ───

  /**
   * Immediate child folders with exact recursive sizes. The first call for a
   * library sweeps it (~1 request per 5,000 items); every drill-down after
   * that is served from the in-memory aggregate with no network at all.
   */
  getFolderChildren(
    siteUrl: string,
    library: LibraryInfo,
    parentServerRelativeUrl: string,
    onProgress?: storageMetrics.FolderChildProgress,
    options?: storageMetrics.WalkOptions,
  ): Promise<FolderStorageNode[]> {
    return storageMetrics.getFolderChildren(
      this.client, siteUrl, library, parentServerRelativeUrl, onProgress, options,
    );
  }

  /**
   * Discards the cached per-library aggregates so the next load re-sweeps
   * from SharePoint. Backs the Explorer's Refresh action.
   */
  clearFolderSizeCache(siteUrl?: string): void {
    storageMetrics.clearAggregateCache(siteUrl);
  }

  /**
   * Version history (size and an approximate count) is always included at
   * no extra cost — see storageMetrics.getFolderFiles — so there is no
   * separate flag to request it here any more; the Explorer's "Include
   * version history" toggle now only controls whether it's DISPLAYED.
   */
  getFolderFiles(
    siteUrl: string,
    library: LibraryInfo,
    serverRelativeUrl: string,
    options?: storageMetrics.WalkOptions,
  ): ReturnType<typeof storageMetrics.getFolderFiles> {
    return storageMetrics.getFolderFiles(this.client, siteUrl, library, serverRelativeUrl, options);
  }

  // ── Storage Report (full recursive scan) ──────────────────────────────────

  scanSite(
    options: ScanOptions,
    onProgress: (progress: ScanProgress) => void,
    onEntry?: (entry: FileEntry) => void,
  ): Promise<ScanResult> {
    return storageScan.scanSite(this.client, options, onProgress, onEntry);
  }
}
