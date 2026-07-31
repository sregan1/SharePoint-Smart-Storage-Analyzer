import { WebPartContext } from '@microsoft/sp-webpart-base';
import {
  FileEntry, FolderStorageNode, LibraryInfo, ScanOptions, ScanProgress,
} from '../models/models';
import { SpApiClient } from './sp/spCore';
import * as siteDiscovery from './sp/siteDiscovery';
import * as storageMetrics from './sp/storageMetrics';
import * as libraryStats from './sp/libraryStats';
import * as folderSizeCache from './sp/folderSizeCache';
import * as storageScan from './sp/storageScan';
import { ScanResult } from './sp/storageScan';

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
   * which now sizes libraries via getLibraryRollups (probe-only, by design),
   * that hidden path would be a whole-site walk before the cheap path even ran.
   */
  getLibraries(siteUrl: string, includeHidden: boolean): Promise<LibraryInfo[]> {
    return siteDiscovery.getLibraries(this.client, siteUrl, includeHidden);
  }

  /**
   * Per-library sizes for the Explorer's site-root treemap. StorageMetrics
   * probe per library and nothing more — see getLibraryRollups for why this
   * must never fall back to a recursive walk.
   */
  getLibraryRollups(siteUrl: string, libraries: LibraryInfo[]): Promise<libraryStats.LibraryRollup[]> {
    return libraryStats.getLibraryRollups(this.client, siteUrl, libraries);
  }

  // ── Folder Explorer (StorageMetrics-driven, lazy per level) ──────────────

  getFolderChildren(
    siteUrl: string,
    parentServerRelativeUrl: string,
    onProgress?: storageMetrics.FolderChildProgress,
  ): Promise<FolderStorageNode[]> {
    return storageMetrics.getFolderChildren(this.client, siteUrl, parentServerRelativeUrl, onProgress);
  }

  /**
   * Discards cached folder sizes so the next load re-measures from SharePoint.
   * Backs the Explorer's Refresh action — without it, a folder that came back
   * unmeasurable or truncated stays that way for the cache's full TTL.
   */
  clearFolderSizeCache(siteUrl?: string): void {
    folderSizeCache.clearCachedFolderChildren(siteUrl);
  }

  getFolderFiles(
    siteUrl: string,
    serverRelativeUrl: string,
    includeVersionHistory?: boolean,
  ): ReturnType<typeof storageMetrics.getFolderFiles> {
    return storageMetrics.getFolderFiles(this.client, siteUrl, serverRelativeUrl, includeVersionHistory);
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
