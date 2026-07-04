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

  // ── Site / permission checks ──────────────────────────────────────────────

  checkCanManageWeb(siteUrl: string): Promise<boolean> {
    return siteDiscovery.checkCanManageWeb(this.client, siteUrl);
  }

  getSiteOwners(siteUrl: string): Promise<siteDiscovery.SiteOwnerInfo[]> {
    return siteDiscovery.getSiteOwners(this.client, siteUrl);
  }

  // ── Library Overview ──────────────────────────────────────────────────────

  getLibrariesWithStats(siteUrl: string, includeHidden: boolean): Promise<LibraryInfo[]> {
    return libraryStats.getLibrariesWithStats(this.client, siteUrl, includeHidden);
  }

  // ── Folder Explorer (StorageMetrics-driven, lazy per level) ──────────────

  getFolderChildren(
    siteUrl: string,
    parentServerRelativeUrl: string,
    onProgress?: storageMetrics.FolderChildProgress,
  ): Promise<FolderStorageNode[]> {
    return storageMetrics.getFolderChildren(this.client, siteUrl, parentServerRelativeUrl, onProgress);
  }

  getFolderFiles(siteUrl: string, serverRelativeUrl: string): ReturnType<typeof storageMetrics.getFolderFiles> {
    return storageMetrics.getFolderFiles(this.client, siteUrl, serverRelativeUrl);
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
