import { SpApiClient, folderApi, odata, SYSTEM_FOLDER_NAMES, TaskQueue } from './spCore';
import { FileEntry, LibraryInfo, ScanOptions } from '../../models/models';
import { ageInDays, classify } from '../../utils/archivalClassification';

export interface WalkLibraryResult {
  // Folders that could not be read and were skipped rather than aborting the
  // whole walk — see the catch block below.
  skippedFolders: number;
  // Per-file version-history fetches that failed — see the version-fetch
  // catch block below. Only ever nonzero when options.includeVersionHistory.
  skippedVersions: number;
  // What actually went wrong for each skipped folder (list view threshold,
  // 403, throttling exhausted, etc.) — always uncapped here; storageScan.ts
  // caps what gets stored in the report summary.
  skippedFolderDetails: { url: string; error: string }[];
}

// Sums SP.FileVersion.Size across a file's retained old versions. Called
// through a separate, lower-concurrency queue than the folder walk (below)
// since it doubles per-file request volume and this is opt-in/expensive.
async function fetchVersionSizeBytes(
  client: SpApiClient,
  siteUrl: string,
  fileServerRelativeUrl: string,
): Promise<number> {
  const url = `${siteUrl}/_api/web/GetFileByServerRelativePath(decodedUrl='${encodeURIComponent(
    odata(fileServerRelativeUrl),
  )}')/Versions?$select=Size`;
  const versions = await client.getJsonPaged(url);
  return versions.reduce((sum, v) => sum + Number(v.Size ?? 0), 0);
}

// Full recursive walk of a library's Files/Folders, needed because the
// Storage Report and its Excel/CSV export require a per-file row (size,
// dates, author) — unlike the interactive Folder Explorer, which can lean on
// StorageMetrics rollups and never needs to touch every file.
//
// A single TaskQueue is shared across the whole library so recursion depth
// never multiplies concurrency (each folder-visit enqueues its subfolders as
// new queue tasks rather than spawning nested worker pools).
export async function walkLibrary(
  client: SpApiClient,
  siteUrl: string,
  library: LibraryInfo,
  options: ScanOptions,
  onEntry: (entry: FileEntry) => void,
): Promise<WalkLibraryResult> {
  const queue = new TaskQueue(options.scanConcurrency);
  // Separate, lower-concurrency queue for the opt-in version-history
  // fetches, so enabling that toggle doesn't multiply the folder walk's own
  // concurrency on top of itself (each additional queue adds its own
  // in-flight requests against the same SPO throttling budget).
  const versionQueue = options.includeVersionHistory
    ? new TaskQueue(Math.max(1, Math.floor(options.scanConcurrency / 2)))
    : undefined;
  let skippedFolders = 0;
  let skippedVersions = 0;
  const skippedFolderDetails: { url: string; error: string }[] = [];

  function visitFolder(serverRelativeUrl: string): void {
    queue.add(async () => {
      // Cooperative cancellation: checked at the top of every queued task
      // (not passed into the in-flight HTTP requests themselves) so a
      // cancel takes effect between folders rather than needing to abort a
      // live fetch.
      if (options.signal?.aborted) return;
      try {
        const [foldersData, filesData] = await Promise.all([
          client.getJsonPaged(
            `${folderApi(siteUrl, serverRelativeUrl)}/Folders?$select=Name,ServerRelativeUrl&$top=5000`,
            options.signal,
          ),
          client.getJsonPaged(
            `${folderApi(siteUrl, serverRelativeUrl)}/Files?$select=Name,ServerRelativeUrl,Length,TimeCreated,TimeLastModified,Author/Title,Author/LoginName&$expand=Author&$top=5000`,
            options.signal,
          ),
        ]);

        const entries: FileEntry[] = filesData.map((f) => {
          const timeLastModified = f.TimeLastModified as string;
          const ageDays = ageInDays(timeLastModified);
          return {
            name: f.Name,
            serverRelativeUrl: f.ServerRelativeUrl,
            libraryTitle: library.title,
            sizeBytes: Number(f.Length ?? 0),
            timeCreated: f.TimeCreated,
            timeLastModified,
            authorLoginName: f.Author?.LoginName,
            authorDisplayName: f.Author?.Title,
            ageDays,
            tier: classify(ageDays, options.staleDays, options.veryStaleDays),
          };
        });

        if (versionQueue) {
          // Resolve this folder's version sizes as a batch before publishing
          // any entry, so onEntry never sees a FileEntry whose
          // versionSizeBytes is later mutated out from under the caller.
          await Promise.all(
            entries.map((entry) => new Promise<void>((resolve) => {
              versionQueue.add(async () => {
                try {
                  entry.versionSizeBytes = await fetchVersionSizeBytes(client, siteUrl, entry.serverRelativeUrl);
                } catch {
                  skippedVersions++;
                }
                resolve();
              });
            })),
          );
        }

        for (const entry of entries) onEntry(entry);

        if (options.signal?.aborted) return;
        for (const folder of foldersData) {
          if (SYSTEM_FOLDER_NAMES.has(String(folder.Name).toLowerCase())) continue;
          visitFolder(folder.ServerRelativeUrl);
        }
      } catch (err: any) {
        // Inaccessible or unreadable subtree (403, throttling exhausted,
        // list view threshold, etc.) — skip just this folder rather than
        // aborting the whole library walk, but count it AND record the
        // actual error so the caller can show why, not just how many.
        skippedFolders++;
        const message = err?.message ?? String(err);
        skippedFolderDetails.push({ url: serverRelativeUrl, error: message });
        // eslint-disable-next-line no-console
        console.warn(`[SmartStorageAnalyzer] Skipped folder ${serverRelativeUrl}: ${message}`);
      }
    });
  }

  visitFolder(library.serverRelativeUrl);
  await queue.drain();
  if (versionQueue) await versionQueue.drain();
  return { skippedFolders, skippedVersions, skippedFolderDetails };
}
