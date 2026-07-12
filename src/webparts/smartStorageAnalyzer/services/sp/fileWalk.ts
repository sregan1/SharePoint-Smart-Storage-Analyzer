import { SpApiClient, folderApi, SYSTEM_FOLDER_NAMES, TaskQueue } from './spCore';
import { FileEntry, LibraryInfo, ScanOptions } from '../../models/models';
import { ageInDays, classify } from '../../utils/archivalClassification';

export interface WalkLibraryResult {
  // Folders that could not be read and were skipped rather than aborting the
  // whole walk — see the catch block below.
  skippedFolders: number;
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
  let skippedFolders = 0;

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

        for (const f of filesData) {
          const timeLastModified = f.TimeLastModified as string;
          const ageDays = ageInDays(timeLastModified);
          const entry: FileEntry = {
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
          onEntry(entry);
        }

        if (options.signal?.aborted) return;
        for (const folder of foldersData) {
          if (SYSTEM_FOLDER_NAMES.has(String(folder.Name).toLowerCase())) continue;
          visitFolder(folder.ServerRelativeUrl);
        }
      } catch {
        // Inaccessible or unreadable subtree (403, throttling exhausted,
        // list view threshold, etc.) — skip just this folder rather than
        // aborting the whole library walk, but count it so the caller can
        // warn that totals are partial instead of silently under-reporting.
        skippedFolders++;
      }
    });
  }

  visitFolder(library.serverRelativeUrl);
  await queue.drain();
  return { skippedFolders };
}
