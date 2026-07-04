import { SpApiClient, folderApi, valueArray, TaskQueue } from './spCore';
import { FileEntry, LibraryInfo, ScanOptions } from '../../models/models';
import { ageInDays, classify } from '../../utils/archivalClassification';

const SYSTEM_FOLDER_NAMES = new Set(['forms']);

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
): Promise<void> {
  const queue = new TaskQueue(options.scanConcurrency);

  function visitFolder(serverRelativeUrl: string): void {
    queue.add(async () => {
      const [foldersData, filesData] = await Promise.all([
        client.getJson(`${folderApi(siteUrl, serverRelativeUrl)}/Folders?$select=Name,ServerRelativeUrl`),
        client.getJson(
          `${folderApi(siteUrl, serverRelativeUrl)}/Files?$select=Name,ServerRelativeUrl,Length,TimeCreated,TimeLastModified,Author/Title,Author/LoginName&$expand=Author`,
        ),
      ]);

      for (const f of valueArray(filesData)) {
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

      for (const folder of valueArray(foldersData)) {
        if (SYSTEM_FOLDER_NAMES.has(String(folder.Name).toLowerCase())) continue;
        visitFolder(folder.ServerRelativeUrl);
      }
    });
  }

  visitFolder(library.serverRelativeUrl);
  await queue.drain();
}
