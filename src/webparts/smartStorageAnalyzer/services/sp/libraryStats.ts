import { SpApiClient, TaskQueue } from './spCore';
import { getStorageMetrics, getShallowEstimate } from './storageMetrics';
import { getLibraries } from './siteDiscovery';
import { LibraryInfo } from '../../models/models';

// Name-based default match only (no size fallback) — used internally to
// decide whether getLibrariesWithStats can skip computing sizes entirely.
// Kept separate from pickDefaultLibrary below, which additionally falls
// back to the largest library and would otherwise always short-circuit
// this check (defeating the point).
function findNamedDefault(libs: LibraryInfo[]): LibraryInfo | undefined {
  const norm = (s: string): string => s.toLowerCase().replace(/%20/g, ' ');
  const byUrl = libs.find((l) => norm(l.serverRelativeUrl).endsWith('/shared documents'));
  if (byUrl) return byUrl;
  return libs.find((l) => l.title === 'Documents' || l.title === 'Shared Documents');
}

// Resolves the library list for a site, with a rolled-up size per library
// only when one is actually needed. The Explorer's library switcher never
// displays library sizes — it's just a row of name buttons — and
// pickDefaultLibrary only falls back to size (picking the largest) when no
// library named "Documents"/"Shared Documents" exists. Computing a full
// recursive rollup for every library up front — the single biggest
// contributor to the "Loading libraries…" delay on first load — was pure
// waste in the overwhelmingly common case where that named default exists.
// So: try the name-based pick first (cheap, no network calls beyond the
// list itself); only compute sizes for every library if that comes up empty.
export async function getLibrariesWithStats(
  client: SpApiClient,
  siteUrl: string,
  includeHidden: boolean,
): Promise<LibraryInfo[]> {
  const libraries = await getLibraries(client, siteUrl, includeHidden);
  if (findNamedDefault(libraries)) return libraries;

  // Shared across every library's fallback walk so concurrency stays capped
  // at scanConcurrency instead of multiplying if several libraries need one
  // at once (see getShallowEstimate in storageMetrics.ts).
  const fallbackQueue = new TaskQueue(client.scanConcurrency);

  const tasks = libraries.map((lib) => async () => {
    const metrics = await getStorageMetrics(client, siteUrl, lib.serverRelativeUrl);
    const resolved = metrics ?? (await getShallowEstimate(client, siteUrl, lib.serverRelativeUrl, fallbackQueue));
    const withStats: LibraryInfo = {
      ...lib,
      totalSizeBytes: resolved.totalSizeBytes,
      sizeSource: metrics ? 'storageMetrics' : 'estimate',
      lastModified: resolved.lastModified ?? lib.lastModified,
    };
    return withStats;
  });

  const results = await client.runConcurrent(tasks, client.scanConcurrency);
  return results
    .filter((l): l is LibraryInfo => !!l)
    .sort((a, b) => (b.totalSizeBytes ?? 0) - (a.totalSizeBytes ?? 0));
}

// Best-effort pick of the site's default document library, used to seed the
// Explorer view without a picker screen. Only falls through to libs[0] (by
// size, when getLibrariesWithStats had to compute it) when no library is
// named "Documents"/"Shared Documents".
export function pickDefaultLibrary(libs: LibraryInfo[]): LibraryInfo | undefined {
  if (libs.length === 0) return undefined;
  return findNamedDefault(libs) ?? libs[0];
}
