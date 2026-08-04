import { SpApiClient } from './spCore';
import { getStorageMetrics, getLibraryAggregate, WalkOptions } from './storageMetrics';
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

  // No named default exists, so every library's size is actually needed to
  // pick one (by largest, via pickDefaultLibrary below). Delegate to
  // getLibraryRollups rather than duplicating its fallback-walk logic here:
  // this function used to call getShallowEstimate per library with NO shared
  // walk context, so each one fell back to its own private, fresh queue —
  // on a site with many libraries and no conventional default, that
  // reintroduced the exact unbounded-concurrency-multiplication problem the
  // shared queue exists to prevent. It also trusted getStorageMetrics
  // returning undefined as the only fallback trigger, missing the common
  // case where it reports a stale 0 for a library that actually has content
  // — getLibraryRollups already checks totalSizeBytes > 0, not just whether
  // the call succeeded.
  const rollups = await getLibraryRollups(client, siteUrl, libraries);
  return rollups
    .map((r): LibraryInfo => ({
      ...r.library,
      totalSizeBytes: r.totalSizeBytes,
      lastModified: r.lastModified ?? r.library.lastModified,
      // Not displayed anywhere on this path (see the function comment above —
      // the library switcher is name-only), so collapsing the more granular
      // storageMetrics/estimate distinction getLibraryRollups tracks
      // internally loses nothing a user could see.
      sizeSource: r.sizeUnknown ? 'unknown' : 'estimate',
    }))
    .sort((a, b) => (b.totalSizeBytes ?? 0) - (a.totalSizeBytes ?? 0));
}

export interface LibraryRollup {
  library: LibraryInfo;
  totalSizeBytes: number;
  fileCount?: number;
  lastModified?: string;
  // True when the library holds items but its size could not be established
  // at all. Deliberately distinct from "0 bytes".
  sizeUnknown: boolean;
  // True when the size came from a walk that was canceled partway (or hit a
  // listing's page valve), so it's a floor ("at least this much") rather
  // than exact.
  sizeApproximate?: boolean;
}

// Per-library rollups for the Explorer's library-level (site root) treemap.
//
// StorageMetrics first: one request per library, and for the landing screen
// that is the right trade — it avoids sweeping a library the user may never
// open. Its background job lags recent activity, so a stale 0 for a library
// that clearly has items is common; when that happens the library is swept
// flat (listItems.ts) and the exact total taken from the aggregate.
//
// The sweep is shared with the drill-down path via the aggregate cache, so a
// library measured here is already resident when the user opens it — the
// landing screen effectively pre-warms whatever it had to measure the hard
// way.
export async function getLibraryRollups(
  client: SpApiClient,
  siteUrl: string,
  libraries: LibraryInfo[],
  options?: WalkOptions,
): Promise<LibraryRollup[]> {
  let sweptItems = 0;

  const tasks = libraries.map((library) => async (): Promise<LibraryRollup> => {
    if (options?.signal?.aborted) {
      return {
        library,
        totalSizeBytes: 0,
        fileCount: 0,
        lastModified: library.lastModified,
        sizeUnknown: true,
      };
    }
    // StorageMetrics is a per-FOLDER probe (GetFolderByServerRelativePath(...)
    // /StorageMetrics) — the Recycle Bin has no real folder behind its
    // pseudo-path, so the probe would either 404 or (worse) resolve to
    // whatever real folder happens to share that path and report the wrong
    // number. Go straight to the sweep for it.
    const genuinelyEmpty = !library.isRecycleBin && (library.itemCount ?? 0) === 0;
    const metrics = library.isRecycleBin
      ? undefined
      : await getStorageMetrics(client, siteUrl, library.serverRelativeUrl);
    if (metrics && metrics.totalSizeBytes > 0) {
      return {
        library,
        totalSizeBytes: metrics.totalSizeBytes,
        fileCount: metrics.fileCount,
        lastModified: metrics.lastModified ?? library.lastModified,
        sizeUnknown: false,
      };
    }
    if (genuinelyEmpty) {
      return {
        library,
        totalSizeBytes: 0,
        fileCount: 0,
        lastModified: library.lastModified,
        sizeUnknown: false,
      };
    }

    try {
      const { aggregate, partial } = await getLibraryAggregate(client, siteUrl, library, {
        signal: options?.signal,
        // Several libraries can be sweeping at once, so report a combined
        // running total rather than letting each one reset the counter.
        onWalkProgress: (n) => {
          sweptItems += n;
          options?.onWalkProgress?.(sweptItems);
        },
      });
      const root = aggregate.root;
      // A canceled sweep that never got a single page is "never looked at",
      // not "measured and found empty" — reporting it as an approximate 0
      // would sort it to the bottom looking confirmed-near-empty, which is
      // exactly what sizeUnknown exists to prevent.
      const nothingToShow = partial && root.totalSizeBytes === 0 && root.fileCount === 0;
      return {
        library,
        totalSizeBytes: root.totalSizeBytes,
        fileCount: root.fileCount,
        lastModified: root.lastModified ?? library.lastModified,
        sizeUnknown: nothingToShow,
        sizeApproximate: !nothingToShow && partial,
      };
    } catch (err: any) {
      // The library could not be enumerated at all (see listItems.ts). There
      // is no recursive-walk fallback any more, so report it honestly.
      // eslint-disable-next-line no-console
      console.warn(`[SmartStorageAnalyzer] Could not measure ${library.title}: ${err?.message ?? String(err)}`);
      return {
        library,
        totalSizeBytes: 0,
        fileCount: 0,
        lastModified: library.lastModified,
        sizeUnknown: true,
      };
    }
  });

  const results = await client.runConcurrent(tasks, client.scanConcurrency);
  return results
    .filter((r): r is LibraryRollup => !!r)
    .sort((a, b) => b.totalSizeBytes - a.totalSizeBytes);
}

// Best-effort pick of the site's default document library, used to seed the
// Explorer view without a picker screen. Only falls through to libs[0] (by
// size, when getLibrariesWithStats had to compute it) when no library is
// named "Documents"/"Shared Documents". The Recycle Bin pseudo-library is
// excluded from consideration entirely — it should never be what a user
// lands on by default.
export function pickDefaultLibrary(libs: LibraryInfo[]): LibraryInfo | undefined {
  const real = libs.filter((l) => !l.isRecycleBin);
  if (real.length === 0) return undefined;
  return findNamedDefault(real) ?? real[0];
}
