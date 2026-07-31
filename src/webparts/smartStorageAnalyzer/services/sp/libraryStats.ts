import { SpApiClient, TaskQueue } from './spCore';
import {
  getStorageMetrics, getShallowEstimate, newWalkBudget, LIBRARY_ROLLUP_WALK_DEADLINE_MS,
} from './storageMetrics';
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
  // budget, so each one fell back to its own private, fresh budget — on a
  // site with many libraries and no conventional default, that reintroduced
  // the exact unbounded-multiplication problem the shared budget exists to
  // prevent. It also trusted getStorageMetrics returning undefined as the
  // only fallback trigger, missing the common case where it reports a stale
  // 0 for a library that actually has content — getLibraryRollups already
  // checks totalSizeBytes > 0, not just whether the call succeeded.
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
  // True when the size came from a walk that stopped at its folder budget, so
  // it's a floor ("at least this much") rather than exact.
  sizeApproximate?: boolean;
}

// Per-library rollups for the Explorer's library-level (site root) treemap.
//
// StorageMetrics first, then a BUDGETED live walk when the rollup is missing or
// reads 0 for a library that clearly has items (its background job lags recent
// activity, so a stale 0 is common).
//
// This originally refused to walk at all, on the reasoning that a library-level
// fallback means scanning an entire library — turning "draw the opening screen"
// into "scan the whole site". That objection was about an UNBOUNDED walk, and no
// longer applies: the walk now stops at a shared folder budget, and requests
// coalesce through $batch. On a real tenant the probe-only version left the
// landing screen reporting "Unknown" for four libraries, which is strictly less
// useful than an honest floor. So: measure what the budget allows and mark the
// result approximate; only report Unknown when even that fails.
//
// One budget across ALL libraries, so this whole view stays bounded no matter
// how many libraries have stale rollups.
export async function getLibraryRollups(
  client: SpApiClient,
  siteUrl: string,
  libraries: LibraryInfo[],
): Promise<LibraryRollup[]> {
  const fallbackQueue = new TaskQueue(client.scanConcurrency);
  const budget = newWalkBudget(client, LIBRARY_ROLLUP_WALK_DEADLINE_MS);

  const tasks = libraries.map((library) => async (): Promise<LibraryRollup> => {
    const genuinelyEmpty = (library.itemCount ?? 0) === 0;
    const metrics = await getStorageMetrics(client, siteUrl, library.serverRelativeUrl);
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

    const walked = await getShallowEstimate(
      client, siteUrl, library.serverRelativeUrl, fallbackQueue, budget,
    );
    // Same reasoning as getFolderChildren's neverMeasured fix: with many
    // libraries sharing ONE budget, a later library's own walk can hit the
    // budget already exhausted by earlier libraries before it lists anything
    // at all. A zero total from that isn't a confirmed empty library or even
    // a real partial floor — it's "never looked at", and must be reported as
    // Unknown rather than an approximate 0 that sorts to the bottom looking
    // confirmed-near-empty. Errored/truncated WITH a nonzero partial total is
    // still a usable floor, same as elsewhere.
    const nothingToShow = (!!walked.errored || !!walked.truncated)
      && walked.totalSizeBytes === 0 && walked.fileCount === 0;
    return {
      library,
      totalSizeBytes: walked.totalSizeBytes,
      fileCount: walked.fileCount,
      lastModified: walked.lastModified ?? library.lastModified,
      sizeUnknown: nothingToShow,
      sizeApproximate: !nothingToShow && (!!walked.truncated || !!walked.errored),
    };
  });

  const results = await client.runConcurrent(tasks, client.scanConcurrency);
  return results
    .filter((r): r is LibraryRollup => !!r)
    .sort((a, b) => b.totalSizeBytes - a.totalSizeBytes);
}

// Best-effort pick of the site's default document library, used to seed the
// Explorer view without a picker screen. Only falls through to libs[0] (by
// size, when getLibrariesWithStats had to compute it) when no library is
// named "Documents"/"Shared Documents".
export function pickDefaultLibrary(libs: LibraryInfo[]): LibraryInfo | undefined {
  if (libs.length === 0) return undefined;
  return findNamedDefault(libs) ?? libs[0];
}
