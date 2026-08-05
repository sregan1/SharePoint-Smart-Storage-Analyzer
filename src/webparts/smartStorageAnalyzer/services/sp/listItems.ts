import { SpApiClient, valueArray } from './spCore';
import { LibraryInfo } from '../../models/models';

// Bulk, FLAT enumeration of a document library's contents.
//
// This replaces the recursive folder walk that used to back both the Storage
// Report and the Explorer's fallback sizing. That walk cost 2 requests per
// FOLDER (Files + Folders) regardless of how few files the folder held, so
// its cost scaled with the shape of the tree rather than the amount of
// content: on a real 2.3TB archive it managed roughly 4 files/second, which
// works out to hours for a single library.
//
// The /items endpoint is not folder-scoped — it returns every item in the
// list, across every folder, in ID order. So one paged sweep replaces the
// entire traversal:
//
//   100,000 files in a deep tree
//     recursive walk : 2 x folder count -> tens of thousands of requests
//     flat sweep     : ceil(items / 5000) -> ~20-25 requests
//
// Folder sizes are then derived from these rows client-side (see
// folderAggregate.ts) at no additional request cost at all.
export const ITEMS_PAGE_SIZE = 5000;

// Deliberately NO $filter and NO $orderby.
//
// Both are evaluated across the whole list before paging is applied, so on a
// list past the 5,000-item view threshold either one throws
// SPQueryThrottledException ("the attempted operation is prohibited because
// it exceeds the list view threshold") — precisely on the large libraries
// this exists to handle. An unfiltered, unordered query pages fine at any
// size because the implicit order is by ID, which is always indexed.
//
// So: fetch everything and split files from folders in memory on FSObjType
// (0 = file, 1 = folder). The folder rows are not waste — the treemap needs
// the folder list anyway, and getting it here means never asking for it
// separately.
// Id leads the list because the version-size side channel (versionSizes.ts)
// joins its own sweep back to these rows by Id — see FlatItem.id. Id is
// always selectable on every list (the shard path below has always relied on
// that), so including it universally costs nothing.
const BASE_FIELDS = 'Id,FileRef,FileLeafRef,FSObjType,Modified,Created,AuthorId,EditorId';
// File size comes from the File/Length navigation property, NOT the
// File_x0020_Size list column. File_x0020_Size looks like the obvious choice
// (it's the "File Size" column visible in the UI) but it's a Computed-type
// field, and SPO's /items REST endpoint frequently rejects a $select naming
// it outright — confirmed against a real tenant, where every attempt
// (with and without SMTotalFileStreamSize) 400'd with "The field or property
// 'File_x0020_Size' does not exist." File/Length is the field the classic
// Files/ collection has always exposed (folderApi's Files?$select=Length in
// the old walk) and is reliable via $expand=File here too. A folder row's
// File is null, which resolves to size 0 — exactly what's wanted, since
// folders are zero-weight containers regardless (see folderAggregate.ts).
const SIZE_EXPAND = 'File';
const SIZE_FIELDS = 'File/Length';
// SMTotalFileStreamSize (current + retained versions) is storage-metrics
// metadata, not a base document-library field, so it's the first thing
// dropped if the full $select is rejected — the size field above is not
// negotiable, since without it there is nothing to report at all.
const VERSION_FIELD = 'SMTotalFileStreamSize';
// OData__UIVersionString is REST's escaped name for the internal
// _UIVersionString field — a standard column on every list, giving the
// current version label (e.g. "5.0", or "12.3" if minor/draft versions are
// in play). The integer part is the major version number; a file at major
// version N has (N-1) prior retained major versions, which is the same
// quantity the per-file .../Versions collection counts (that collection
// also excludes the current version). This gets version COUNT from the same
// bulk request that already gets size, at zero extra cost — replacing what
// used to be a separate request PER FILE (fetchVersionInfo, now removed).
//
// It is an approximation, not identical to querying /Versions directly:
//   - Minor/draft versions (the fractional part) aren't counted.
//   - A library with a configured version-retention LIMIT keeps incrementing
//     the version number even after old versions are purged from storage, so
//     this can OVERSTATE the true retained count in that specific case — it
//     never understates. Size (SMTotalFileStreamSize, above) is unaffected
//     either way: it reflects real storage regardless of how count is
//     derived, so the number that actually matters for a storage audit
//     stays exact. Confirmed acceptable trade-off: the alternative was one
//     request per file, which took over 3 hours and was still incomplete on
//     a real 192,978-item library.
const VERSION_LABEL_FIELD = 'OData__UIVersionString';

export interface FlatItem {
  // Server-relative path, e.g. /sites/x/Shared Documents/A/B/report.docx
  fileRef: string;
  // List item id. Load-bearing for the version-size SIDE CHANNEL
  // (versionSizes.ts): that sweep projects a different field set and must be
  // joined back to these rows, and Id is the only stable join key — FileRef
  // can differ in encoding between projections, and a file moved mid-sweep
  // changes its path but never its id.
  id?: number;
  name: string;
  isFolder: boolean;
  sizeBytes: number;
  // Retained-version bytes only (total minus current), so it is additive to
  // sizeBytes exactly the way the rest of the app treats version history.
  // undefined when the list didn't return SMTotalFileStreamSize.
  versionSizeBytes?: number;
  // Approximate retained MAJOR version count, derived from
  // OData__UIVersionString — see VERSION_LABEL_FIELD above for exactly what
  // this does and doesn't capture. undefined when the list didn't return the
  // version label (folders don't have one; num() would otherwise report 0).
  versionCountApprox?: number;
  // True when the version label PROVES this file has nothing in its /Versions
  // collection, so a per-file version request would certainly return zero and
  // can be skipped outright. See hasNoRetainedVersions — deliberately NOT the
  // same test as `versionCountApprox === 0`.
  versionsProvablyZero?: boolean;
  created: string;
  modified: string;
  authorId?: number;
  editorId?: number;
  // Set instead of authorId/editorId by sources that already have a display
  // name/login rather than a site-user lookup id — currently only
  // recycleBin.ts, whose DeletedBy fields come back as plain strings, not
  // lookup ids. Consumers should prefer these over resolving authorId when
  // present, rather than trying (and failing) a lookup.
  authorDisplayName?: string;
  authorLoginName?: string;
}

// One update from a library's item sweep, fired per page (so roughly once per
// request — pages are 5,000 items).
export interface ItemSweepProgress {
  // Raw rows so far in this library: files AND folders.
  items: number;
  // Of those, actual files (FSObjType 0). The ONLY honest "files" count that
  // exists before the report is allowed to publish entries, which is why the
  // status line can show movement during a long sweep instead of a flat zero.
  files: number;
  pagesDone: number;
  // undefined once unknown OR overrun — see the emitters below. A stale
  // ItemCount that under-estimates makes pagesDone exceed this, and reporting a
  // clamped fraction from that point is what pegged the progress bar at 100%
  // while a third of the work remained.
  pagesTotal?: number;
}

export interface FetchItemsOptions {
  signal?: AbortSignal;
  // Fired per page. See ItemSweepProgress — this used to be a bare
  // `(fetchedSoFar: number)`, which gave the UI no way to distinguish files
  // from folder rows and no denominator of its own.
  onProgress?: (p: ItemSweepProgress) => void;
  // Fired (at most once per library) when SMTotalFileStreamSize could not be
  // selected, so the returned items carry no inline version size. The version
  // ESCALATION itself deliberately lives in the caller (fileWalk.ts) rather
  // than here: versionSizes.ts needs FlatItem from this module, so invoking it
  // from here would make the two files circularly dependent for values, not
  // just types. Keeping the signal one-directional avoids that entirely.
  //
  // Explorer never passes this, so it never pays for a version escalation.
  onVersionFieldUnavailable?: () => void;
}

// Thrown when a library cannot be enumerated at all. The caller reports the
// library as skipped (with this message) rather than silently under-counting
// — there is deliberately no recursive-walk fallback any more.
export class LibraryFetchError extends Error {
  public constructor(public readonly library: string, message: string) {
    super(message);
    this.name = 'LibraryFetchError';
  }
}

function num(v: unknown): number {
  const n = Number(v ?? 0);
  return isFinite(n) ? n : 0;
}

function parseVersionCountApprox(versionLabel: unknown): number | undefined {
  if (versionLabel == null) return undefined;
  const major = parseInt(String(versionLabel).split('.')[0], 10);
  if (!isFinite(major)) return undefined;
  return Math.max(0, major - 1);
}

// True ONLY when the version label proves the /Versions collection is empty:
// exactly major version 1, with no minor/draft part. Anything else — a nonzero
// fractional part, an absent or unparseable label, "0.3" for a never-published
// draft — is treated as unknown, and the file is still measured.
//
// Deliberately STRICTER than parseVersionCountApprox, and not interchangeable
// with it. That function returns 0 for "1.3" as well, because it only ever
// claimed to count MAJOR versions. As a skip test that would be wrong:
// /Versions includes minor versions too, so "1.3" has three retained versions
// with real bytes that would get silently zeroed on any library with
// minor/draft versioning enabled.
//
// The degradation direction here is always "slower, never wrong" — an
// unrecognized label returns false and the file gets a real measurement.
function hasNoRetainedVersions(versionLabel: unknown): boolean {
  if (versionLabel == null) return false;
  const parts = String(versionLabel).split('.');
  const major = parseInt(parts[0], 10);
  if (!isFinite(major) || major !== 1) return false;
  if (parts.length === 1) return true;
  const minor = parseInt(parts[1], 10);
  return isFinite(minor) && minor === 0;
}

function toItem(raw: any): FlatItem {
  const size = num(raw.File?.Length);
  // SMTotalFileStreamSize covers current + retained versions. Guard against
  // it coming back smaller than the current size (it is maintained by the
  // same lagging background job as StorageMetrics) rather than reporting a
  // negative version total.
  const totalWithVersions = raw.SMTotalFileStreamSize != null
    ? num(raw.SMTotalFileStreamSize)
    : undefined;
  // Positive proof of no retained versions. When the bulk field is missing,
  // this is what lets the version pass skip the large majority of files
  // outright instead of spending a request each to confirm zero.
  const provablyZero = hasNoRetainedVersions(raw.OData__UIVersionString);
  return {
    fileRef: String(raw.FileRef ?? ''),
    id: raw.Id != null ? Number(raw.Id) : undefined,
    name: String(raw.FileLeafRef ?? ''),
    isFolder: Number(raw.FSObjType ?? 0) === 1,
    sizeBytes: size,
    // `0` rather than `undefined` when provably zero: undefined means "not
    // measured" everywhere in this app (rendered as "—"), and here there is
    // positive proof rather than absence of information.
    versionSizeBytes: totalWithVersions != null
      ? Math.max(0, totalWithVersions - size)
      : (provablyZero ? 0 : undefined),
    versionCountApprox: parseVersionCountApprox(raw.OData__UIVersionString),
    versionsProvablyZero: provablyZero,
    created: raw.Created as string,
    modified: raw.Modified as string,
    authorId: raw.AuthorId != null ? Number(raw.AuthorId) : undefined,
    editorId: raw.EditorId != null ? Number(raw.EditorId) : undefined,
  };
}

function itemsUrl(siteUrl: string, listId: string, fields: string): string {
  return `${siteUrl}/_api/web/lists(guid'${listId}')/items`
    + `?$select=${fields}&$expand=${SIZE_EXPAND}&$top=${ITEMS_PAGE_SIZE}`;
}

// ── ID-range sharding ──────────────────────────────────────────────────────
// Splits ONE library's sweep into several concurrent requests instead of the
// plain sequential nextLink loop below. Only engaged for libraries large
// enough to need more than one page — see fetchLibraryItems.
//
// Filtering/ordering by Id is the one exception to the "no $filter, no
// $orderby" rule above: the ID column is always indexed on every SharePoint
// list, and Microsoft documents indexed-column filters as exempt from the
// List View Threshold (unlike a filter on an arbitrary column, which forces
// a full-list scan and throws exactly the exception the plain sweep above
// is built to avoid). This is a technique this app hasn't used before,
// though, so it's wrapped in a canary check (below) rather than trusted
// outright — the same caution that would have caught the File_x0020_Size
// field turning out to be unselectable on this tenant, had it been applied
// there too.
const SHARD_MIN_ITEMS = ITEMS_PAGE_SIZE;

interface FieldProbe {
  maxId: number;
  // The winning $select field list (with or without VERSION_FIELD). Id is
  // already the first entry of BASE_FIELDS, so this path no longer prepends
  // it separately.
  fields: string;
  // Whether `fields` includes VERSION_FIELD. NON-OPTIONAL on purpose: when
  // this is false the caller MUST route version size through the escalation
  // (versionSizes.ts) or explicitly warn, and a field that could be quietly
  // left undefined is exactly how the earlier version of this code silently
  // dropped version history for the largest libraries. Make the compiler ask.
  versionFieldIncluded: boolean;
}

function probeUrl(siteUrl: string, listId: string, fields: string): string {
  return `${siteUrl}/_api/web/lists(guid'${listId}')/items`
    + `?$select=${fields}&$expand=${SIZE_EXPAND}&$orderby=Id desc&$top=1`;
}

// One small request that answers two questions at once: the library's
// highest item Id (the upper bound to shard against) and whether
// SMTotalFileStreamSize is selectable on this list (the same field
// negotiation fetchLibraryItems does, just piggybacked here instead of
// costing its own separate request).
//
// TWO attempts, and the history of this function is worth knowing before
// changing it again.
//
// Originally it retried without SMTotalFileStreamSize on a 400 and sharded
// with the reduced field list. That silently dropped version-history size for
// the entire library — and because sharding only engages on LARGE libraries,
// it dropped it exactly where most version storage lives. The fix at the time
// was to give up on sharding altogether when the full field set was rejected,
// on the reasoning that losing concurrency was better than losing data.
//
// That reasoning was correct *while there was no other way to get version
// size*. There is now (versionSizes.ts): a rejected field routes version size
// to a separate un-expanded sweep, or to a bounded per-file pass. So dropping
// the field from the SHARD query no longer loses anything — it just moves
// where the number comes from — and the cost of refusing to shard is real:
// on a 193,915-item library it is ~39 sequential pages instead of ~40
// concurrent ones, which was a large slice of an observed 80-minute scan.
//
// The safety property that must survive any future edit: versionFieldIncluded
// === false is a HARD obligation on the caller to escalate (or warn). It must
// never quietly mean "this library has no version history".
async function probeMaxIdAndFields(
  client: SpApiClient,
  siteUrl: string,
  library: LibraryInfo,
  signal?: AbortSignal,
): Promise<FieldProbe | undefined> {
  const attempts: { fields: string; versionFieldIncluded: boolean }[] = [
    { fields: `${BASE_FIELDS},${SIZE_FIELDS},${VERSION_FIELD},${VERSION_LABEL_FIELD}`, versionFieldIncluded: true },
    { fields: `${BASE_FIELDS},${SIZE_FIELDS},${VERSION_LABEL_FIELD}`, versionFieldIncluded: false },
  ];
  for (let i = 0; i < attempts.length; i++) {
    if (signal?.aborted) return undefined;
    const { fields, versionFieldIncluded } = attempts[i];
    try {
      const data = await client.getJson(probeUrl(siteUrl, library.id!, fields), true);
      const rows = valueArray(data);
      if (rows.length === 0) return { maxId: 0, fields, versionFieldIncluded }; // empty library
      const id = Number(rows[0].Id);
      if (!isFinite(id)) return undefined;
      return { maxId: id, fields, versionFieldIncluded };
    } catch (err: any) {
      // Only a field-shape problem is worth retrying with fewer fields.
      // Anything else — throttling exhausted, permissions, the $orderby/$filter
      // combination itself being rejected — will fail identically the second
      // time, and means the ID-range technique doesn't work on this list at
      // all. Give up on sharding rather than doubling the probe cost.
      const status = /HTTP (\d+)/.exec(err?.message ?? '')?.[1];
      if (status !== '400') {
        // eslint-disable-next-line no-console
        console.warn(
          `[SmartStorageAnalyzer] ${library.title}: shard probe failed, sharding disabled for ` +
          `this library — ${err?.message ?? String(err)}`,
        );
        return undefined;
      }
      // Log the FULL body of the first (version-field-bearing) 400. That body
      // is the only thing distinguishing the two very different causes:
      //   "The field or property 'SMTotalFileStreamSize' does not exist"
      //     -> Storage Metrics isn't projectable on this list at all
      //   an $expand/$orderby/expression complaint
      //     -> the FIELD is fine, the query SHAPE is not
      // This was a bare `catch {}` for a long time, which is precisely why
      // that question went unanswered through several rounds of debugging.
      if (i === 0) {
        // eslint-disable-next-line no-console
        console.warn(
          `[SmartStorageAnalyzer] ${library.title}: shard probe rejected the version field (400) — ` +
          `${err?.message ?? String(err)}`,
        );
      }
    }
  }
  return undefined;
}

function shardUrl(siteUrl: string, listId: string, fields: string, idGreaterThan: number, idAtMost: number): string {
  return `${siteUrl}/_api/web/lists(guid'${listId}')/items`
    + `?$select=${fields}&$expand=${SIZE_EXPAND}`
    + `&$filter=${encodeURIComponent(`Id gt ${idGreaterThan} and Id le ${idAtMost}`)}`
    + `&$orderby=Id asc&$top=${ITEMS_PAGE_SIZE}`;
}

// Fetches one contiguous ID range, paging internally if the shard holds more
// than one page's worth. Prefers the server's own @odata.nextLink when
// provided; if a tenant doesn't return one for a filtered/ordered query,
// falls back to advancing the `Id gt` cursor using the last row's own Id —
// the same dual-path defensiveness recycleBin.ts's paging already uses,
// since neither codepath has verified nextLink support on every tenant.
async function fetchIdRange(
  client: SpApiClient,
  siteUrl: string,
  listId: string,
  fields: string,
  rangeStart: number,
  rangeEnd: number,
  signal: AbortSignal | undefined,
  // Handed the rows themselves rather than just a count, so the shared
  // accumulator can classify files vs folders without a second pass.
  onPage: (pageRows: any[]) => void,
): Promise<any[]> {
  const rows: any[] = [];
  let cursor = rangeStart;
  let next: string | undefined = shardUrl(siteUrl, listId, fields, cursor, rangeEnd);
  for (let page = 0; page < 400 && next && !signal?.aborted; page++) {
    const data = await client.getJson(next, true);
    const pageRows = valueArray(data);
    rows.push(...pageRows);
    onPage(pageRows);
    const nextLink = data?.['odata.nextLink'] ?? data?.['@odata.nextLink'];
    if (nextLink) {
      next = nextLink;
    } else if (pageRows.length < ITEMS_PAGE_SIZE) {
      next = undefined; // last page of this shard
    } else {
      const lastId = Number(pageRows[pageRows.length - 1]?.Id);
      if (!isFinite(lastId) || lastId >= rangeEnd) { next = undefined; }
      else { cursor = lastId; next = shardUrl(siteUrl, listId, fields, cursor, rangeEnd); }
    }
  }
  return rows;
}

// Bounded-concurrency runner that (unlike SpApiClient.runConcurrent)
// re-throws the first task failure instead of swallowing it into
// `undefined`. A silently dropped shard here would under-count the library
// while looking like a complete, exact sweep — worse than the honest
// all-or-nothing fallback in fetchLibraryItemsSharded below.
async function runShardsOrThrow<T>(tasks: (() => Promise<T>)[], concurrency: number): Promise<T[]> {
  const results: T[] = new Array(tasks.length);
  let idx = 0;
  let firstError: any;
  const worker = async (): Promise<void> => {
    while (idx < tasks.length && !firstError) {
      const i = idx++;
      try {
        results[i] = await tasks[i]();
      } catch (err) {
        firstError = err;
        return;
      }
    }
  };
  await Promise.all(Array.from({ length: Math.min(concurrency, tasks.length) }, worker));
  if (firstError) throw firstError;
  return results;
}

// Attempts the sharded sweep. Returns undefined (not a thrown error) when
// sharding should simply not be used for this library — the caller falls
// back to the plain sequential sweep in that case. Throws only if sharding
// started succeeding (the canary shard worked, so the ID-filter technique is
// confirmed to work on this tenant) and then a LATER shard failed — that is
// a real fetch failure, not a "don't bother sharding" signal, so it should
// surface as one, same as the plain sweep's own errors do.
async function fetchLibraryItemsSharded(
  client: SpApiClient,
  siteUrl: string,
  library: LibraryInfo,
  probe: FieldProbe,
  options?: FetchItemsOptions,
): Promise<any[] | undefined> {
  const listId = library.id!;
  // Shard COUNT is deliberately NOT tied to scanConcurrency. Sizing shards
  // that way (maxId / scanConcurrency) was the first version of this and
  // measured ~2.2x on a real 192,000-item library instead of the ~6x its
  // concurrency setting should allow — each shard still spanned ~6-7 pages,
  // so the sweep ran as two lumpy sequential phases (one solo canary shard,
  // then a batch of N-1 shards each STILL paging internally) rather than a
  // continuous pipeline. Sizing shards toward ~1 page each and feeding
  // however many results into the same bounded-concurrency runner keeps the
  // concurrency budget saturated for the whole sweep instead of just its
  // first two phases.
  //
  // itemCount (from getLibraries) estimates density (items per ID unit) —
  // a hint for EFFICIENCY only, same as elsewhere in this file: an
  // under-estimate (heavy historical deletion skewing the ID:item ratio)
  // just means some shards need a second internal page, never a wrong
  // total, since correctness comes from the ranges covering [0, maxId]
  // exhaustively, not from the density guess being right.
  const density = library.itemCount && library.itemCount > 0 ? library.itemCount / probe.maxId : 1;
  const spanForOnePage = Math.max(1, Math.ceil(ITEMS_PAGE_SIZE / Math.max(density, 1e-6)));
  // Upper bound so a badly wrong density estimate can't turn into thousands
  // of tiny requests — worst case some shards fall back to needing extra
  // internal pages, which is exactly the graceful-degradation this guards.
  const MAX_SHARDS = 64;
  const shardCount = Math.min(MAX_SHARDS, Math.max(1, Math.ceil(probe.maxId / spanForOnePage)));
  const span = Math.max(1, Math.ceil(probe.maxId / shardCount));
  const ranges: { start: number; end: number }[] = [];
  for (let i = 0; i < shardCount; i++) {
    const start = i * span;
    if (start >= probe.maxId) break;
    ranges.push({ start, end: i === shardCount - 1 ? probe.maxId : Math.min(probe.maxId, start + span) });
  }
  if (ranges.length === 0) return [];

  // Shared across every shard, so the reported totals are the library's, not
  // one shard's. pagesTotal is EXACT here — the ranges are all computed up
  // front, so the page count is known before the first request goes out, which
  // is strictly better than the sequential path's ItemCount-derived guess.
  let itemsSoFar = 0;
  let filesSoFar = 0;
  let pagesDone = 0;
  const onPage = (pageRows: any[]): void => {
    itemsSoFar += pageRows.length;
    for (const r of pageRows) if (Number(r.FSObjType ?? 0) === 0) filesSoFar++;
    pagesDone++;
    options?.onProgress?.({
      items: itemsSoFar,
      files: filesSoFar,
      pagesDone,
      // A shard needing a second internal page (a density mis-estimate) makes
      // this overrun; report no denominator rather than a fraction over 1.
      pagesTotal: pagesDone <= ranges.length ? ranges.length : undefined,
    });
  };

  // Canary: fetch the first shard ALONE before trusting the ID-filter
  // technique on this tenant at all. Only once it succeeds do the remaining
  // shards fire concurrently — if it throws, sharding is abandoned for this
  // library entirely (undefined, not a re-thrown error) so the caller falls
  // back to the proven sequential sweep, paying only this one shard's cost.
  let canaryRows: any[];
  try {
    canaryRows = await fetchIdRange(
      client, siteUrl, listId, probe.fields, ranges[0].start, ranges[0].end, options?.signal, onPage,
    );
  } catch {
    return undefined;
  }
  if (options?.signal?.aborted || ranges.length === 1) return canaryRows;

  const remaining = ranges.slice(1);
  const shardResults = await runShardsOrThrow(
    remaining.map((r) => () => fetchIdRange(
      client, siteUrl, listId, probe.fields, r.start, r.end, options?.signal, onPage,
    )),
    client.scanConcurrency,
  );
  return canaryRows.concat(...shardResults);
}

// One flat sweep of a library. Resolves with every item (files AND folders);
// throws LibraryFetchError if the library can't be enumerated.
//
// skipBatch is set on every page: $batch exists to amortize many small
// requests, and a 5,000-item page is the opposite shape — see getJson.
export async function fetchLibraryItems(
  client: SpApiClient,
  siteUrl: string,
  library: LibraryInfo,
  options?: FetchItemsOptions,
): Promise<FlatItem[]> {
  if (!library.id) {
    throw new LibraryFetchError(
      library.title,
      'Library has no list id — cannot run a bulk item query against it.',
    );
  }

  // Only worth the extra probe request for a library actually large enough
  // to span multiple pages. itemCount is a hint here, not load-bearing — it
  // can be stale, and being wrong only costs one skipped/attempted probe
  // either direction, never a correctness problem.
  if ((library.itemCount ?? 0) > SHARD_MIN_ITEMS && !options?.signal?.aborted) {
    const probe = await probeMaxIdAndFields(client, siteUrl, library, options?.signal);
    if (probe && probe.maxId > 0) {
      try {
        const rows = await fetchLibraryItemsSharded(client, siteUrl, library, probe, options);
        if (rows) {
          // The shard query had to drop the version field to work. That is now
          // recoverable (the caller escalates via versionSizes.ts) rather than
          // the silent data loss it used to be — but ONLY because the caller is
          // told. See FieldProbe.versionFieldIncluded.
          if (!probe.versionFieldIncluded) options?.onVersionFieldUnavailable?.();
          return rows.map(toItem);
        }
        // undefined: the canary shard failed — the ID-filter technique isn't
        // safe to use on this tenant. Fall through to the plain sequential
        // sweep below rather than sharding.
      } catch (err: any) {
        // A LATER shard failed after the canary already succeeded — this is
        // a genuine fetch failure, not a "don't shard" signal, so it's
        // reported the same way the sequential sweep's own failures are.
        throw new LibraryFetchError(library.title, err?.message ?? String(err));
      }
    }
  }

  // Field negotiation, not a fallback to the old walk: a $select naming an
  // absent field 400s the entire query, and version-history metadata varies
  // by list template (Picture Libraries and Site Pages are not identical to
  // a plain Document Library). Try with SMTotalFileStreamSize first, then
  // the same query without it, and only then give up on the library.
  const attempts = [
    `${BASE_FIELDS},${SIZE_FIELDS},${VERSION_FIELD},${VERSION_LABEL_FIELD}`,
    `${BASE_FIELDS},${SIZE_FIELDS},${VERSION_LABEL_FIELD}`,
  ];

  // Page count derived from ItemCount — a HINT, unlike the sharded path's exact
  // range count. Reported only while it still holds; see the overrun guard in
  // the callback below.
  const pagesHint = library.itemCount && library.itemCount > 0
    ? Math.max(1, Math.ceil(library.itemCount / ITEMS_PAGE_SIZE))
    : undefined;

  let lastError: any;
  for (let i = 0; i < attempts.length; i++) {
    if (options?.signal?.aborted) return [];
    let filesSoFar = 0;
    try {
      const raw = await client.getJsonPaged(
        itemsUrl(siteUrl, library.id, attempts[i]),
        options?.signal,
        // A library needing more than this many pages holds >2,000,000 items,
        // which is far past what SharePoint supports in a single list.
        400,
        true, // skipBatch
        (fetchedSoFar, pageRows, pagesDone) => {
          // One integer compare per row, over rows already parsed and resident
          // — immeasurable next to the JSON.parse of a 5,000-row page.
          for (const r of pageRows) if (Number(r.FSObjType ?? 0) === 0) filesSoFar++;
          options?.onProgress?.({
            items: fetchedSoFar,
            files: filesSoFar,
            pagesDone,
            // A stale ItemCount that under-estimates makes pagesDone overrun.
            // Dropping the denominator from that point is not cosmetic: a
            // clamped fraction pegs the bar at 100% while work continues, which
            // is precisely the bug this rework exists to fix.
            pagesTotal: pagesHint != null && pagesDone <= pagesHint ? pagesHint : undefined,
          });
        },
      );
      const items = raw.map(toItem);
      // The reduced field set won, so nothing here carries inline version
      // size. Tell the caller, which owns the escalation decision (and, when
      // version history wasn't requested at all, the warning).
      if (i > 0) options?.onVersionFieldUnavailable?.();
      return items;
    } catch (err: any) {
      lastError = err;
      // Only a field-shape problem is worth retrying with fewer fields. A
      // throttle/permission/threshold failure will fail identically the
      // second time, so surface it immediately instead of doubling the cost.
      const status = /HTTP (\d+)/.exec(err?.message ?? '')?.[1];
      if (status !== '400') break;
      // Log the body on the FIRST attempt only — that's the one carrying the
      // version field, so its message is what says whether the field is
      // absent from the list or merely incompatible with this query's shape.
      // The status alone can't distinguish those, and they lead to completely
      // different fallback strategies (see versionSizes.ts).
      if (i === 0) {
        // eslint-disable-next-line no-console
        console.warn(
          `[SmartStorageAnalyzer] ${library.title}: full field set rejected (400) — ` +
          `${err?.message ?? String(err)}`,
        );
      }
    }
  }

  throw new LibraryFetchError(
    library.title,
    lastError?.message ?? String(lastError ?? 'Unknown error'),
  );
}

// Author/Editor come back as numeric lookup ids. Resolving them means one
// request per SITE (the user information list is small and shared by every
// library) instead of the $expand=Author that used to ride along on every
// single file row — by far the most expensive part of the old per-file
// query.
export async function fetchSiteUsers(
  client: SpApiClient,
  siteUrl: string,
  signal?: AbortSignal,
): Promise<Map<number, { title: string; loginName: string }>> {
  const map = new Map<number, { title: string; loginName: string }>();
  try {
    const data = await client.getJson(
      `${siteUrl}/_api/web/siteusers?$select=Id,Title,LoginName&$top=${ITEMS_PAGE_SIZE}`,
    );
    if (signal?.aborted) return map;
    for (const u of valueArray(data)) {
      const id = Number(u.Id);
      if (!isFinite(id)) continue;
      map.set(id, { title: String(u.Title ?? ''), loginName: String(u.LoginName ?? '') });
    }
  } catch {
    // Names are presentational — a report with blank authors is far better
    // than no report, so this never fails a scan.
  }
  return map;
}
