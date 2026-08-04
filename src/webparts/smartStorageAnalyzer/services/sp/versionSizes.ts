import { SpApiClient, valueArray, odata, SMALL_BATCH_SIZE } from './spCore';
import { LibraryInfo, VersionScanMode } from '../../models/models';
import { FlatItem, ITEMS_PAGE_SIZE } from './listItems';
import { QUICK_VERSION_FILE_LIMIT } from '../../utils/settingsBounds';

// Getting version-history SIZE for a library whose bulk field doesn't work.
//
// The happy path lives in listItems.ts: SMTotalFileStreamSize rides along in
// the main item sweep and version size costs nothing at all. This module is
// everything that happens when that field is rejected — which, on a real
// tenant, it is: every attempt against one 193,915-item "Documents" library
// 400s.
//
// The history here matters, because the obvious fix is a trap. The original
// implementation answered a rejected bulk field by fetching every file's
// /Versions collection individually. That is ~194,000 requests for one
// library; measured on this tenant it ran 80 minutes without finishing, and a
// comment in fileWalk.ts records an earlier attempt at the same thing running
// "3+ hours and still incomplete". Correct and unusable.
//
// So the order of preference is: prove a BULK mechanism works before spending
// a single per-file request, and treat per-file as a bounded last resort
// rather than the fallback. Each mechanism is introduced the way the rest of
// this codebase introduces uncertain SharePoint techniques (see
// probeMaxIdAndFields and the canary shard in listItems.ts): probe it small,
// adopt it only on proof, degrade cleanly when it fails.

export type VersionSizeStrategyKind =
  // Nothing to do — the main sweep already returned SMTotalFileStreamSize.
  | 'inline'
  // The field IS selectable, just not alongside $expand=File. Sweep it on its
  // own and join by Id. ~1 request per 5,000 items.
  | 'items-side-channel'
  // OData won't project the field at all, but the CAML/view renderer will.
  // Also ~1 request per 5,000 items, just POSTs instead of GETs.
  | 'render-list-data'
  // No bulk mechanism works. One request per file, capped in Quick mode.
  | 'per-file'
  // Nothing works and per-file wasn't permitted — report honestly, never zero.
  | 'none';

export interface VersionSizeStrategy {
  kind: VersionSizeStrategyKind;
}

export interface VersionFillOptions {
  signal?: AbortSignal;
  mode: VersionScanMode;
  // Fired with (done, total) for the per-file path only — the bulk paths
  // resolve a whole library in a handful of requests and have nothing
  // meaningful to report in between. Already throttled at the source (see
  // PROGRESS_EVERY_N / PROGRESS_EVERY_MS), so callers can wire it straight to
  // setState without adding their own coalescing.
  onProgress?: (done: number, total: number) => void;
  // One file's version fetch failed (throttling exhausted, transient error).
  // Kept in scope rather than failing the library, same as a skipped folder.
  onSkipped?: () => void;
}

export interface VersionFillResult {
  strategy: VersionSizeStrategyKind;
  // Attempted and failed. Retrying (or lowering concurrency) could fix these.
  skipped: number;
  // Never attempted, because Quick mode's budget ran out. Retrying changes
  // nothing — only switching to Full does. Deliberately a separate number
  // from `skipped`: conflating them would tell the user to fix throttling for
  // a condition that has nothing to do with throttling.
  unmeasured: number;
}

// ── Progress throttling ───────────────────────────────────────────────────
// The per-file path can complete thousands of requests. Emitting progress per
// completion would put a setState storm on a thread that is simultaneously
// holding a ~194,000-element array — the same reasoning that made the Storage
// Report's own file counter a ref flushed on a timer rather than per-file
// state (see scannedRef in StorageReportView.tsx).
const PROGRESS_EVERY_N = 50;
const PROGRESS_EVERY_MS = 500;

// Version size is (current + retained) minus current, so it needs BOTH the
// metrics field and the file's own length. The side-channel sweep only fetches
// the former; the latter is already on the FlatItem from the main sweep.
const SIDE_CHANNEL_FIELDS = `Id,SMTotalFileStreamSize`;

// The internal field name, as CAML ViewFields spells it. NOT the REST-escaped
// OData__ form — RenderListDataAsStream goes through the view/CAML renderer,
// which uses raw internal names.
const VERSION_FIELD_CAML = 'SMTotalFileStreamSize';
// RenderListDataAsStream pages by echoing a NextHref back; this bounds the
// loop the same way getJsonPagedMeta's maxPages does.
const RLDAS_MAX_PAGES = 400;
// Sampled by the probe. Enough to see past a run of folders (which legitimately
// report zero) before deciding the column is unpopulated.
const RLDAS_PROBE_ROWS = 100;

function num(v: unknown): number {
  const n = Number(v ?? 0);
  return isFinite(n) ? n : 0;
}

// RenderListDataAsStream returns field values as DISPLAY-formatted strings,
// not raw numbers, so they can carry locale thousands separators ("1,234,567"
// or "1.234.567"). Byte counts are integers, so stripping every non-digit
// recovers the value correctly under either convention.
//
// Returns undefined when the value is absent, or contains letters — the latter
// matters: if SharePoint hands back a pre-formatted size like "1.5 MB" there is
// no safe way to invert it, and guessing would fabricate storage figures. The
// caller treats undefined as "this mechanism is unusable" rather than as zero.
function parseDisplayInteger(v: unknown): number | undefined {
  if (v == null) return undefined;
  const raw = String(v).trim();
  if (raw === '') return undefined;
  if (/[a-z]/i.test(raw)) return undefined;
  const digits = raw.replace(/[^\d]/g, '');
  if (digits === '') return undefined;
  const n = Number(digits);
  return isFinite(n) ? n : undefined;
}

function rldasUrl(siteUrl: string, listId: string, nextHref?: string): string {
  const base = `${siteUrl}/_api/web/lists(guid'${listId}')/RenderListDataAsStream`;
  // NextHref arrives already query-string shaped ("?Paged=TRUE&p_ID=5000&…"),
  // and the base URL has no query of its own, so appending it verbatim is
  // correct — rewriting it would risk dropping a paging token.
  return nextHref ? `${base}${nextHref}` : base;
}

function rldasBody(rowLimit: number): unknown {
  return {
    parameters: {
      // Scope='RecursiveAll' is MANDATORY and the most dangerous thing to omit
      // here. Without it the view is folder-scoped and returns only the root
      // folder's rows — which looks like a perfectly successful sweep while
      // silently missing almost the entire library. That failure mode would be
      // indistinguishable from "this library has little version history".
      ViewXml:
        `<View Scope='RecursiveAll'>`
        + `<ViewFields><FieldRef Name='ID'/><FieldRef Name='${VERSION_FIELD_CAML}'/></ViewFields>`
        + `<RowLimit Paged='TRUE'>${rowLimit}</RowLimit>`
        + `</View>`,
      DatesInUtc: true,
    },
  };
}

/**
 * Decides how to get version size for one library, given that the main sweep
 * did NOT return it inline.
 *
 * Costs at most one small request. Never throws: an unusable library resolves
 * to a strategy, not an error, because a report that omits version history is
 * still a useful report.
 */
export async function probeVersionSizeStrategy(
  client: SpApiClient,
  siteUrl: string,
  library: LibraryInfo,
  signal?: AbortSignal,
): Promise<VersionSizeStrategy> {
  if (signal?.aborted || !library.id) return { kind: 'per-file' };

  // P1: the version field entirely on its own — no $expand, no $orderby, no
  // $filter.
  //
  // This is the probe that was missing for several rounds of debugging, and
  // it is the whole reason this module exists. EVERY previously-observed
  // failure of SMTotalFileStreamSize on this tenant came from a URL that also
  // carried `&$expand=File` (itemsUrl, probeUrl and shardUrl all append it
  // unconditionally). A storage-metrics field colliding with a
  // navigation-property expand is a known shape of SPO 400 — so "the field
  // 400s" and "the field 400s WHEN EXPANDED" were never distinguished, and
  // the app fell back to 194,000 requests on the strength of that conflation.
  //
  // If this succeeds, one library's version history costs ~39 requests
  // instead of ~194,000.
  try {
    const url = `${siteUrl}/_api/web/lists(guid'${library.id}')/items`
      + `?$select=${SIDE_CHANNEL_FIELDS}&$top=1`;
    const data = await client.getJson(url, true);
    const rows = valueArray(data);
    // An empty library proves nothing about the field, but there is also
    // nothing to measure — either strategy is a no-op, so take the cheap one.
    if (rows.length === 0) return { kind: 'items-side-channel' };
    // Presence, not truthiness: a genuinely zero-version file legitimately
    // reports 0, and `0` must not be read as "field missing". A row that
    // simply lacks the property is the failure case.
    if ('SMTotalFileStreamSize' in rows[0] && rows[0].SMTotalFileStreamSize != null) {
      // eslint-disable-next-line no-console
      console.info(
        `[SmartStorageAnalyzer] ${library.title}: version-history size available via a ` +
        'separate un-expanded sweep — the field is fine, it just cannot be combined with ' +
        '$expand=File. Using the bulk side channel (no per-file requests).',
      );
      return { kind: 'items-side-channel' };
    }
  } catch {
    // Fall through. No logging here: fetchLibraryItems has already logged the
    // full 400 body for this library (see its field-negotiation catch), so a
    // second message would just be noise.
  }

  // P2: RenderListDataAsStream — the endpoint SharePoint's own modern UI
  // drives. It projects fields through the view/CAML renderer rather than the
  // OData projector, so it can return columns that $select rejects outright.
  //
  // Built because P1 failed on a real tenant with
  //   "The field or property 'SMTotalFileStreamSize' does not exist."
  // on a bare `?$select=Id,SMTotalFileStreamSize&$top=1` — proving the earlier
  // theory (that the field merely clashed with $expand=File) wrong. A field
  // OData claims not to exist is exactly the case where a different projector
  // is worth one request.
  //
  // Tempered expectation, recorded honestly: "does not exist" may mean the
  // column is absent from the list SCHEMA, which RLDAS reads too — in which
  // case this fails the same way and we fall through to per-file. One request
  // to find out.
  try {
    // A sample, not one row. The field can be present on every row and
    // populated on none — and one row is not enough to tell the difference,
    // especially when the first rows of a library are often folders.
    const data = await client.postJson(rldasUrl(siteUrl, library.id), rldasBody(RLDAS_PROBE_ROWS));
    const rows: any[] = Array.isArray(data?.Row) ? data.Row : [];
    if (rows.length === 0) return { kind: 'render-list-data' }; // nothing to measure
    // Two separate things have to be true, and conflating them is the mistake
    // this check exists to avoid:
    //   1. the field comes back parseable at all (RLDAS will happily omit it), and
    //   2. at least one row is NON-ZERO.
    // SMTotalFileStreamSize counts the current version too, so any real file
    // must report more than zero. An all-zero sample means the column exists in
    // the schema but nothing populates it — which passes a presence-only check
    // and then yields max(0, 0 - fileSize) = 0 for every file in the library.
    // That is a confident, completely wrong 0 B, and it is exactly what a
    // presence-only canary let through on a real tenant.
    const parsed = rows.map((r) => parseDisplayInteger(r[VERSION_FIELD_CAML]));
    const usable = parsed.filter((v): v is number => v !== undefined);
    if (usable.length > 0 && usable.some((v) => v > 0)) {
      // eslint-disable-next-line no-console
      console.info(
        `[SmartStorageAnalyzer] ${library.title}: version-history size available via ` +
        'RenderListDataAsStream (OData will not project the field, the view renderer will). ' +
        'Using the bulk CAML sweep — no per-file requests.',
      );
      return { kind: 'render-list-data' };
    }
    // eslint-disable-next-line no-console
    console.warn(
      `[SmartStorageAnalyzer] ${library.title}: RenderListDataAsStream returned ` +
      `${usable.length} of ${rows.length} sampled rows with a parseable ${VERSION_FIELD_CAML}, ` +
      `all zero${usable.length === 0 ? '/absent' : ''} — the column is not populated with storage ` +
      'data on this list. Falling back to per-file measurement.',
    );
  } catch (err: any) {
    // eslint-disable-next-line no-console
    console.warn(
      `[SmartStorageAnalyzer] ${library.title}: RenderListDataAsStream probe failed — ` +
      `${err?.message ?? String(err)}. Falling back to per-file measurement.`,
    );
  }

  return { kind: 'per-file' };
}

/**
 * Fills versionSizeBytes / versionCountApprox on `items` in place, using
 * whichever strategy was probed. Mutates rather than returning a copy because
 * the caller's aggregate and FileEntry mapping both already reference these
 * exact objects.
 */
export async function applyVersionSizes(
  client: SpApiClient,
  siteUrl: string,
  library: LibraryInfo,
  items: FlatItem[],
  options: VersionFillOptions,
  strategy: VersionSizeStrategy,
): Promise<VersionFillResult> {
  if (strategy.kind === 'inline' || strategy.kind === 'none') {
    return { strategy: strategy.kind, skipped: 0, unmeasured: 0 };
  }
  if (strategy.kind === 'items-side-channel') {
    return fillFromSideChannel(client, siteUrl, library, items, options);
  }
  if (strategy.kind === 'render-list-data') {
    return fillFromRenderListData(client, siteUrl, library, items, options);
  }
  return fillPerFile(client, siteUrl, items, options);
}

// Cross-checks a bulk strategy's OUTPUT against what the version labels
// already told us, and reports whether the numbers are self-consistent.
//
// This exists because a probe can only prove a field is PRESENT and numeric —
// it cannot prove the field means what we assume. A field that exists in the
// list schema but is never populated returns 0, which sails through any
// presence check and then produces max(0, 0 - fileSize) = 0 for every single
// file: a total of 0 B that looks like a successful bulk sweep. That happened
// on a real tenant with RenderListDataAsStream.
//
// The check is free, and it is the one contradiction the data can actually
// expose: OData__UIVersionString (from the main sweep) independently says which
// files have retained versions. If it says a population of files DOES have
// them and the bulk field says every one of those is zero bytes, the bulk field
// is not measuring version storage and must not be trusted.
//
// Note what this deliberately does NOT flag: a library where every file is at
// version 1.0. There, `expected` is 0, there is no contradiction, and 0 B is
// the correct answer rather than a failure.
function versionTotalsLookCredible(
  library: LibraryInfo,
  items: FlatItem[],
  strategyLabel: string,
): boolean {
  // Files the labels say should have retained version bytes.
  const expected = items.filter((i) => !i.isFolder && !i.versionsProvablyZero);
  if (expected.length === 0) {
    // eslint-disable-next-line no-console
    console.info(
      `[SmartStorageAnalyzer] ${library.title}: every file is at version 1.0 — no retained ` +
      'versions exist, so a version-history total of 0 B is correct for this library.',
    );
    return true;
  }
  const nonZero = expected.filter((i) => (i.versionSizeBytes ?? 0) > 0).length;
  if (nonZero > 0) {
    // eslint-disable-next-line no-console
    console.info(
      `[SmartStorageAnalyzer] ${library.title}: ${strategyLabel} measured version history for ` +
      `${nonZero.toLocaleString()} of ${expected.length.toLocaleString()} files whose version ` +
      'label indicates retained versions.',
    );
    return true;
  }
  // eslint-disable-next-line no-console
  console.warn(
    `[SmartStorageAnalyzer] ${library.title}: ${strategyLabel} returned zero version bytes for ` +
    `all ${expected.length.toLocaleString()} files that DO have retained versions according to ` +
    'their version number. The field exists but is not populated with version storage on this ' +
    'list, so it is being discarded and version history measured per file instead.',
  );
  return false;
}

// Applies a { itemId -> total bytes including versions } map onto the rows,
// shared by both bulk strategies so they can't drift on the details that
// matter: joining by Id, never overwriting a proven zero, and clamping a
// lagging total that reads below the current file size.
function applyTotalsById(items: FlatItem[], totalById: Map<number, number>): number {
  let unmeasured = 0;
  for (const item of items) {
    if (item.isFolder) continue;
    // Already proven zero from the version label — leave it alone rather than
    // letting a lagging metrics figure overwrite a certain value.
    if (item.versionsProvablyZero) continue;
    const total = item.id != null ? totalById.get(item.id) : undefined;
    if (total == null) {
      // In the main sweep but not this one (created/deleted in between, or an
      // id that didn't parse). "Not measured" — never zero.
      unmeasured++;
      continue;
    }
    // SMTotalFileStreamSize is maintained by a lagging background job and can
    // read smaller than the current file size; that must not become a negative
    // version total.
    item.versionSizeBytes = Math.max(0, total - item.sizeBytes);
  }
  return unmeasured;
}

// ── Bulk: side-channel sweep ──────────────────────────────────────────────

async function fillFromSideChannel(
  client: SpApiClient,
  siteUrl: string,
  library: LibraryInfo,
  items: FlatItem[],
  options: VersionFillOptions,
): Promise<VersionFillResult> {
  const url = `${siteUrl}/_api/web/lists(guid'${library.id}')/items`
    + `?$select=${SIDE_CHANNEL_FIELDS}&$top=${ITEMS_PAGE_SIZE}`;

  let rows: any[];
  try {
    // skipBatch for the same reason the main sweep does it: a 5,000-row page
    // is the wrong shape for $batch, which exists to amortize many SMALL
    // requests (see getJson's docstring).
    rows = await client.getJsonPaged(url, options.signal, 400, true);
  } catch (err: any) {
    // The probe succeeded and the sweep then failed — throttling, most
    // likely. Report the library's version history as unmeasured rather than
    // falling through to ~194,000 per-file requests, which is exactly the
    // outcome this module exists to avoid. Not fatal: file sizes are intact.
    // eslint-disable-next-line no-console
    console.warn(
      `[SmartStorageAnalyzer] ${library.title}: version-history side-channel sweep failed — ` +
      `${err?.message ?? String(err)}. Version history is not included for this library.`,
    );
    const files = items.filter((i) => !i.isFolder && !i.versionsProvablyZero);
    return { strategy: 'none', skipped: 0, unmeasured: files.length };
  }

  // Join by Id, never by path: FileRef can differ in encoding between two
  // projections of the same list, and a file moved between the two sweeps
  // changes its path but keeps its id.
  const totalById = new Map<number, number>();
  for (const row of rows) {
    const id = Number(row.Id);
    if (!isFinite(id) || row.SMTotalFileStreamSize == null) continue;
    totalById.set(id, num(row.SMTotalFileStreamSize));
  }

  const unmeasured = applyTotalsById(items, totalById);
  // Same output validation as the RLDAS path — a present-but-unpopulated field
  // fails identically whichever projector returned it.
  if (!versionTotalsLookCredible(library, items, 'the un-expanded items sweep')) {
    clearBulkVersionSizes(items);
    return fillPerFile(client, siteUrl, items, options);
  }

  return { strategy: 'items-side-channel', skipped: 0, unmeasured };
}

// ── Bulk: RenderListDataAsStream (CAML/view renderer) ─────────────────────

async function fillFromRenderListData(
  client: SpApiClient,
  siteUrl: string,
  library: LibraryInfo,
  items: FlatItem[],
  options: VersionFillOptions,
): Promise<VersionFillResult> {
  const totalById = new Map<number, number>();
  let nextHref: string | undefined;
  try {
    for (let page = 0; page < RLDAS_MAX_PAGES; page++) {
      if (options.signal?.aborted) break;
      const data = await client.postJson(
        rldasUrl(siteUrl, library.id!, nextHref),
        rldasBody(ITEMS_PAGE_SIZE),
      );
      const rows: any[] = Array.isArray(data?.Row) ? data.Row : [];
      for (const row of rows) {
        const id = Number(row.ID);
        if (!isFinite(id)) continue;
        const total = parseDisplayInteger(row[VERSION_FIELD_CAML]);
        if (total === undefined) continue;
        totalById.set(id, total);
      }
      // RLDAS signals "more rows" by handing back a NextHref to echo. Absent
      // (or unchanged) means done — the unchanged check is a loop guard, since
      // a repeated token would otherwise re-fetch the same page forever.
      const href = typeof data?.NextHref === 'string' ? data.NextHref : undefined;
      if (!href || href === nextHref) break;
      nextHref = href;
    }
  } catch (err: any) {
    // The probe worked and the sweep then failed — throttling, most likely.
    // Report the library's version history as unmeasured rather than silently
    // dropping into ~194,000 per-file requests, which is the exact outcome
    // this module exists to prevent. File sizes themselves are unaffected.
    // eslint-disable-next-line no-console
    console.warn(
      `[SmartStorageAnalyzer] ${library.title}: RenderListDataAsStream sweep failed — ` +
      `${err?.message ?? String(err)}. Version history is not included for this library.`,
    );
    const files = items.filter((i) => !i.isFolder && !i.versionsProvablyZero);
    return { strategy: 'none', skipped: 0, unmeasured: files.length };
  }

  const unmeasured = applyTotalsById(items, totalById);
  // A successful sweep is not the same as a MEANINGFUL one — see
  // versionTotalsLookCredible. If the field turns out not to carry version
  // storage on this list, discard it and pay for per-file measurement rather
  // than reporting a confident 0 B.
  if (!versionTotalsLookCredible(library, items, 'RenderListDataAsStream')) {
    clearBulkVersionSizes(items);
    return fillPerFile(client, siteUrl, items, options);
  }

  return { strategy: 'render-list-data', skipped: 0, unmeasured };
}

// Undoes a rejected bulk strategy's writes before falling back, so the per-file
// pass sees "not measured" rather than a stale 0 it would then skip over.
// Proven zeros (from the version label) are left alone — those are real.
function clearBulkVersionSizes(items: FlatItem[]): void {
  for (const item of items) {
    if (item.isFolder || item.versionsProvablyZero) continue;
    item.versionSizeBytes = undefined;
  }
}

// ── Last resort: one request per file ─────────────────────────────────────

// Sums SP.FileVersion.Size across a file's retained versions, and counts them.
// The /Versions collection excludes the current version, so it IS exactly the
// retained history and the count is free from the same call.
//
// $top is set explicitly so a heavily-versioned file resolves in one request
// rather than paging; the default page size would make some files cost two.
async function fetchVersionInfo(
  client: SpApiClient,
  siteUrl: string,
  fileServerRelativeUrl: string,
  signal: AbortSignal | undefined,
): Promise<{ sizeBytes: number; count: number }> {
  const url = `${siteUrl}/_api/web/GetFileByServerRelativePath(decodedUrl='${encodeURIComponent(
    odata(fileServerRelativeUrl),
  )}')/Versions?$select=Size&$top=1000`;
  // Passing the signal only helps the multi-page case — getJsonPagedMeta
  // checks it BETWEEN pages, and with $top=1000 above almost every file is a
  // single page. The responsive part of cancellation is the per-task check in
  // the worker loop below.
  //
  // 'small' is the whole reason this path is tolerable at all: these responses
  // are a few numbers each, so they coalesce 50-to-an-envelope instead of 20,
  // cutting round trips by 60% on exactly the libraries that need thousands of
  // them (see BATCH_MAX_SMALL in spCore.ts).
  const versions = await client.getJsonPaged(url, signal, 400, false, undefined, 'small');
  return {
    sizeBytes: versions.reduce((sum, v) => sum + num(v.Size), 0),
    count: versions.length,
  };
}

async function fillPerFile(
  client: SpApiClient,
  siteUrl: string,
  items: FlatItem[],
  options: VersionFillOptions,
): Promise<VersionFillResult> {
  // Files whose version label already proves zero are skipped outright. This
  // is the single largest reduction available on the per-file path and it
  // costs nothing: the label came free in the main sweep, and in a typical
  // library the large majority of files sit at version 1.0.
  const candidates = items.filter((i) => !i.isFolder && !i.versionsProvablyZero);

  let unmeasured = 0;
  let targets = candidates;
  if (options.mode === 'quick' && candidates.length > QUICK_VERSION_FILE_LIMIT) {
    // Spend the budget on the LARGEST files. Retained-version bytes correlate
    // strongly with current file size, so the biggest N files hold most of
    // the version storage; capping an arbitrary N would bound the cost just
    // as well but leave the answer close to meaningless.
    //
    // Copy before sorting — `items` is the caller's array and its order is
    // load-bearing elsewhere (folder aggregation walks it as-is).
    targets = candidates.slice().sort((a, b) => b.sizeBytes - a.sizeBytes).slice(0, QUICK_VERSION_FILE_LIMIT);
    unmeasured = candidates.length - targets.length;
  }
  if (targets.length === 0) {
    return { strategy: 'per-file', skipped: 0, unmeasured };
  }

  const total = targets.length;
  let done = 0;
  let skipped = 0;
  let lastEmitAt = 0;
  let lastEmitDone = 0;
  const emit = (force: boolean): void => {
    const now = Date.now();
    if (!force && done - lastEmitDone < PROGRESS_EVERY_N && now - lastEmitAt < PROGRESS_EVERY_MS) return;
    lastEmitAt = now;
    lastEmitDone = done;
    options.onProgress?.(done, total);
  };
  emit(true);

  // How many workers run here is NOT a throttling decision, and treating it
  // like one is a mistake worth spelling out, because it was made here.
  //
  // This used to be floor(scanConcurrency / 2) — 3 workers on a default
  // setup — by analogy with the old pre-batching version queue, where each
  // worker really did mean one HTTP request. It doesn't any more: these calls
  // are coalesced into $batch envelopes of up to SMALL_BATCH_SIZE. An envelope
  // only fills from requests pending *at the same moment*, so 3 workers
  // produced 3-member batches and the 50-member ceiling did nothing at all.
  // Measured on a real tenant: ~3.5 files/second, i.e. ~15 hours for a
  // 194,000-file library.
  //
  // So workers are sized to KEEP THE ENVELOPES FULL, several deep, and the
  // actual request rate stays governed where it always was — acquireSlot in
  // spCore.ts, which bounds HTTP requests in flight (one per envelope, not one
  // per file) and is what the throttle backoff already shrinks. A worker
  // waiting on a batch holds no slot, so oversubscribing here cannot deadlock
  // and cannot outrun the governor; it only stops the governor from being fed
  // one request at a time.
  const VERSION_PENDING_TARGET = SMALL_BATCH_SIZE * 3;
  const concurrency = Math.min(targets.length, Math.max(client.scanConcurrency, VERSION_PENDING_TARGET));
  // Index-based workers rather than an array of closures: Full mode can have
  // ~194,000 targets, and materializing a closure per file (which is what
  // runConcurrent's signature requires) is pure waste at that size.
  let next = 0;
  const worker = async (): Promise<void> => {
    for (;;) {
      if (options.signal?.aborted) return;
      const i = next++;
      if (i >= targets.length) return;
      const file = targets[i];
      try {
        const info = await fetchVersionInfo(client, siteUrl, file.fileRef, options.signal);
        file.versionSizeBytes = info.sizeBytes;
        // An exact count from /Versions beats the label-derived estimate,
        // which can overstate on a library with a version-retention limit.
        file.versionCountApprox = info.count;
      } catch {
        skipped++;
        options.onSkipped?.();
      }
      done++;
      emit(false);
    }
  };
  await Promise.all(Array.from({ length: Math.min(concurrency, targets.length) }, worker));

  // A cancel mid-pass leaves the remaining targets genuinely unmeasured —
  // count them so the report says so rather than implying they were zero.
  if (options.signal?.aborted) unmeasured += Math.max(0, total - done);
  emit(true);

  return { strategy: 'per-file', skipped, unmeasured };
}
