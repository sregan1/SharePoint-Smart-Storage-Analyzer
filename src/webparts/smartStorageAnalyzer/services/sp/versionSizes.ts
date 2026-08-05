import { SpApiClient, valueArray, odata, SMALL_BATCH_SIZE } from './spCore';
import { LibraryInfo } from '../../models/models';
import { FlatItem, ITEMS_PAGE_SIZE } from './listItems';

// Getting version-history SIZE for a library whose bulk field doesn't work.
//
// The happy path lives in listItems.ts: SMTotalFileStreamSize rides along in
// the main item sweep and version size costs nothing at all. This module is
// everything that happens when that field is rejected.
//
// WHAT IS ACTUALLY TRUE ON A REAL TENANT (measured, not theorised — an earlier
// version of this comment asserted a theory that turned out to be wrong, and
// that wrong theory shaped the code for several rounds):
//   - `$select=…,SMTotalFileStreamSize` 400s with "The field or property
//     'SMTotalFileStreamSize' does not exist." — WITH and WITHOUT $expand=File.
//     So it is not an $expand collision, as was first assumed; the column is
//     simply not projectable through OData on these lists.
//   - RenderListDataAsStream DOES return the column, but every row reads 0 —
//     present in the schema, never populated. A probe that only checks
//     "parseable and non-zero somewhere" cannot tell that apart from real data,
//     and adopting it produced a confident 0 B for a whole 184,859-file library.
//   - Per-file /Versions works and is exact. With files at version 1.0 skipped
//     for free (see FlatItem.versionsProvablyZero), a 184,859-file library has
//     only 8,168 candidates, which completes quickly once requests are properly
//     coalesced into $batch envelopes.
//
// So: prove a BULK mechanism works before spending a single per-file request,
// and make the proof cheap AND decisive. Each mechanism is introduced the way
// this codebase introduces every uncertain SharePoint technique (see
// probeMaxIdAndFields and the canary shard in listItems.ts): probe small, adopt
// only on proof, degrade cleanly.

export type VersionSizeStrategyKind =
  // Nothing to do — the main sweep already returned SMTotalFileStreamSize.
  | 'inline'
  // The field IS selectable on its own, just not alongside $expand=File.
  // Sweep it separately and join by Id. ~1 request per 5,000 items.
  | 'items-side-channel'
  // OData won't project the field at all, but the CAML/view renderer will.
  // Also ~1 request per 5,000 items, just POSTs instead of GETs.
  | 'render-list-data'
  // No bulk mechanism works. One request per file — for EVERY file that could
  // have retained versions. There used to be a Quick mode capping this; see
  // ScanOptions.includeVersionHistory for why measurement retired it.
  | 'per-file'
  // Nothing worked. Report honestly as unmeasured, never as zero.
  | 'none';

export interface VersionSizeStrategy {
  kind: VersionSizeStrategyKind;
  // Short human phrase for WHY this kind was chosen, e.g. 'OData will not
  // project the field; the view renderer will'. Surfaced in the progress detail
  // and stored on the summary, so the report explains itself without the user
  // reading the console.
  reason: string;
}

export type VersionStage = 'probe' | 'bulk-sweep' | 'validate' | 'per-file';

export interface VersionProgressEvent {
  stage: VersionStage;
  // Always present. A stage with nothing to say is a stage that looks frozen —
  // which is exactly what happened when the bulk paths reported nothing.
  detail: string;
  done?: number;
  // undefined ⇒ indeterminate. Never clamp to make a bar look tidy.
  total?: number;
  unit?: 'pages' | 'files';
  skipped?: number;
}

export interface VersionFillOptions {
  signal?: AbortSignal;
  // Fired by EVERY stage, not just per-file.
  //
  // This used to be documented as per-file-only because "the bulk paths resolve
  // a whole library in a handful of requests and have nothing meaningful to
  // report in between". That was the entire bug: on a 184,859-item library a
  // bulk path is ~37 sequential heavyweight CAML POSTs, and the UI went silent
  // for all of them — bar pegged at 100%, "~0s remaining", for eleven minutes.
  //
  // Already coalesced at the source for the per-file path (PROGRESS_EVERY_N /
  // PROGRESS_EVERY_MS), so callers can wire this straight to setState.
  onProgress?: (e: VersionProgressEvent) => void;
  // One file's version fetch failed (throttling exhausted, transient error).
  // Kept in scope rather than failing the library, same as a skipped folder.
  onSkipped?: () => void;
}

export interface VersionFillResult {
  strategy: VersionSizeStrategyKind;
  // Attempted and failed. Retrying (or lowering concurrency) could fix these.
  skipped: number;
  // NEVER ATTEMPTED. Three causes, none of which retrying fixes: the scan was
  // canceled mid-pass; no mechanism worked for the list at all (strategy
  // 'none'); or the row was in the main item sweep but absent from the bulk
  // version sweep. Deliberately separate from `skipped` — that one IS worth
  // retrying, and conflating them tells the user to change a setting that
  // cannot help.
  unmeasured: number;
}

// ── Progress throttling ───────────────────────────────────────────────────
// The per-file path can complete thousands of requests. Emitting progress per
// completion would put a setState storm on a thread that is simultaneously
// holding a ~194,000-element array — the same reasoning that made the Storage
// Report's own file counter a ref flushed on a timer rather than per-file
// state (see scannedRef in StorageReportView.tsx).
const PROGRESS_EVERY_N = 50;
// Deliberately NOT backed by a timer of its own. During a throttle window
// nothing completes, so no timer here could fire either — it is the VIEW's own
// 500ms ticker plus SpApiClient.throttleRemainingSeconds that keeps the UI alive
// and explains the pause. Adding a second timer in the data layer would produce
// identical repeated events and still not say why nothing is moving.
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

function rldasBody(rowLimit: number, restrictToIds?: number[]): unknown {
  // Restricting to specific ids is what makes the probe DECISIVE — see
  // pickProbeSample. ID is always indexed, so an In filter is exempt from the
  // list view threshold on exactly the grounds listItems.ts already documents
  // for its `Id gt/le` shard filters.
  const where = restrictToIds?.length
    ? `<Query><Where><In><FieldRef Name='ID'/><Values>`
      + restrictToIds.map((id) => `<Value Type='Counter'>${id}</Value>`).join('')
      + `</Values></In></Where></Query>`
    : '';
  return {
    parameters: {
      // Scope='RecursiveAll' is MANDATORY and the most dangerous thing to omit
      // here. Without it the view is folder-scoped and returns only the root
      // folder's rows — which looks like a perfectly successful sweep while
      // silently missing almost the entire library. That failure mode would be
      // indistinguishable from "this library has little version history".
      ViewXml:
        `<View Scope='RecursiveAll'>`
        + where
        + `<ViewFields><FieldRef Name='ID'/><FieldRef Name='${VERSION_FIELD_CAML}'/></ViewFields>`
        + `<RowLimit Paged='TRUE'>${rowLimit}</RowLimit>`
        + `</View>`,
      DatesInUtc: true,
    },
  };
}

interface ProbeSample { id: number; sizeBytes: number; }

// Rows the probe can be FALSIFIED against: files whose version label says they
// DO have retained versions, so a populated SMTotalFileStreamSize must read
// strictly greater than the current file size.
//
// Largest first, deliberately. A big file with versions has the biggest absolute
// margin over its current size, which is what keeps a lagging-but-real metrics
// value from being mistaken for an unpopulated column.
//
// Sampling the FIRST rows of the library instead (what the probe used to do)
// does not work: those are overwhelmingly folders and ancient single-version
// files, so the intersection with "files that should have versions" is usually
// empty and the probe learns nothing.
//
// Copies before sorting — `items` is the caller's array and folder aggregation
// walks it in order.
function pickProbeSample(items: FlatItem[], want: number): ProbeSample[] {
  return items
    .filter((i) => !i.isFolder && !i.versionsProvablyZero && i.id != null && i.sizeBytes > 0)
    .slice()
    .sort((a, b) => b.sizeBytes - a.sizeBytes)
    .slice(0, want)
    .map((i) => ({ id: i.id!, sizeBytes: i.sizeBytes }));
}

// Verdict on a sampled set of (known size, reported total) pairs.
//
// Asymmetric ON PURPOSE. A wrong REJECT costs a slow-but-correct per-file pass.
// A wrong ACCEPT produces a confident 0 B for an entire library — the failure
// this module exists to prevent, and one a presence-only check already let
// through on a real tenant.
//
//   value > sizeBytes        -> proof the column carries version storage
//   value === 0              -> proof it is present but unpopulated
//   0 < value <= sizeBytes   -> inconclusive for one row, damning across a
//                               largest-first sample: the metrics job lags, but
//                               not for the N biggest versioned files at once.
function sampleProvesVersionStorage(pairs: { sizeBytes: number; value: number }[]): boolean {
  if (pairs.length === 0) return false;
  const proven = pairs.filter((p) => p.value > p.sizeBytes).length;
  return proven >= Math.max(1, Math.ceil(0.10 * pairs.length));
}

/**
 * Decides how to get version size for one library, given that the main sweep
 * did NOT return it inline.
 *
 * Costs at most two small requests, and is DECISIVE — it never adopts a
 * mechanism that a ~37-request sweep would then have to discard. That mattered:
 * before, this accepted RLDAS on "the column parses and something is non-zero",
 * swept 184,859 items across ~37 CAML POSTs, and threw all of it away.
 *
 * `items` comes from the main sweep, which has already finished by the time this
 * runs — so known file sizes are available to falsify a bulk field against, and
 * that is what makes the decision cheap AND safe.
 *
 * Never throws: an unusable library resolves to a strategy, not an error,
 * because a report that omits version history is still a useful report.
 */
export async function probeVersionSizeStrategy(
  client: SpApiClient,
  siteUrl: string,
  library: LibraryInfo,
  items: FlatItem[],
  options: VersionFillOptions,
): Promise<VersionSizeStrategy> {
  if (options.signal?.aborted || !library.id) {
    return { kind: 'per-file', reason: 'canceled or no list id' };
  }

  const sample = pickProbeSample(items, RLDAS_PROBE_ROWS);
  // No file in the library can have retained versions, so there is nothing to
  // measure and nothing any mechanism could prove. Take the free answer.
  if (sample.length === 0) {
    return {
      kind: 'items-side-channel',
      reason: 'no file in this library has retained versions',
    };
  }
  const sizeById = new Map(sample.map((s) => [s.id, s.sizeBytes]));

  // ── P1: the field on its own — no $expand, no $orderby, no $filter ───────
  // Kept as a presence-only test, deliberately. It costs one request; on this
  // tenant it fails outright with "The field or property does not exist", and if
  // it ever DOES succeed the sweep costs the same ~37 requests as the main one
  // with versionTotalsLookCredible as the safety net. Falsifying it properly
  // would need a second request to buy very little.
  options.onProgress?.({
    stage: 'probe',
    detail: 'checking whether the version field selects on its own',
  });
  try {
    const url = `${siteUrl}/_api/web/lists(guid'${library.id}')/items`
      + `?$select=${SIDE_CHANNEL_FIELDS}&$top=1`;
    const rows = valueArray(await client.getJson(url, true));
    // Presence, not truthiness: a genuinely zero-version file legitimately
    // reports 0, and `0` must not be read as "field missing".
    if (rows.length > 0 && 'SMTotalFileStreamSize' in rows[0] && rows[0].SMTotalFileStreamSize != null) {
      const reason = 'the field selects on its own, just not alongside $expand=File';
      // eslint-disable-next-line no-console
      console.info(`[SmartStorageAnalyzer] ${library.title}: ${reason} — using the bulk side channel.`);
      return { kind: 'items-side-channel', reason };
    }
  } catch {
    // Fall through. fetchLibraryItems has already logged the full 400 body for
    // this library, so a second message would just be noise.
  }

  // ── P2: RenderListDataAsStream, falsified against known sizes ────────────
  // RLDAS projects through the view/CAML renderer rather than the OData
  // projector, so it can return columns $select rejects. On this tenant it DOES
  // return the column — and every row reads 0, because the column exists in the
  // schema and nothing populates it.
  //
  // So the probe asks for exactly the rows that can disprove that: the largest
  // files whose version labels say they HAVE retained versions. For those,
  // SMTotalFileStreamSize must exceed the current file size. An all-zero (or
  // merely equal) answer is proof of an unpopulated column, not of a library
  // without version history.
  options.onProgress?.({
    stage: 'probe',
    detail: `checking the view renderer against ${sample.length} known-versioned files`,
  });
  try {
    const ids = sample.map((s) => s.id);
    let rows = await rldasProbeRows(client, siteUrl, library.id, ids);
    // Canary, per this file's own discipline: the ID-restricted CAML `In` filter
    // is a technique this codebase hasn't used before. If it errors or matches
    // nothing, fall back to an unrestricted sample rather than concluding
    // anything from its silence.
    if (rows === undefined || rows.length === 0) {
      rows = await rldasProbeRows(client, siteUrl, library.id, undefined);
    }
    if (rows && rows.length > 0) {
      const pairs = rows
        .map((r) => ({
          sizeBytes: sizeById.get(Number(r.ID)),
          value: parseDisplayInteger(r[VERSION_FIELD_CAML]),
        }))
        // Only rows we know a size for can falsify anything. An unrestricted
        // fallback sample will mostly drop out here, which is correct — it is
        // weaker evidence and should behave that way.
        .filter((p): p is { sizeBytes: number; value: number } =>
          p.sizeBytes !== undefined && p.value !== undefined);

      if (sampleProvesVersionStorage(pairs)) {
        const reason = 'OData will not project the field; the view renderer will';
        // eslint-disable-next-line no-console
        console.info(`[SmartStorageAnalyzer] ${library.title}: ${reason} — using the bulk CAML sweep.`);
        return { kind: 'render-list-data', reason };
      }
      // eslint-disable-next-line no-console
      console.info(
        `[SmartStorageAnalyzer] ${library.title}: the view renderer returns ${VERSION_FIELD_CAML} ` +
        `but it is not populated (checked ${pairs.length} files that DO have retained versions; ` +
        'none reported more than its current size). Measuring per file instead — which is correct, ' +
        'just slower. No bulk sweep was spent finding this out.',
      );
    }
  } catch (err: any) {
    // eslint-disable-next-line no-console
    console.warn(
      `[SmartStorageAnalyzer] ${library.title}: RenderListDataAsStream probe failed — ` +
      `${err?.message ?? String(err)}. Measuring per file instead.`,
    );
  }

  return {
    kind: 'per-file',
    reason: 'no bulk mechanism reports version storage on this list',
  };
}

// One RLDAS probe request. Returns undefined (not an empty array) when the
// request itself failed, so the caller can tell "the technique is unavailable"
// from "the technique worked and matched nothing".
async function rldasProbeRows(
  client: SpApiClient,
  siteUrl: string,
  listId: string,
  restrictToIds: number[] | undefined,
): Promise<any[] | undefined> {
  try {
    const data = await client.postJson(
      rldasUrl(siteUrl, listId),
      rldasBody(RLDAS_PROBE_ROWS, restrictToIds),
    );
    return Array.isArray(data?.Row) ? data.Row : [];
  } catch {
    return undefined;
  }
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
// KEEP THIS even though probeVersionSizeStrategy is now decisive. The probe
// samples; this sees the whole library. A lucky sample — 100 files whose metrics
// happened to be populated on a list where most are not — would pass the probe
// and produce a badly under-counted total, and this is the only thing that would
// catch it. It costs nothing: no requests, one pass over rows already in memory.
//
// The contradiction it exposes: OData__UIVersionString (from the main sweep)
// independently says which files have retained versions. If it says a population
// of files DOES have them and the bulk field says every one of those is zero
// bytes, the bulk field is not measuring version storage.
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

  const pagesHint = library.itemCount && library.itemCount > 0
    ? Math.max(1, Math.ceil(library.itemCount / ITEMS_PAGE_SIZE))
    : undefined;

  let rows: any[];
  try {
    // skipBatch for the same reason the main sweep does it: a 5,000-row page
    // is the wrong shape for $batch, which exists to amortize many SMALL
    // requests (see getJson's docstring).
    //
    // The onPage callback is the fix for a silent window: this argument used to
    // be omitted, so a ~37-request sweep of a large library reported absolutely
    // nothing while it ran.
    rows = await client.getJsonPaged(url, options.signal, 400, true, (_n, _rows, pagesDone) => {
      options.onProgress?.({
        stage: 'bulk-sweep',
        detail: pagesHint
          ? `reading version metadata — page ${pagesDone} of ~${pagesHint}`
          : `reading version metadata — page ${pagesDone}`,
        done: pagesDone,
        total: pagesHint != null && pagesDone <= pagesHint ? pagesHint : undefined,
        unit: 'pages',
      });
    });
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
  if (!await validateOrFallBack(library, items, 'the un-expanded items sweep', options)) {
    return fillPerFile(client, siteUrl, items, options);
  }

  return { strategy: 'items-side-channel', skipped: 0, unmeasured };
}

// Announces the validation stage, runs it, and rolls back a rejected bulk result
// so the per-file pass sees "not measured" rather than a stale zero.
//
// Returns true to keep the bulk result, false when the caller must fall back.
async function validateOrFallBack(
  library: LibraryInfo,
  items: FlatItem[],
  strategyLabel: string,
  options: VersionFillOptions,
): Promise<boolean> {
  options.onProgress?.({
    stage: 'validate',
    detail: 'checking the bulk version figures against version labels',
  });
  // Yield one macrotask so the label above actually paints.
  // versionTotalsLookCredible is synchronous over every row in the library
  // (~185,000 on a real one); without a yield, the emit that announces it
  // commits state the browser never gets a frame to render, and a rejected bulk
  // result would jump straight from "sweeping" to "measuring per file" with no
  // explanation of the pause in between.
  await new Promise((r) => setTimeout(r, 0));

  if (versionTotalsLookCredible(library, items, strategyLabel)) return true;
  clearBulkVersionSizes(items);
  return false;
}

// ── Bulk: RenderListDataAsStream (CAML/view renderer) ─────────────────────

async function fillFromRenderListData(
  client: SpApiClient,
  siteUrl: string,
  library: LibraryInfo,
  items: FlatItem[],
  options: VersionFillOptions,
): Promise<VersionFillResult> {
  const pagesHint = library.itemCount && library.itemCount > 0
    ? Math.max(1, Math.ceil(library.itemCount / ITEMS_PAGE_SIZE))
    : undefined;
  const totalById = new Map<number, number>();
  let nextHref: string | undefined;
  let rowsSoFar = 0;
  try {
    for (let page = 0; page < RLDAS_MAX_PAGES; page++) {
      if (options.signal?.aborted) break;
      // Emitted BEFORE the request, not only after it. Each of these is a
      // 5,000-row CAML render that can take seconds; reporting only on
      // completion leaves the slowest moment of each page unaccounted for.
      options.onProgress?.({
        stage: 'bulk-sweep',
        detail: pagesHint
          ? `reading version metadata — page ${page + 1} of ~${pagesHint}`
          : `reading version metadata — page ${page + 1}`,
        done: page,
        total: pagesHint != null && page < pagesHint ? pagesHint : undefined,
        unit: 'pages',
      });
      const data = await client.postJson(
        rldasUrl(siteUrl, library.id!, nextHref),
        rldasBody(ITEMS_PAGE_SIZE),
      );
      const rows: any[] = Array.isArray(data?.Row) ? data.Row : [];
      rowsSoFar += rows.length;
      for (const row of rows) {
        const id = Number(row.ID);
        if (!isFinite(id)) continue;
        const total = parseDisplayInteger(row[VERSION_FIELD_CAML]);
        if (total === undefined) continue;
        totalById.set(id, total);
      }
      options.onProgress?.({
        stage: 'bulk-sweep',
        detail: `read ${rowsSoFar.toLocaleString()} rows of version metadata`,
        done: page + 1,
        total: pagesHint != null && page + 1 <= pagesHint ? pagesHint : undefined,
        unit: 'pages',
      });
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
  if (!await validateOrFallBack(library, items, 'RenderListDataAsStream', options)) {
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
  // Two free eliminations, applied before a single request is issued:
  //
  //  - versionsProvablyZero: the version label already proves the file has
  //    nothing in its /Versions collection. This is the single largest reduction
  //    available anywhere on this path — on a real 184,859-file library it left
  //    only 8,168 candidates — and it costs nothing, because the label rode
  //    along in the main sweep.
  //  - sizeBytes === 0: a zero-byte file (a placeholder, a .url stub, an empty
  //    OneNote section) has no version bytes worth a request. Removes requests,
  //    never changes a total.
  const files = items.filter((i) => !i.isFolder);
  const targets = files.filter((i) => !i.versionsProvablyZero && i.sizeBytes > 0);
  const provablyZero = files.length - targets.length;

  // The ONLY writer of `unmeasured` on this path now. There used to be a Quick
  // mode that capped `targets` and counted the remainder here; every file that
  // can have retained versions is measured now, so the only way to leave one
  // unmeasured is to be canceled partway.
  let unmeasured = 0;
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
    options.onProgress?.({
      stage: 'per-file',
      detail: `${done.toLocaleString()} of ${total.toLocaleString()} files`
        + (skipped > 0 ? ` · ${skipped.toLocaleString()} failed` : ''),
      done,
      total,
      unit: 'files',
      skipped,
    });
  };
  // Announced before any request, and it answers the question the numbers
  // otherwise provoke — why a 184,859-file library only measures 8,168.
  options.onProgress?.({
    stage: 'per-file',
    detail: `${total.toLocaleString()} of ${files.length.toLocaleString()} files need measuring`
      + (provablyZero > 0 ? ` — ${provablyZero.toLocaleString()} have no retained versions` : ''),
    done: 0,
    total,
    unit: 'files',
  });

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
  // Index-based workers rather than an array of closures: this path is now
  // uncapped, so a pathological library could have ~190,000 targets, and
  // materializing a closure per file (which runConcurrent's signature requires)
  // is pure waste at that size.
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
