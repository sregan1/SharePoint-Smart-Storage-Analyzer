import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { clampConcurrency } from '../../utils/settingsBounds';

// Escape single-quotes in OData string literals (SQL-style doubling).
export function odata(s: string): string {
  return s.replace(/'/g, "''");
}

// API bases for a folder/file addressed by server-relative path. The
// *ByServerRelativePath(decodedUrl=...) forms plus URI-encoding handle names
// containing &, #, % and other characters that break the legacy
// GetFolderByServerRelativeUrl('...') form (& truncates the query string,
// # and % break URL parsing — items in such paths silently vanished).
export function folderApi(siteUrl: string, serverRelativeUrl: string): string {
  return `${siteUrl}/_api/web/GetFolderByServerRelativePath(decodedUrl='${encodeURIComponent(odata(serverRelativeUrl))}')`;
}

// Folder names SharePoint creates automatically that never hold user content.
// Shared by every recursive walk (fileWalk.ts, storageMetrics.ts) so the
// exclusion list can't drift between them.
export const SYSTEM_FOLDER_NAMES = new Set(['forms']);

// Single shared work queue with a global concurrency cap. Unlike nested
// runConcurrent pools (which multiply: N workers each spawning N more per
// recursion level), tasks here can enqueue follow-up tasks — e.g. recursive
// folder walks — while total in-flight work stays capped at `concurrency`.
export class TaskQueue {
  private active = 0;
  private pending: (() => Promise<void>)[] = [];
  private idleResolvers: (() => void)[] = [];

  constructor(private readonly concurrency: number) {}

  add(task: () => Promise<void>): void {
    this.pending.push(task);
    this.pump();
  }

  private pump(): void {
    while (this.active < this.concurrency && this.pending.length > 0) {
      const task = this.pending.shift()!;
      this.active++;
      const onDone = (): void => {
        this.active--;
        this.pump();
        if (this.active === 0 && this.pending.length === 0) {
          this.idleResolvers.splice(0).forEach((resolve) => resolve());
        }
      };
      // Individual task errors don't stop the queue.
      task().then(onDone, onDone);
    }
  }

  drain(): Promise<void> {
    if (this.active === 0 && this.pending.length === 0) return Promise.resolve();
    return new Promise((resolve) => this.idleResolvers.push(resolve));
  }
}

// Normalise top-level value arrays: SPO REST returns a direct array with
// odata=nometadata; legacy verbose mode wraps it in { results: [] } or { value: [] }.
export function valueArray(data: any): any[] {
  if (Array.isArray(data)) return data;
  if (Array.isArray(data?.value)) return data.value;
  if (Array.isArray(data?.results)) return data.results;
  return [];
}

// Known system/infrastructure library URL suffixes (lowercased, site-relative).
// Checked as a suffix so they match regardless of site path prefix.
export const SYSTEM_LIB_SUFFIXES = [
  '/formservertemplates', // Form Templates
  '/style library',       // Style Library
];

// Returns true if this list entry should be treated as a system/hidden library
// and excluded when includeHidden is false. NoCrawl is deliberately NOT
// treated as system: admins sometimes mark large archival libraries NoCrawl
// to hide them from search, which is exactly what a storage audit needs to
// see — those are included and flagged via LibraryInfo.noCrawl instead.
export function isSystemLibrary(lib: any): boolean {
  if (lib.IsSiteAssetsLibrary) return true;
  const url = ((lib.RootFolder?.ServerRelativeUrl) ?? '').toLowerCase();
  return SYSTEM_LIB_SUFFIXES.some((s) => url.endsWith(s));
}

// Base templates that behave like document libraries (have Files/Folders and
// can be walked): 101 = Document Library, 109 = Picture Library, 119 = Site
// Pages. Everything else is a generic list, not a storage-relevant target.
export const LIBRARY_TEMPLATES = [101, 109, 119];

// Identifies this app's traffic to SharePoint Online. Browsers forbid setting
// User-Agent from script (it's a Fetch-spec forbidden header), so the
// documented "decorate your traffic" guidance can't be followed in its usual
// form from a client-side web part — X-ClientService-ClientTag is the header
// SPO does accept from the browser for client identification. Treat this as
// diagnosability (it shows up in tenant request logs, which is what lets an
// admin or Microsoft support attribute throttling to this web part) rather
// than as a throttling exemption.
const CLIENT_TAG = 'SmartStorageAnalyzer/1.2.0';

// Which $batch envelope a coalesced request belongs to. Purely a size policy:
// 'small' promises a tiny, bounded response and so tolerates a bigger envelope
// (see BATCH_MAX_SMALL). Groups are also never mixed within one envelope, so a
// 5,000-row member can't delay 49 one-line ones.
export type BatchGroup = 'default' | 'small';

// How many 'small'-group requests ride in one $batch envelope.
//
// Exported because a caller cannot benefit from this without knowing it. A
// batch only fills if at least this many logical requests are pending at the
// same moment, and the number of pending requests is set by the CALLER's
// worker count — so a caller running 3 workers gets 3-member batches no matter
// how high this is. That mistake cost a real scan ~15x its throughput; see
// VERSION_PENDING_TARGET in versionSizes.ts for the other half of the contract.
export const SMALL_BATCH_SIZE = 50;

// Shared API client: SPFx context plus the throttling-aware fetch helpers and
// user-tunable scan settings. All sp/ modules take this as their first argument.
export class SpApiClient {
  public readonly context: WebPartContext;
  private _scanConcurrency = 6;

  // ── Global request governor ─────────────────────────────────────────────
  // Every request in the app funnels through getJson, so bounding things here
  // bounds ALL traffic regardless of how many TaskQueue/runConcurrent pools
  // happen to be alive at once. That matters because each pool previously
  // enforced only its OWN limit and nothing enforced a tenant-wide ceiling:
  // the Explorer alone ran a StorageMetrics probe pool (deliberately doubled
  // to 2x) with a scanConcurrency-sized fallback-walk pool nested inside it,
  // and a Storage Report scan ran its folder queue plus a version-history
  // queue — so a setting of "6" could mean 16+ requests in flight. On a large
  // tenant that is enough to stay permanently throttled.
  //
  // Slots are acquired around the individual HTTP call only, never held
  // across a task's lifetime. That's what keeps this deadlock-free: a
  // walkLibrary task awaiting its versionQueue children holds a TaskQueue
  // slot but zero request slots, so the children can always acquire one.
  private inFlight = 0;
  private slotWaiters: (() => void)[] = [];

  // While this timestamp is in the future, EVERY request waits — not just the
  // one that got throttled. The previous per-request-only retry left all
  // other workers hammering the same throttled tenant, which is what turns a
  // brief throttle into a sustained one.
  private throttledUntilMs = 0;
  private consecutiveThrottles = 0;

  /**
   * Max concurrent API requests during scans. Settable from Settings.
   * SharePoint Online doesn't publish a fixed throttling threshold (it's
   * dynamic and resource-based), so this is a starting point, not a
   * measured limit — paired with the escalating-but-fast retry backoff in
   * getJson() below so pushing this higher fails gracefully instead of
   * stalling on a flat 10s wait per throttled request.
   *
   * Clamped on write: TaskQueue.pump() only ever runs a task while
   * `active < concurrency`, so a NaN/0/negative value reaching it (a
   * corrupted localStorage value, or a caller bypassing the Settings UI's
   * own bounds) would never run anything and hang every scan/folder load
   * forever instead of just misbehaving.
   */
  public get scanConcurrency(): number { return this._scanConcurrency; }
  public set scanConcurrency(value: number) { this._scanConcurrency = clampConcurrency(value); }

  constructor(context: WebPartContext) {
    this.context = context;
  }

  // Retries on 429/503, honoring the server's Retry-After header when it
  // sends one. When it doesn't, fall back to an escalating schedule — most
  // throttling is transient and clears quickly, so the first retry should be
  // fast; the wait only grows if it keeps happening.
  //
  // This budget is deliberately generous. The previous one (3 attempts,
  // 2/5/10s, ~17s total) was far too small for a large tenant, where SPO's
  // own Retry-After is routinely 30-120s: a throttled scan didn't slow down
  // and recover, it GAVE UP, marking folders unreadable while the server was
  // merely asking it to wait. Waiting minutes and succeeding beats failing in
  // seconds, since the alternative is a report with holes in it.
  private static readonly RETRY_BACKOFF_SECONDS = [2, 5, 10, 20, 30, 45, 60, 60];
  // Cap on honoring a single Retry-After. Past this, failing that one folder
  // (and reporting it as unmeasurable) beats blocking the whole queue behind
  // one pathological wait.
  private static readonly MAX_RETRY_AFTER_SECONDS = 120;

  // Effective in-flight ceiling, reduced while throttling is active: halved
  // per consecutive throttle down to a floor of 1. SPO throttling is
  // resource-based and tenant-wide, so the only thing that actually clears it
  // is offering less load — retrying at the same rate prolongs it. Recovers a
  // step at a time as clean responses come back (see noteSuccess).
  private effectiveConcurrency(): number {
    const divisor = Math.pow(2, Math.min(this.consecutiveThrottles, 3));
    return Math.max(1, Math.floor(this._scanConcurrency / divisor));
  }

  private async acquireSlot(): Promise<void> {
    while (this.inFlight >= this.effectiveConcurrency()) {
      await new Promise<void>((resolve) => { this.slotWaiters.push(resolve); });
    }
    this.inFlight++;
  }

  private releaseSlot(): void {
    this.inFlight--;
    this.slotWaiters.shift()?.();
  }

  // Blocks until the global throttle window has passed. Sleeps in short
  // slices rather than one long timer so a window that gets shortened (or a
  // scan the user cancels) isn't stuck waiting on a stale deadline.
  private async waitOutThrottle(): Promise<void> {
    for (;;) {
      const remaining = this.throttledUntilMs - Date.now();
      if (remaining <= 0) return;
      await new Promise((r) => setTimeout(r, Math.min(remaining, 500)));
    }
  }

  private noteThrottled(waitSeconds: number): void {
    this.consecutiveThrottles++;
    // Extend, never shorten: a concurrent request that got a smaller
    // Retry-After must not pull the shared gate in and let everyone resume
    // early while the tenant is still shedding load.
    this.throttledUntilMs = Math.max(this.throttledUntilMs, Date.now() + waitSeconds * 1000);
    // eslint-disable-next-line no-console
    console.warn(
      `[SmartStorageAnalyzer] Throttled by SharePoint — pausing all requests for ${waitSeconds}s ` +
      `and reducing concurrency to ${this.effectiveConcurrency()} (was ${this._scanConcurrency}).`,
    );
  }

  private noteSuccess(): void {
    if (this.consecutiveThrottles > 0) {
      this.consecutiveThrottles--;
      // Waking waiters here matters: effectiveConcurrency just went UP, so
      // slots exist that nobody would otherwise be notified about until an
      // unrelated request happened to finish.
      this.slotWaiters.splice(0, this.slotWaiters.length).forEach((w) => w());
    }
  }

  /** True while the shared throttle gate is holding requests back. */
  public get isThrottled(): boolean { return this.throttledUntilMs > Date.now(); }

  // 406 IS A THROTTLING SIGNAL — do not "fix" this by treating it as a bad
  // request. When SPO decides to shed load it redirects the call to its HTML
  // throttle page (/_layouts/15/Throttle.htm); we asked for JSON, HTML is not
  // acceptable, so the browser surfaces 406 instead of 429. Console evidence:
  // hundreds of "GET /_layouts/15/Throttle.htm ... 406" entries interleaved
  // one-to-one with our own failures, on folders with entirely ordinary names
  // ("IP", "CSRs", "2019"). A genuinely malformed path yields 400/404, not 406.
  //
  // Getting this wrong is costly in both directions: it made every throttle
  // look like a permanent per-folder defect (so no backoff, no gate, instant
  // give-up), and it told users to stop adjusting the one setting that
  // actually helps.
  private static isThrottleResponse(status: number): boolean {
    return status === 429 || status === 503 || status === 406;
  }

  // ── Request coalescing via SharePoint's $batch endpoint ─────────────────
  // The recursive folder walks are the dominant source of request volume:
  // 2 requests per folder (Files + Folders), and the Explorer's fallback walk
  // recurses a whole subtree per subfolder whose StorageMetrics rollup reads 0
  // — which the rollup commonly does. Opening one folder with 30 subfolders of
  // ~200 folders each is ~12,000 requests to paint a single screen.
  //
  // Rather than restructure those walks (their recursion carries subtle
  // deadlock- and cache-poisoning invariants), coalescing lives here at the
  // request layer: callers keep calling getJson one URL at a time, and
  // whatever overlaps within a short window is bundled into one $batch POST.
  // The walks already run scanConcurrency folders at once, so overlap is the
  // normal case and grouping happens naturally — every call site benefits
  // (both walks, the StorageMetrics probes, and version-history lookups)
  // without any of them changing.
  private batchQueue: {
    url: string; group: BatchGroup; resolve: (v: any) => void; reject: (e: any) => void;
  }[] = [];
  private batchTimer: ReturnType<typeof setTimeout> | undefined;
  private batchSeq = 0;
  // SPO documents a higher per-batch ceiling, but large batches are slower to
  // fail and harder to attribute when one member misbehaves. 20 captures
  // nearly all of the win at a fraction of the blast radius.
  //
  // That reasoning holds for the DEFAULT group, whose members (folder
  // listings, StorageMetrics probes) return payloads of unbounded size. It
  // does not hold for requests whose responses are known to be tiny and
  // bounded — see BATCH_MAX_SMALL.
  private static readonly BATCH_MAX = 20;
  // A second, larger ceiling for call sites that can promise a small
  // response. The per-file version fetch (`/Versions?$select=Size&$top=1000`)
  // qualifies: a handful of numbers per file. On a library that has to measure
  // thousands of files individually, going 20 -> 50 per envelope removes 60%
  // of the round trips, which on a throttled tenant is the difference that
  // matters most.
  //
  // 50 rather than SPO's documented 100, deliberately: both batch failure
  // paths below (envelope failure, part-count mismatch) degrade a failed batch
  // into N INDIVIDUAL requests. At 100 that is a 100-request amplification
  // burst aimed at a tenant that just signalled it was unhappy. 50 keeps most
  // of the win at half the amplification, and can be raised later if this
  // tenant demonstrates it tolerates it.
  private static readonly BATCH_MAX_SMALL = SMALL_BATCH_SIZE;

  private static batchMaxFor(group: BatchGroup): number {
    return group === 'small' ? SpApiClient.BATCH_MAX_SMALL : SpApiClient.BATCH_MAX;
  }
  // Long enough for concurrent walk tasks to land in the same batch, short
  // enough to be invisible next to a SharePoint round trip.
  private static readonly BATCH_WINDOW_MS = 15;

  /** Escape hatch: set false to bypass $batch entirely (one request per call). */
  public batchingEnabled = true;
  private batchFallbackWarned = false;

  /**
   * `skipBatch` opts a single call out of coalescing. Batching exists to
   * amortize MANY SMALL requests; a bulk read that returns thousands of list
   * items is the opposite shape — bundling it buys nothing (there is nothing
   * to amortize against) and actively hurts: the whole multipart response
   * must be buffered and string-split before any member resolves, and one
   * oversized member delays every unrelated request that landed in the same
   * 15ms window. Bulk paged reads (see listItems.ts) pass true.
   */
  public async getJson(url: string, skipBatch = false, group: BatchGroup = 'default'): Promise<any> {
    if (!this.batchingEnabled || skipBatch) return this.getJsonDirect(url);
    return new Promise<any>((resolve, reject) => {
      this.batchQueue.push({ url, group, resolve, reject });
      // Count only this request's OWN group against its own ceiling. A plain
      // length check on the mixed queue would either flush a 'small' group
      // long before it filled (because unrelated 'default' requests pushed the
      // total over 20) or let it overshoot — both of which defeat the point of
      // having two ceilings at all.
      let queuedInGroup = 0;
      for (const item of this.batchQueue) if (item.group === group) queuedInGroup++;
      if (queuedInGroup >= SpApiClient.batchMaxFor(group)) this.flushBatches();
      else if (this.batchTimer === undefined) {
        this.batchTimer = setTimeout(() => this.flushBatches(), SpApiClient.BATCH_WINDOW_MS);
      }
    });
  }

  private flushBatches(): void {
    if (this.batchTimer !== undefined) {
      clearTimeout(this.batchTimer);
      this.batchTimer = undefined;
    }
    const queued = this.batchQueue.splice(0, this.batchQueue.length);
    if (queued.length === 0) return;

    // Partitioned by (site, group).
    //
    // Site, because a $batch POST goes to one web's endpoint and cannot span
    // site collections — a subsite-inclusive scan has requests for several
    // webs in flight at once and mixing them would fail the whole envelope.
    //
    // Group, because the two groups have different size ceilings, and because
    // keeping them apart is what stops one unbounded 'default' member from
    // holding up 49 tiny 'small' ones inside the same multipart response.
    const partitions = new Map<string, typeof queued>();
    for (const item of queued) {
      const site = item.url.split('/_api/')[0];
      const key = `${item.group} ${site}`;
      const bucket = partitions.get(key);
      if (bucket) bucket.push(item);
      else partitions.set(key, [item]);
    }
    partitions.forEach((items, key) => {
      const site = key.substring(key.indexOf(' ') + 1);
      const max = SpApiClient.batchMaxFor(items[0].group);
      for (let i = 0; i < items.length; i += max) {
        void this.sendBatch(site, items.slice(i, i + max));
      }
    });
  }

  private async sendBatch(siteUrl: string, group: typeof this.batchQueue): Promise<void> {
    // Not worth the multipart envelope for a lone request.
    if (group.length === 1) {
      try { group[0].resolve(await this.getJsonDirect(group[0].url)); }
      catch (err) { group[0].reject(err); }
      return;
    }

    const boundary = `batch_ssa_${++this.batchSeq}`;
    const body = group.map((item) => (
      `--${boundary}\r\n` +
      'Content-Type: application/http\r\n' +
      'Content-Transfer-Encoding: binary\r\n\r\n' +
      `GET ${item.url} HTTP/1.1\r\n` +
      'Accept: application/json;odata=nometadata\r\n\r\n'
    )).join('') + `--${boundary}--\r\n`;

    let parts: { status: number; body: string }[];
    try {
      parts = await this.postBatch(siteUrl, boundary, body);
    } catch (err: any) {
      // The batch envelope itself failed (or the endpoint is unavailable on
      // this tenant). Fall back to individual requests rather than failing
      // every caller in the group — batching is an optimization and must
      // never be the reason a folder reads as unreadable.
      //
      // Warned once, because a persistent envelope-level failure (a rejected
      // header, a tenant with $batch disabled) would otherwise degrade every
      // request to unbatched with no signal at all — the scan would just be
      // as slow as before for no apparent reason.
      //
      // The message used to claim this disabled batching "for the rest of this
      // session", which was never true — the fallback is per-group, and the
      // next group batches normally. That wording sent a real debugging
      // session looking for a permanent switch that doesn't exist.
      if (!this.batchFallbackWarned) {
        this.batchFallbackWarned = true;
        // eslint-disable-next-line no-console
        console.warn(
          `[SmartStorageAnalyzer] $batch envelope failed (${err?.message ?? String(err)}) — ` +
          'running this group as individual requests. Batching stays enabled for later groups; ' +
          'this warning appears only once.',
        );
      }
      await Promise.all(group.map(async (item) => {
        try { item.resolve(await this.getJsonDirect(item.url)); }
        catch (e) { item.reject(e); }
      }));
      return;
    }

    // Batch responses come back in request order. If the count doesn't line
    // up, the mapping is untrustworthy — resolving the wrong body against the
    // wrong caller would silently attribute one folder's contents to another,
    // so fall back rather than guess.
    if (parts.length !== group.length) {
      // eslint-disable-next-line no-console
      console.warn(
        `[SmartStorageAnalyzer] $batch returned ${parts.length} parts for ${group.length} requests — ` +
        'falling back to individual requests for this group.',
      );
      await Promise.all(group.map(async (item) => {
        try { item.resolve(await this.getJsonDirect(item.url)); }
        catch (err) { item.reject(err); }
      }));
      return;
    }

    await Promise.all(group.map(async (item, i) => {
      const part = parts[i];
      // A member throttled inside the batch is retried on its own, through
      // the normal governor/backoff path, so it gets the same treatment an
      // unbatched request would.
      if (SpApiClient.isThrottleResponse(part.status)) {
        try { item.resolve(await this.getJsonDirect(item.url)); }
        catch (err) { item.reject(err); }
        return;
      }
      if (part.status >= 400) {
        item.reject(new Error(`HTTP ${part.status} on ${item.url} — ${part.body.substring(0, 300)}`));
        return;
      }
      try { item.resolve(part.body ? JSON.parse(part.body) : {}); }
      catch (err) { item.reject(err); }
    }));
  }

  private async postBatch(
    siteUrl: string,
    boundary: string,
    body: string,
    attempt = 0,
  ): Promise<{ status: number; body: string }[]> {
    await this.waitOutThrottle();
    await this.acquireSlot();
    let resp;
    try {
      // jsonRequest/jsonResponse MUST both be off. The v1 configuration is
      // JSON-oriented, and a $batch payload is multipart/mixed both ways —
      // leaving jsonResponse on makes SPHttpClient reject the call outright
      // with "ISPHttpClientConfiguration.jsonResponse is enabled, which
      // requires the OData-Version header to be 3.0 or 4.0".
      //
      // Blanking `odata-version` by hand (the previous attempt at this) does
      // NOT work — it trips that same guard, so every batch failed and
      // sendBatch quietly fell back to unbatched requests. Because the
      // fallback is silent-by-design, the only symptom was that batching
      // never actually happened; the one-time warning below exists to make
      // exactly that visible.
      const batchConfig = SPHttpClient.configurations.v1.overrideWith({
        jsonRequest: false,
        jsonResponse: false,
      });
      resp = await this.context.spHttpClient.post(`${siteUrl}/_api/$batch`, batchConfig, {
        headers: {
          'Content-Type': `multipart/mixed; boundary=${boundary}`,
          'X-ClientService-ClientTag': CLIENT_TAG,
        },
        body,
      });
    } finally {
      this.releaseSlot();
    }

    // Throttling of the batch POST as a whole. RETRY THE ENVELOPE — do not
    // throw, because the caller's failure path fans a rejected batch out into
    // one individual request per member.
    //
    // That fan-out is right for a STRUCTURAL failure (a tenant with $batch
    // disabled, a rejected header) where the envelope will never work. It is
    // catastrophically wrong for a 429: it answers "you are sending too much"
    // by turning one request into up to BATCH_MAX_SMALL of them, against a
    // tenant that is already shedding load. Observed on a real tenant — one
    // throttled batch, then backoff escalating 10s -> 20s -> 44s as the
    // amplified individual requests kept arriving.
    //
    // The shared gate is tripped either way, so this retry waits out the same
    // window every other request does.
    if (SpApiClient.isThrottleResponse(resp.status)) {
      const fallback = SpApiClient.RETRY_BACKOFF_SECONDS[
        Math.min(attempt, SpApiClient.RETRY_BACKOFF_SECONDS.length - 1)
      ];
      const header = parseInt(resp.headers.get('Retry-After') ?? '', 10);
      this.noteThrottled(Math.min(
        isNaN(header) ? fallback : Math.max(header, 1),
        SpApiClient.MAX_RETRY_AFTER_SECONDS,
      ));
      if (attempt < SpApiClient.RETRY_BACKOFF_SECONDS.length) {
        return this.postBatch(siteUrl, boundary, body, attempt + 1);
      }
      // Budget exhausted. Now the caller's per-member fallback is the lesser
      // evil: each member gets its own retry budget and its own slot, which at
      // least makes progress instead of failing the whole group outright.
      throw new Error(`HTTP ${resp.status} on $batch`);
    }
    if (!resp.ok) throw new Error(`HTTP ${resp.status} on $batch`);
    this.noteSuccess();
    return SpApiClient.parseBatchResponse(await resp.text());
  }

  // Splits a multipart/mixed batch response into its inner HTTP responses.
  // The response boundary is generated by the server and differs from the
  // request's, so it's matched by shape rather than assumed.
  private static parseBatchResponse(text: string): { status: number; body: string }[] {
    const out: { status: number; body: string }[] = [];
    for (const chunk of text.split(/--batchresponse_[0-9a-zA-Z._-]+/)) {
      const status = /HTTP\/1\.1\s+(\d{3})/.exec(chunk);
      if (!status) continue;
      // Inner headers end at the first blank line after the status line; the
      // remainder is the body (JSON for everything this app requests).
      const headerEnd = chunk.indexOf('\r\n\r\n', status.index);
      const raw = headerEnd === -1 ? '' : chunk.substring(headerEnd + 4);
      out.push({ status: parseInt(status[1], 10), body: raw.trim() });
    }
    return out;
  }

  public async getJsonDirect(url: string, attempt = 0): Promise<any> {
    // Both gates, in this order: wait out any tenant-wide throttle window
    // first, then take a slot from the global in-flight ceiling.
    await this.waitOutThrottle();
    await this.acquireSlot();
    let resp;
    try {
      resp = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1, {
        headers: { 'X-ClientService-ClientTag': CLIENT_TAG },
      });
    } finally {
      // Released as soon as the response comes back — retry backoff waiting
      // must NOT hold a slot, or a throttled tenant would keep the ceiling
      // fully occupied by requests that are only sleeping.
      this.releaseSlot();
    }
    return this.classifyResponse(resp, url, attempt, () => this.getJsonDirect(url, attempt + 1));
  }

  /**
   * A JSON POST with exactly the same governor discipline as getJsonDirect:
   * throttle gate, in-flight slot released before any backoff wait, and the
   * shared 429/503/406/non-JSON classification below.
   *
   * Added for RenderListDataAsStream (see versionSizes.ts), which is a POST
   * because its ViewXml is a request body. Deliberately NOT built on
   * postBatch — that one is hard-wired to multipart/mixed and must keep
   * jsonRequest/jsonResponse off, which is the opposite of what's needed here.
   *
   * SPFx's SPHttpClient supplies X-RequestDigest itself, so there's no digest
   * plumbing to get wrong.
   */
  public async postJson(url: string, body: unknown, attempt = 0): Promise<any> {
    await this.waitOutThrottle();
    await this.acquireSlot();
    let resp;
    try {
      resp = await this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
        headers: {
          'X-ClientService-ClientTag': CLIENT_TAG,
          'Accept': 'application/json;odata=nometadata',
          'Content-Type': 'application/json;odata=nometadata',
        },
        body: JSON.stringify(body),
      });
    } finally {
      this.releaseSlot();
    }
    return this.classifyResponse(resp, url, attempt, () => this.postJson(url, body, attempt + 1));
  }

  // Shared response handling for getJsonDirect and postJson.
  //
  // Factored out rather than duplicated ON PURPOSE: this is where every
  // throttling rule in the app actually lives (406 being a throttle signal, a
  // 200-with-HTML being one too, tripping the shared gate even on the final
  // attempt). Two copies would drift, and a POST path that mishandled
  // throttling would poison the shared governor for every GET as well.
  private async classifyResponse(
    resp: SPHttpClientResponse,
    url: string,
    attempt: number,
    retry: () => Promise<any>,
  ): Promise<any> {
    if (SpApiClient.isThrottleResponse(resp.status)) {
      const fallback = SpApiClient.RETRY_BACKOFF_SECONDS[
        Math.min(attempt, SpApiClient.RETRY_BACKOFF_SECONDS.length - 1)
      ];
      // 406 (the throttle-page redirect) carries no Retry-After, so it always
      // falls through to the escalating schedule.
      const header = parseInt(resp.headers.get('Retry-After') ?? '', 10);
      const waitSeconds = Math.min(
        isNaN(header) ? fallback : Math.max(header, 1),
        SpApiClient.MAX_RETRY_AFTER_SECONDS,
      );
      // Trip the shared gate even on the final attempt: other in-flight work
      // still benefits from backing off, whether or not THIS request retries.
      this.noteThrottled(waitSeconds);
      if (attempt < SpApiClient.RETRY_BACKOFF_SECONDS.length) {
        // retry(), never getJsonDirect — otherwise a throttled POST would
        // quietly retry itself as a GET and lose its body.
        return retry();
      }
      throw new Error(
        `HTTP ${resp.status} on ${url} — SharePoint is throttling this account ` +
        `(still blocked after ${SpApiClient.RETRY_BACKOFF_SECONDS.length} attempts). ` +
        'Lower Concurrent API requests in Settings and try again shortly.',
      );
    }

    if (!resp.ok) {
      const txt = await resp.text();
      // Include the URL — 406/400-class failures are almost always specific
      // to one folder/file (a name with a problematic character, or an
      // over-long path), and the body SharePoint returns rarely says which;
      // the URL itself is usually the only clue pointing at the culprit.
      throw new Error(`HTTP ${resp.status} on ${url} — ${txt.substring(0, 300)}`);
    }

    // A 200 with an HTML body — an interstitial/error page, not the JSON
    // that was asked for — is a real, observed shape distinct from the
    // documented 429/503/406 throttle responses: under heavy sustained load
    // (many concurrent shard requests — see listItems.ts) SharePoint can
    // return one of these instead, with no throttle-signaling status code
    // to catch above. Calling resp.json() on it crashes with "Unexpected
    // token '<'" rather than surfacing anything actionable. Content-Type is
    // the reliable signal here (unlike sniffing the body), and treating it
    // exactly like a recognized throttle — trip the shared gate, retry with
    // the same escalating backoff — is the correct response either way:
    // whether this really is throttling in disguise or some other transient
    // interstitial, backing off and retrying is what recovers it.
    const contentType = resp.headers.get('Content-Type') ?? '';
    if (!contentType.includes('json')) {
      const fallback = SpApiClient.RETRY_BACKOFF_SECONDS[
        Math.min(attempt, SpApiClient.RETRY_BACKOFF_SECONDS.length - 1)
      ];
      this.noteThrottled(fallback);
      if (attempt < SpApiClient.RETRY_BACKOFF_SECONDS.length) {
        return retry();
      }
      throw new Error(
        `HTTP ${resp.status} on ${url} — SharePoint returned a non-JSON response ` +
        `(likely throttling) and retries were exhausted. Lower Concurrent API requests ` +
        'in Settings and try again shortly.',
      );
    }

    this.noteSuccess();
    return resp.json();
  }

  // Fetches a collection endpoint and follows server-side paging links so
  // results beyond the $top page size are not silently dropped, reporting
  // whether it had to give up before the last page. maxPages is a safety
  // valve against runaway loops on enormous collections — at $top=5000 the
  // default of 400 covers 2,000,000 items before it ever kicks in, so a real
  // truncation here means a genuinely enormous single folder, not routine
  // library size.
  public async getJsonPagedMeta(
    url: string,
    signal?: AbortSignal,
    maxPages = 400,
    skipBatch = false,
    onPage?: (fetchedSoFar: number) => void,
    group: BatchGroup = 'default',
  ): Promise<{ items: any[]; truncated: boolean }> {
    const all: any[] = [];
    let next: string | undefined = url;
    let page = 0;
    for (; next && page < maxPages && !signal?.aborted; page++) {
      const data = await this.getJson(next, skipBatch, group);
      all.push(...valueArray(data));
      onPage?.(all.length);
      next = data?.['odata.nextLink'] ?? data?.['@odata.nextLink'] ?? data?.d?.__next;
    }
    const truncated = !!next && page >= maxPages;
    if (truncated) {
      // eslint-disable-next-line no-console
      console.warn(`[SmartStorageAnalyzer] getJsonPaged: stopped after ${maxPages} pages with more results still available — ${url}`);
    }
    return { items: all, truncated };
  }

  // Convenience wrapper for the many call sites that don't need to react to
  // truncation themselves (small/bounded collections like a file's version
  // history) — see getJsonPagedMeta for callers that must propagate it
  // (e.g. as sizeApproximate) instead of silently under-counting.
  public async getJsonPaged(
    url: string,
    signal?: AbortSignal,
    maxPages = 400,
    skipBatch = false,
    onPage?: (fetchedSoFar: number) => void,
    group: BatchGroup = 'default',
  ): Promise<any[]> {
    return (await this.getJsonPagedMeta(url, signal, maxPages, skipBatch, onPage, group)).items;
  }

  public async runConcurrent<T>(
    tasks: (() => Promise<T | undefined>)[],
    concurrency = 5,
  ): Promise<(T | undefined)[]> {
    if (tasks.length === 0) return [];
    const results: (T | undefined)[] = new Array(tasks.length);
    let idx = 0;
    const worker = async () => {
      while (idx < tasks.length) {
        const i = idx++;
        try { results[i] = await tasks[i](); }
        catch { results[i] = undefined; }
      }
    };
    await Promise.all(Array.from({ length: Math.min(concurrency, tasks.length) }, worker));
    return results;
  }
}
