import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';
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

  // 406 from SharePoint REST is usually a malformed/rejected request (a
  // folder or file name with a character the OData parser or IIS request
  // filtering chokes on, or an over-long path) — not transient in the way
  // 429/503 are. But it has also been observed to clear on its own on SPO's
  // front end under load, so one short retry is worth it before giving up;
  // unlike 429/503 there's no Retry-After to honor, so a single fixed pause.
  private static readonly RETRY_406_DELAY_SECONDS = 3;

  public async getJson(url: string, attempt = 0): Promise<any> {
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

    if (resp.status === 429 || resp.status === 503) {
      const fallback = SpApiClient.RETRY_BACKOFF_SECONDS[
        Math.min(attempt, SpApiClient.RETRY_BACKOFF_SECONDS.length - 1)
      ];
      const header = parseInt(resp.headers.get('Retry-After') ?? '', 10);
      const waitSeconds = Math.min(
        isNaN(header) ? fallback : Math.max(header, 1),
        SpApiClient.MAX_RETRY_AFTER_SECONDS,
      );
      // Trip the shared gate even on the final attempt: other in-flight work
      // still benefits from backing off, whether or not THIS request retries.
      this.noteThrottled(waitSeconds);
      if (attempt < SpApiClient.RETRY_BACKOFF_SECONDS.length) {
        return this.getJson(url, attempt + 1);
      }
      throw new Error(
        `HTTP ${resp.status} on ${url} — still throttled after ` +
        `${SpApiClient.RETRY_BACKOFF_SECONDS.length} attempts. Lower Concurrent API requests in Settings.`,
      );
    }

    if (resp.status === 406 && attempt < 1) {
      await new Promise((r) => setTimeout(r, SpApiClient.RETRY_406_DELAY_SECONDS * 1000));
      return this.getJson(url, attempt + 1);
    }
    if (!resp.ok) {
      const txt = await resp.text();
      // Include the URL — 406/400-class failures are almost always specific
      // to one folder/file (a name with a problematic character, or an
      // over-long path), and the body SharePoint returns rarely says which;
      // the URL itself is usually the only clue pointing at the culprit.
      throw new Error(`HTTP ${resp.status} on ${url} — ${txt.substring(0, 300)}`);
    }
    this.noteSuccess();
    return resp.json();
  }

  // Fetches a collection endpoint and follows server-side paging links so
  // results beyond the $top page size are not silently dropped. maxPages is a
  // safety valve against runaway loops on enormous collections.
  public async getJsonPaged(url: string, signal?: AbortSignal, maxPages = 50): Promise<any[]> {
    const all: any[] = [];
    let next: string | undefined = url;
    let page = 0;
    for (; next && page < maxPages && !signal?.aborted; page++) {
      const data = await this.getJson(next);
      all.push(...valueArray(data));
      next = data?.['odata.nextLink'] ?? data?.['@odata.nextLink'] ?? data?.d?.__next;
    }
    if (next && page >= maxPages) {
      // eslint-disable-next-line no-console
      console.warn(`[SmartStorageAnalyzer] getJsonPaged: stopped after ${maxPages} pages with more results still available — ${url}`);
    }
    return all;
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
