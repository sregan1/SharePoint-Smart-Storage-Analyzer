# Changelog

All notable changes to this project are documented here.

---

## [1.2.0] — 2026-07-31

### Added

- **Refresh button in the Explorer**
  Folder and library sizes are cached briefly after loading. Refresh clears
  the cache for the current site and re-measures whatever is currently on
  screen, for when files have just been added, deleted, or moved and the
  cached numbers no longer match.

- **Version Count**
  Alongside Version History Size, the Explorer's List View and the Storage
  Report now also show a **Version Count** column — how many older versions
  of a file are being retained (the current version doesn't count) — in the
  UI and in every export (Explorer and Storage Report, Excel and CSV). The
  Excel exports also gained a human-readable Version History Size column
  next to the existing raw-bytes one, matching the existing Size/Size(bytes)
  pattern. Export column headers were also standardized to "Version History
  Size" everywhere (previously "Version history" in places).

- **Site name in export filenames**
  Explorer and Storage Report exports (Excel and CSV) now prefix the
  filename with the site's name, so files exported from different sites
  don't overwrite each other when saved to the same folder.

- **"Other (N items)" folding, explained**
  The Treemap's "Other" cell (the smallest items folded together once a
  folder or the library-root view has more items than can be usefully drawn
  as separate squares) now has a hover tooltip explaining what it is, plus a
  persistent note above the Treemap with a one-click **Switch to List view**
  button — the List view has no such folding and shows every item
  individually.

- **A distinct "≥" (at least) indicator for partially-measured sizes**
  A folder or library whose fallback measurement stopped at this view's
  request budget, rather than finishing, now shows as "≥ &lt;size&gt;" in
  both the Treemap and List View — a floor, not an estimate — with a tooltip
  explaining that opening the folder directly, or raising **Concurrent API
  requests** in Settings, measures deeper. This is now clearly distinguished
  from **"Unknown"**, which means nothing could be measured at all.

- **Version history size in the Explorer Treemap**
  The existing "Include version history size" toggle (previously List View
  only) now also applies to the Treemap: file squares are sized by file +
  version history combined, so a file with a lot of retained version history
  visibly stands out. Hovering a cell (or its on-screen label, for large
  enough squares) breaks the total back down, e.g. "24.6 MB (18.2 MB file +
  6.4 MB version history)". The toggle itself moved to a shared location
  above both views instead of living only in the List View.

- **Clarified that folder totals never include version history**
  An info tooltip next to the version-history toggle, and a similarly-styled
  one next to Storage Report's "Total size" tile, spell out that both
  figures are exact sums of current file content only — SharePoint has no
  recursive version-history rollup, so a folder or site total never includes
  it regardless of whether the toggle is on. Version history size is always
  a separate, additive number.

- **Processing indicator while the Treemap measures its container**
  A brief spinner now covers the Treemap during the one-frame window before
  its container has been measured, instead of that window rendering as a
  blank box that could look identical to "this folder is empty."

- **GitHub Actions release workflow**
  Pushing a `vX.Y.Z` tag now builds the `.sppkg` and publishes it to the
  repo's Releases page automatically (`.github/workflows/release.yml`) —
  tagging a version is now sufficient to produce a release artifact, with no
  separate manual build/upload step.

- **The Explorer now opens on a site-wide library treemap**
  Previously it dropped straight into the default document library, and
  switching libraries meant using the button row. Libraries are now treemap
  squares in their own right, sized by their storage rollup, so the opening
  screen answers "which library is the storage in?" before any drilling. The
  breadcrumb gained an **All libraries** root, and the button row is still
  there for direct jumps.

  Library sizes come from SharePoint's rollup only — this deliberately does
  *not* fall back to a live recursive walk the way folder-level sizing does.
  At folder level that fallback is bounded by one subtree; at library level it
  would mean walking an entire library, so one library with a lagging rollup
  would turn "draw the opening screen" into "scan the whole site". A library
  SharePoint hasn't reported a size for shows as "Unknown" (striped) and is
  measured exactly when opened, capping this view at one probe per library.
  A zero rollup is disambiguated from a genuinely empty library using the
  list's own `ItemCount`.

### Changed

- **Loading states in the Explorer now always show motion**
  A large library could show a featureless gray box for a long stretch, or a
  progress bar that appeared frozen. Three causes: the loading gate tested
  only "is a fetch in flight", which is false in the render between selecting
  a folder and the effect that starts loading it — so an *empty treemap* was
  drawn, complete with "This folder has no subfolders or files"; the progress
  bar rendered at `value={0}`, which draws an empty track indistinguishable
  from a stalled one; and each unit of progress requires one child's entire
  recursive subtree, so the count legitimately sits still for minutes. Now the
  gate tests whether the data on screen actually belongs to the current
  folder, the bar is indeterminate/animated until progress genuinely advances,
  and an elapsed-seconds counter guarantees visible movement regardless.

- **Long stalls now say when they're waiting on SharePoint**
  If the request governor is waiting out throttling, the Explorer says so
  explicitly rather than looking hung — a 60s+ throttle wait was previously
  indistinguishable from a crash.

- **Throttling resilience on large sites — substantially reworked**
  Large tenants could throttle the web part badly enough that it stopped
  working rather than merely slowing down. Four separate causes, all in the
  request layer:

  - *No global request ceiling.* Every concurrency pool enforced only its own
    limit, and several were alive simultaneously — the Explorer ran a
    StorageMetrics probe pool (deliberately doubled to `2x`) with a
    `scanConcurrency`-sized fallback-walk pool nested inside it, while a
    Storage Report scan added its folder queue plus a version-history queue.
    A setting of "6" could mean 16+ requests in flight. There is now one
    ceiling enforced across every pool, in `SpApiClient` itself.
  - *The retry budget was far too small.* 3 attempts at 2/5/10s (~17s total)
    against a tenant whose own `Retry-After` is routinely 30–120s meant
    requests exhausted their retries and **gave up** — marking folders
    unreadable while the server was only asking us to wait. Now 8 attempts
    (2/5/10/20/30/45/60/60s), honoring `Retry-After` up to 120s per attempt.
  - *Throttling was self-sustaining.* A 429 backed off only the request that
    received it while every other worker kept hammering at full concurrency.
    A 429/503 now trips a shared gate that pauses **all** requests for the
    `Retry-After` window, and halves the effective ceiling per consecutive
    throttle (floor 1), recovering a step at a time on clean responses.
  - *No request batching.* Recursive folder walks dominate request volume at
    2 requests per folder, and the Explorer's fallback walk recurses an entire
    subtree for each subfolder whose StorageMetrics rollup reads 0 (which it
    commonly does) — opening one folder with 30 such subfolders of ~200
    folders each was ~12,000 requests to paint a single screen. Requests now
    coalesce into SharePoint's `$batch` endpoint, cutting round trips by
    roughly an order of magnitude with no loss of accuracy.

  Batching is transparent to callers (it lives at the request layer, so the
  recursive walks were left untouched rather than restructured), groups only
  requests targeting the same site — `$batch` cannot span site collections —
  and falls back to individual requests if the endpoint is unavailable, so it
  can never be the reason a folder reads as unreadable.

- **HTTP 406 correctly identified as throttling, not a bad item name**
  Further investigation (console/network evidence against a real large
  tenant) showed 406 is SharePoint redirecting an over-limit request to an
  HTML throttle page, which then fails content negotiation — not a rejected
  folder/file name as originally assumed. 406 is now handled identically to
  429/503: absorbed by the shared throttle gate and retried with the same
  backoff schedule, rather than retried once and then surfaced as a per-item
  error.

- **Right-sized the folder-walk fallback request budget for large sites**
  The fixed 400-folder budget bounding a fallback measurement walk (used when
  SharePoint's `StorageMetrics` rollup is stale or unavailable) was never
  really a request-count safety valve — it was a wall-clock limit wearing a
  request-count costume, since the same 400 folders costs seconds on a fast
  tenant but minutes on one that's actively throttling. It's now derived from
  **Concurrent API requests** (150 folders per unit of concurrency, floor 300,
  ceiling 2400 — 900 at the default of 6, vs. a fixed 400 before), and a
  wall-clock deadline is now the primary bound alongside it (45s for an
  interactive folder drill-in, 20s for the library-root rollups, 15s for the
  post-batch retry pass, which previously had no budget or deadline at all).
  Raising Concurrent API requests now measures deeper before falling back to
  an approximate "≥" result, instead of the folder-count budget and the
  concurrency setting silently disagreeing with each other.

- **`getLibrariesWithStats` now delegates to `getLibraryRollups`**
  Used only on sites with no library named "Documents"/"Shared Documents",
  this previously ran its own fallback walk per library with no shared
  budget — on a site with many such libraries, each one got its own fresh
  400-folder budget, reintroducing the exact unbounded-multiplication problem
  the shared budget exists to prevent. It also trusted `StorageMetrics`
  returning `undefined` as the only fallback trigger, missing the common case
  where it returns a stale 0 for a library that actually has content. It now
  delegates entirely to `getLibraryRollups`, which already gets both right.

- **Storage Report View/Export buttons respond immediately**
  Clicking **View** or **Export Excel/CSV** on a large report could sit doing
  nothing for a noticeable moment before anything happened, because the
  heavy work (loading thousands of rows, building a workbook) is synchronous
  and blocked the browser before it could paint a spinner. Both buttons now
  show a spinner on the very next frame, before that work starts.

### Fixed

- **Files appeared seconds before their folders in the Explorer**
  The folder-load completion handler cleared its loading flag *before*
  bumping the cache version the folder list is derived from. React 17 does not
  batch state updates made from a promise callback, so there was a render
  where the loading gate had already opened but the folder data had not yet
  been picked up — showing the (much faster) file results against a stale or
  empty folder list until a later render caught up.

- **Throttled folders displayed a confident "0 bytes, 0 files"**
  When a folder's size lookup exhausted its retries, the failure was absorbed
  and the folder rendered as an ordinary empty folder — indistinguishable from
  one that genuinely had no content. Such folders are now marked as
  unmeasured: striped rather than solid-filled in the Treemap, "Unknown"
  instead of a size in the List View, both explained on hover, with a warning
  banner above the view. They are also no longer cached, so revisiting retries
  instead of showing "unknown" for the cache's full TTL.

- **HTTP 406 responses failed immediately with no retry and no attribution**
  Every failed request's error now includes the URL that produced it —
  previously only the status and response body were reported, which rarely
  identify which item was at fault. (406's retry behavior was further
  corrected afterward — see **Changed**, above — once it was clear a single
  retry wasn't the right fix.)

- **A folder whose files failed to load could misreport as empty**
  The Explorer's per-folder file fetch silently treated any failure
  (throttling, a transient error) as "this folder has no files," which could
  contradict the folder's own size/file-count shown one level up. It now
  surfaces the same kind of error banner the folder-listing fetch already
  used for this exact failure mode.

- **Version History reporting could silently show nothing**
  When per-file version-history lookups failed (most often throttling) during
  a Storage Report scan, the count was tracked internally but never shown —
  the Version History Size tile and column simply read 0/blank for every
  file, with no indication anything had gone wrong. This is now included in
  the same partial-scan warning banner used for skipped folders/subsites.

- **"≥ 0 B" — a folder or library that was never actually measured could look
  like a confirmed near-empty one**
  When several folders (in `getFolderChildren`) or libraries (in
  `getLibraryRollups`) share one measurement budget, a later item's own walk
  could hit that budget already exhausted by earlier siblings, before it had
  listed anything at all. The result was reported as a genuine — if
  approximate — `0 bytes` floor: rendered as "≥ 0 B", sorted to the bottom,
  and drawn as a near-invisible treemap square, indistinguishable from an
  item that really was confirmed almost-empty. This case is now detected and
  reported as **"Unknown"** instead, consistent with every other case where
  nothing could be measured.

- **Exporting from the Explorer's root library view used stale folder data**
  The List View's Export Excel/CSV buttons always read from the currently
  drilled-in folder's data, even when the view on screen was actually the
  site-root library treemap — exporting from the root silently produced a
  stale or unrelated folder's listing instead of the libraries shown.

- **Exporting a cross-site viewed Storage Report used the wrong site**
  Reopening a saved scan from another site (via **Show reports from other
  sites**) and then exporting used the current page's site rather than the
  viewed report's own site, which could put the wrong site's name on the
  export or fail outright depending on permissions.

### Removed

- **"Estimate" size badge references in the docs**
  README and USER-GUIDE described a visible "Estimate" badge for when
  SharePoint's Storage Metrics rollup is unavailable; the underlying
  fallback estimate still exists and is used automatically, but no such
  badge is rendered anywhere in the UI. The docs no longer claim otherwise.

---

## [1.1.0] — 2026-07-11

### Added

- **Cancel a running Storage Report scan**
  A Cancel button appears while a scan runs. Cancellation is cooperative
  (checked between queued folders/libraries/sites), so it takes effect within
  moments rather than instantly interrupting in-flight requests. Partial
  results collected before cancellation are shown but not saved to history.

- **Live file-count progress during a scan**
  The "files scanned" counter now updates continuously while a single library
  is being walked, instead of sitting at 0 until the whole library finishes.

- **Partial-scan warnings**
  A scan that could not read some folders (permission errors, throttling
  exhaustion, etc.) or subsites now completes with a visible warning stating
  how many were skipped, instead of silently under-reporting totals.

- **Scan history: site awareness**
  Each saved scan now shows which site it was run against. The history list
  and the "compare two scans" feature default to the current site, with an
  explicit toggle to show reports from other (sub)sites; comparing two
  reports from different sites now shows a caution instead of a silent,
  potentially meaningless diff.

- **Scan history: bounded storage**
  A scan whose file count would strain `IndexedDB` quota now saves only its
  stale/very-stale rows (flagged "Partial" in the history list) — summaries
  and scan comparisons are unaffected, since both only ever need stale-tier
  data.

- **Scan history: viewing context**
  Reopening a saved scan now shows its timestamp, site, and displays the
  archival-status legend using the thresholds *that scan* was classified
  with, rather than whatever the current Settings happen to be.

- **Keyboard accessibility**
  Treemap folder cells and sortable table headers are now focusable and
  operable from the keyboard (Enter/Space), with `aria-sort` reflecting the
  active sort column.

### Fixed

- **Invalid concurrency/threshold values could hang the whole web part**
  A corrupted `localStorage` value, or typing an out-of-range number directly
  into a Settings spin button, could set scan concurrency to `NaN`/`0` —
  which the scan queue treats as "never run anything," hanging every scan and
  folder load indefinitely, permanently (the bad value was persisted). All
  three settings (concurrency, stale/very-stale day thresholds) are now
  clamped at every layer they pass through: on read, in the Settings UI, and
  in the scan engine itself.

- **A failed scan of the site itself could report as an empty success**
  If the requested site's own document library listing failed (a transient
  error, throttling), the scan previously completed with a "successful" empty
  report instead of an error. Subsites still fail silently and are now
  counted instead (see "Added" above).

- **A scan whose results couldn't be saved to history reported as "Scan
  failed"** even though the scan itself succeeded and its results were on
  screen (e.g. an oversized report hitting `IndexedDB` quota). These are now
  reported as separate, non-fatal warnings.

- **A folder that failed to load could be cached as permanently empty**
  A transient error while resolving a folder's contents (throttling, a
  permission error) was cached as "this folder is empty" for the cache's
  10-minute lifetime, and the Explorer showed no indication anything had
  gone wrong. Failed folders are no longer cached, and the Explorer now
  shows the underlying error instead of a false empty state.

- **Storage Report scans could silently under-report totals**
  Folders that failed mid-walk (permission errors, throttling, the list view
  threshold) were dropped from the report with no indication — see the new
  partial-scan warning above.

- **Folder/file listings were never paginated**
  Every `Files`/`Folders` API call fetched a single page; a folder with an
  extremely large number of items could silently lose entries beyond that
  page. These calls now follow SharePoint's paging links.

- **Picture Libraries and Site Pages were never scanned**
  The library query filtered to Document Libraries only, while the rest of
  the app already treated Picture Libraries and Site Pages as scannable —
  both can hold real storage and are now included everywhere.

- **Changing archival thresholds didn't update already-loaded folders**
  The Explorer cached a folder's file list with its Active/Stale/Very-stale
  classification already computed; changing the thresholds in Settings and
  revisiting a previously-viewed folder kept showing the old classification.
  Classification is now computed at display time from the current
  thresholds.

- **CSV export was vulnerable to formula injection**
  A file or author name starting with `=`, `+`, `-`, or `@` would be
  interpreted as a formula by Excel/Sheets when the exported CSV was opened.
  Such values are now neutralized with the standard leading-quote mitigation.

- **Theme updates could apply a stale color**
  The in-app theme was only recomputed when the site's primary brand color
  changed, missing updates to the other palette shades.

- **A very large folder listing could crash the Explorer**
  Computing the largest size in a folder's file list used
  `Math.max(1, ...values)`, which can throw once the argument count gets
  large enough; replaced with a plain reduction.

- **Scan history trimming could get stuck over its own limit**
  Adding a new scan when history was already over its 10-report cap deleted
  only one old report regardless of how far over the cap it was, so it could
  never shrink back down. It now deletes as many as needed to settle at the
  cap, and a rejected save no longer leaves the database connection open.

### Changed

- **Export dates are now real dates, not locale strings**
  Excel exports write typed date cells (sortable/filterable), and CSV
  exports use ISO `yyyy-mm-dd` instead of locale-formatted strings.

---

## [1.0.0] — 2026-07-03

### Added

- **Explorer — Treemap and List views**
  A WizTree-style treemap of the current library, with click-to-zoom drill-down
  into any folder, plus a sortable List view of the same folder's contents
  (folders and files together, largest first by default). Both views share the
  same drill-down state — switching between them keeps you in the same folder.

- **Library switcher and breadcrumb navigation**
  A button row switches between every document library on the site without
  leaving the Explorer. Breadcrumbs let you jump back to any ancestor folder
  in one click.

- **Folder size and file count shown on treemap cells**
  Folder cells display their total rolled-up size and file count (recursive,
  not just immediate children) directly on the cell — and in the tooltip —
  instead of size alone.

- **Archival status legend with real thresholds**
  A shared status legend shown on the Treemap, List, and Storage Report spells
  out the actual configured day thresholds — e.g. "Stale (180–364d)" — instead
  of just naming the Active/Stale/Very-stale tiers.

- **Excel and CSV export**
  The Explorer's List view and the Storage Report's results can both be
  exported to a styled `.xlsx` workbook or a plain `.csv` file.

- **Storage Report — full site scan**
  Scans a site (and optionally its subsites) for every file, classifying each
  as Active, Stale, or Very stale. Includes configurable scope (subsites,
  hidden/system libraries) and scan concurrency, a live progress bar with
  file count and elapsed timer, and a sortable results table with a toggle to
  show only archival candidates.

- **Scan history and report compare**
  Every completed scan is saved automatically (IndexedDB, capped at the 10
  most recent). Past scans can be reopened without re-scanning, and any two
  saved scans can be compared to see size change, new archival candidates,
  and resolved items over time.

- **Settings screen**
  A dedicated in-app Settings page for archival day thresholds, scan scope,
  and scan concurrency, plus instructions for setting the web part's default
  view via the SharePoint property pane.

- **Progress bar with live folder count**
  Loading a folder's sizes shows a progress bar and a "N of M folders"
  counter once the folder count is known, instead of an unlabeled spinner.

- **Session cache for folder sizes**
  Resolved folder sizes are cached for 10 minutes (`sessionStorage`), so
  revisiting a folder — including after a page refresh — doesn't repeat the
  same expensive size computation.

- **Minimum clickable cell size in the treemap**
  The smallest items in the treemap are floored to a minimum on-screen size
  so they stay clickable instead of shrinking to unclickable slivers.

### Fixed

- **Folders not appearing in the treemap**
  Any item — folder or file — whose resolved size was zero was filtered out
  of the treemap layout entirely, rather than rendering as a small but
  visible, clickable cell. A folder with a rollup of 0 bytes (empty, or a
  failed size lookup) would silently vanish.

- **Folders showing 0 files and 0 bytes**
  The fallback size estimate used when SharePoint's `StorageMetrics` rollup
  isn't available only summed files directly inside a folder, never its
  subfolders — so a folder whose content lived entirely in nested subfolders
  (the common case) reported nothing. The fallback now walks the whole
  subtree recursively.

- **Folders still showing 0/0 after the above fix**
  `StorageMetrics` is computed by a periodic SharePoint background job that
  can report a stale "0 bytes" for a folder that demonstrably has content.
  The rollup is now trusted only when it reports actual content (bytes > 0);
  otherwise a live recursive walk is used instead of accepting the stale
  zero.

- **A folder taking up to a minute to load**
  The live-walk fallback was opening its own bounded worker pool nested
  inside an already-bounded outer loop, multiplying concurrency instead of
  capping it and tripping SharePoint's request throttling. Every fallback
  walk triggered by one folder load now shares a single queue, so total
  concurrency stays capped regardless of how many folders need the fallback.

- **Treemap rendering long thin rectangles instead of near-square cells**
  The squarified-treemap layout algorithm had the "which side is fixed"
  logic backwards — it stretched each row across the box's longer dimension
  instead of its shorter one, the part of the algorithm that actually keeps
  cells square. Swapped to match the correct algorithm.

- **Web part icon not appearing in the "Add a web part" picker**
  The configured `officeFabricIconFontName` was not a valid Fluent/Fabric
  icon name, so the picker rendered no icon at all.

- **Build/serve failing with `ENOENT` after the project rename**
  `config/config.json` still referenced the pre-rename source folder and
  class names, so a fresh `gulp bundle`/`gulp serve` failed to find them
  even though the running dev server (started before the rename) hadn't
  noticed.

- **Excel export silently doing nothing on failure**
  A failed one-time load of the Excel-generation library was cached
  permanently with no retry — every future export attempt (from either
  export button) would immediately re-fail. Neither export button caught or
  surfaced errors either, so a failure produced no visible feedback at all.

- **Every page load computing sizes for every library on the site**
  The library switcher only ever displays library names, never sizes, but
  the app was computing a full recursive size rollup for every library up
  front regardless — the single biggest contributor to the initial loading
  delay. Size computation is now skipped whenever a named default library
  ("Documents" / "Shared Documents") is found.

### Changed

- **Project and web part renamed** from "Smart Data Analyzer" to "Smart
  Storage Analyzer" throughout — web part title and icon, class names,
  bundle output name, `localStorage` key prefixes, and the scan-history
  database name.

- **Explorer defaults to the Treemap view**, with an explicit progress
  indicator while folder sizes resolve instead of a blank area.

### Removed

- **Editable site URL in the header.** The web part now always analyzes the
  site it's placed on; the "Change URL" control and its associated state
  were removed to keep the header simple and match standard SPFx usage.

---

## [0.0.1] — Initial internal build

First internal build. Core features: Explorer (treemap + list drill-down)
and Storage Report (site scan with archival tiering).
