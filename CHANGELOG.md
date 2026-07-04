# Changelog

All notable changes to this project are documented here.

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
