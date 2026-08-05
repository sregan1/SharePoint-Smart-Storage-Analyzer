# SharePoint Smart Storage Analyzer — Storage Usage & Archival Report Web Part

[![Website](https://img.shields.io/badge/Website-sharepointsmartsolutions.com-blue)](https://sharepointsmartsolutions.com/smart-data-analyzer) [![User Guide](https://img.shields.io/badge/User%20Guide-USER--GUIDE.md-informational)](USER-GUIDE.md) [![Download](https://img.shields.io/badge/Download-Latest%20Release-CA5010?logo=github&logoColor=white)](../../releases/latest) [![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

A free, open-source SPFx web part that helps SharePoint site owners find storage that's safe to archive — browsing libraries and folders, and scanning a site for stale files — entirely from the browser, with no PowerShell or third-party tools required.

![SPFx](https://img.shields.io/badge/SPFx-1.21.1-0078D4?logo=microsoft&logoColor=white) ![React](https://img.shields.io/badge/React-17.0.1-61DAFB?logo=react&logoColor=black) ![TypeScript](https://img.shields.io/badge/TypeScript-5.3-3178C6?logo=typescript&logoColor=white) ![Fluent UI](https://img.shields.io/badge/Fluent%20UI-v9-0078D4?logo=microsoft&logoColor=white) ![ExcelJS](https://img.shields.io/badge/ExcelJS-4.4-107C10)

---

## Features

> **Requires Site Owner access.** Storage totals rely on the same right as classic Site Settings → Storage Metrics. Anyone without Site Owner access will see the web part with a clear warning banner instead of data.

### Home

The web part opens on a landing screen with three cards — **Tree View**, **List View**, and **Storage Report** — each with a short description of what it does. A site administrator can instead configure the web part to open directly on any one of the three; see [Configuration](#configuration).

![Home screen with Tree View, List View, and Storage Report cards](docs/screenshots/00_home.png)

### Tree View & List View

Two entry points into the same screen — a WizTree-style treemap of every document library on the site, sized by storage, or the same libraries and folders as a sortable table — then drills down through folders to individual files. A tab switches between the two without losing your place.

![Explorer treemap view showing folders and files sized by storage weight, color-coded by archival status](docs/screenshots/01_explorer_treemap.png)

| Feature | Description |
|---|---|
| **Site-wide library treemap** | Opening view sizes every library on the site by its storage rollup, answering "which library is the storage in?" before any drill-down. Libraries SharePoint hasn't yet reported a size for render as "Unknown" and are measured exactly on open — deliberately, so the root view costs one probe per library rather than a full site walk |
| **Treemap drill-down** | Click any library or folder square to zoom into it; square size reflects storage weight at a glance. A folder measured only partially (its walk hit this view's request budget) shows as "≥ &lt;size&gt;" — a floor, not an estimate — distinct from "Unknown" (nothing could be measured) |
| **"Other" folding** | Folders/libraries beyond the largest ~40 fold into a single "Other (N items)" cell instead of drawing slivers too small to see or click; a note above the treemap offers one click through to the List view, which shows every item individually |
| **Library switcher** | A button row switches between every document library on the site without leaving the view |
| **Refresh** | Clears cached folder/library sizes for the current site and re-measures what's on screen, for when content has changed since the last load |
| **Breadcrumb navigation** | Jump back to any ancestor folder in one click |
| **List View** | Toggle to a sortable table of the same folder's contents — folders and files together, largest first by default |
| **Version history size & count** | Optional per-folder toggle that adds each file's retained-version storage (and how many old versions are retained) on top of its current size — sized into the Treemap's file squares and shown as its own columns in the List View. Folder totals never include it (no recursive rollup exists for it) |
| **Excel / CSV export** | Export the current List View (name, size, item count, modified date, archival status, version history size and count if enabled) to `.xlsx` or `.csv`; filenames are prefixed with the site name |
| **Archival status** | Files are tagged Active / Stale / Very stale based on configurable last-modified thresholds, shown in both the treemap and the list |

Both view modes share the same drill-down state — switching from Treemap to List (or vice versa) keeps you in the same folder.

![Explorer list view showing a sortable table of folders and files with size, item count, modified date, and archival status](docs/screenshots/02_explorer_list.png)

### Storage Report

Scan a site — and optionally its subsites — and export a report of archival candidates.

![Storage Report results showing summary tiles for total size, files scanned, and stale/very-stale counts, plus a sortable file-level results table](docs/screenshots/05_report_results.png)

| Feature | Description |
|---|---|
| **Configurable scope** | Include subsites and hidden/system libraries in the scan |
| **Concurrent, throttling-aware scan** | Adjustable request concurrency with a stage-by-stage progress display (what's being read, a per-stage count, an ETA once there's enough data to estimate from) and an explicit "paused, waiting on SharePoint" state instead of an apparent freeze |
| **Cancelable scans** | Stop a running scan and still see the partial results collected so far (not saved to history) |
| **Version history size & count** | Optional toggle that adds per-file version-history size and count columns, plus summary tiles for the total size and total retained-version count across the scan — additive to Total size, not included in it. Free on most libraries (it rides the same bulk read); on a list that won't report it in bulk, only files that can actually have retained versions are measured individually |
| **Partial-scan reporting** | Folders/subsites that fail to read (permissions, throttling), and files whose version history specifically couldn't be read, are called out with a warning, expandable per-item error details, and a copy-to-clipboard action, instead of silently under-reporting |
| **Archival tiering** | Every file is classified Active, Stale, or Very stale based on configurable last-modified thresholds |
| **In-browser results table** | Sortable results with a toggle to show only archival candidates |
| **Excel export** | Color-coded `.xlsx` workbook with a Summary sheet and a full file-level Details sheet; filename prefixed with the site name |
| **CSV export** | Plain-text alternative for scripted processing; filename prefixed with the site name |
| **Scan history** | Past scans persist in IndexedDB (10 most recent), with a cross-site visibility toggle. Every saved report keeps its complete file listing — there's no row cap; if browser storage genuinely runs out of room, the *oldest* report's listing (never its summary) is evicted to make room, flagged with a "No file list" badge |
| **Report compare** | Diff two saved scans to see size change, new archival candidates, and resolved items over time — with a warning if the two scans are from different sites |

---

## Prerequisites (for Development Only)

| Requirement | Detail |
|---|---|
| **Node.js** | 18.x (`>=18.17.1 <19.0.0`) |
| **gulp-cli** | Install globally: `npm install -g gulp-cli` |
| **SharePoint** | Online (Microsoft 365) |
| **SPFx** | 1.21.1 |
| **Permissions to deploy** | Site Owner or above |

---

## Development Setup

```bash
# Install dependencies
npm install

# Edit config/serve.json and set initialPage to your hosted workbench URL:
# "initialPage": "https://<your-tenant>.sharepoint.com/sites/<your-site>/_layouts/workbench.aspx"

# Start the local dev server (opens hosted workbench)
gulp serve
```

The local workbench at `https://localhost:4321/temp/workbench.html` does not have SharePoint REST API access. Use the **hosted workbench** URL above for full functionality.

---

## Build & Deploy

```bash
# Production bundle (minified, ship mode)
gulp bundle --ship

# Create the .sppkg deployment package
gulp package-solution --ship

# Or run both in one step:
npm run ship
```

The package is written to `sharepoint/solution/smart-storage-analyzer.sppkg`.

Pushing a `vX.Y.Z` tag also triggers a GitHub Actions workflow (`.github/workflows/release.yml`) that builds this same package and publishes it to the repo's [Releases](../../releases) page automatically.

**Deploy to SharePoint:**

1. Upload `smart-storage-analyzer.sppkg` to the tenant or site App Catalog.
2. Click **Deploy** when prompted.
3. Navigate to the SharePoint page where you want to add the web part, click **Edit**, and add **Smart Storage Analyzer** from the web part picker.

---

## Configuration

To change web part settings, put the page in **Edit** mode, click the web part pencil icon, and use the property pane. Everything else — scope, archival thresholds, and scan concurrency — is configured from the in-app **Settings** screen.

**General (property pane)**

| Setting | Default | Description |
|---|---|---|
| **Default view on open** | Home | The screen shown when the web part first loads. Options: Home, Tree View, List View, Storage Report |

**Scope (in-app Settings)**

| Setting | Default | Description |
|---|---|---|
| **Include subsites in Storage Report scans** | Off | Also walks every subsite beneath the current site. Does not affect Tree View / List View |
| **Include system and hidden libraries** | Off | Includes Style Library, Form Templates, and other libraries normally hidden from default views. Applies to all tools |

**Archival thresholds (in-app Settings)**

| Setting | Default | Description |
|---|---|---|
| **Stale after (days)** | 180 | Files not modified in this many days are flagged "Stale" — a candidate for review |
| **Very stale after (days)** | 365 | Files not modified in this many days are flagged "Very stale" — a strong archival candidate |

**Performance (in-app Settings)**

| Setting | Default | Description |
|---|---|---|
| **Concurrent API requests** | 6 | How many SharePoint API requests run in parallel during scans and folder loads (1–15). SharePoint's throttling limit is dynamic, not fixed — the app retries automatically on throttling (HTTP 429/503/406), but very high values can still net out slower. Also sizes how deep Tree View / List View's fallback folder/library measurement goes before falling back to a "≥" (at least) result |

---

## Project Structure

```
src/webparts/smartStorageAnalyzer/
├── components/
│   ├── App.tsx                        # Root component — routing, header nav, permission check, theme wiring
│   ├── HomeView.tsx                    # Landing screen — Tree View / List View / Storage Report cards
│   ├── ExplorerView.tsx               # Tree View & List View — library switcher, treemap + list drill-down, export
│   ├── StorageReportView.tsx          # Scan configuration, results, history, compare
│   ├── SettingsView.tsx               # Full-page settings screen
│   └── shared/                        # Shared UI: StorageTable, SizeBar, Treemap, tier badges
├── services/
│   ├── StorageAnalyzerService.ts      # Facade over the sp/ modules (stable public API)
│   ├── sp/                            # API client, site discovery, storage metrics, library stats, scan
│   ├── ExcelExportService.ts          # ExcelJS workbook generation (lazy-loaded chunk)
│   └── ReportHistoryService.ts        # IndexedDB-backed scan history
├── utils/
│   ├── archivalClassification.ts      # Age calculation and Active/Stale/VeryStale tiering
│   ├── reportDiff.ts                  # Pure diff between two stored reports
│   └── treemapLayout.ts               # Squarified treemap layout algorithm
├── models/
│   └── models.ts                      # Shared TypeScript interfaces
└── SmartStorageAnalyzerWebPart.ts       # SPFx entry point, property pane, theme wiring

config/
├── package-solution.json              # Solution ID, version
└── serve.json                         # Local dev server config — set initialPage here
```

---

## Key Dependencies

| Package | Purpose |
|---|---|
| `@microsoft/sp-webpart-base` | SPFx web part base class and framework integration |
| `@fluentui/react-components` | Fluent UI v9 — components and theme tokens |
| `react` / `react-dom` | UI rendering (v17) |
| `exceljs` | In-browser Excel workbook generation for report exports (lazy-loaded) |

---

## Troubleshooting

**"I see a yellow warning banner and no data"** — The signed-in account does not have Site Owner access. Storage totals rely on the same right as classic Site Settings → Storage Metrics. Ask a site owner to grant Owner access or run the tool on your behalf.

**"gulp serve opens but the web part shows no data"** — The local workbench (`localhost:4321`) cannot authenticate to SharePoint REST. Switch to the hosted workbench: edit `config/serve.json` and set `initialPage` to `https://<tenant>.sharepoint.com/_layouts/15/workbench.aspx`.

**"npm install fails" or build errors about Node version** — This project requires Node 18.x exactly (`>=18.17.1 <19.0.0`). Run `node --version` to confirm. Use `nvm` or `nvm-windows` to switch versions.

**"The Storage Report scan takes a very long time"** — Scan time scales with file count and, if enabled, the number of subsites. Version history is normally free (it rides along in the same bulk read), but a list that won't report it in bulk has to be measured one file at a time; the progress display names the stage that's running and gives a per-stage estimate. Narrow the scope in Settings (disable subsites, hidden libraries, or version history) or lower scan concurrency if you're seeing throttling (HTTP 429) errors.

**"An error mentions HTTP 406"** — this is SharePoint throttling, not a bad request: a 406 occurs when SharePoint redirects an over-limit request to an HTML throttle page instead of the JSON response that was asked for. `spCore.ts` treats 406 identically to 429/503 — it's absorbed by the shared throttle gate and retried with backoff, not surfaced as a per-item error. If it persists, lower **Concurrent API requests** in Settings.

---

## Limitations

- All tools require **Site Owner** access. Members, Visitors, Limited Access users, and guests cannot use any feature.
- Archival tiering is based on **last modified date only** — SharePoint does not expose a reliable last-accessed signal at scale.
- This tool reports and browses storage; it does not move, archive, or delete anything.
- Runs entirely as the signed-in user — results reflect that user's view and access.
- Scan history persists in **IndexedDB** in the browser, capped at the 10 most recent scans. Clearing browser data removes all saved scan results.

---

## License

[MIT](LICENSE) © 2026 Sean Regan
