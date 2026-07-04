# SharePoint Smart Storage Analyzer — User Guide

**Version 1.0.0**
**Applies to:** SharePoint Online

---

## Table of Contents

1. [Overview](#overview)
2. [Who Is This For?](#who-is-this-for)
3. [Getting Started](#getting-started)
4. [Explorer](#explorer)
5. [Storage Report](#storage-report)
6. [Understanding Archival Status](#understanding-archival-status)
7. [Settings](#settings)
8. [Web Part Configuration](#web-part-configuration)
9. [Security & Privacy](#security--privacy)
10. [Frequently Asked Questions](#frequently-asked-questions)
11. [Troubleshooting](#troubleshooting)

---

## Overview

**SharePoint Smart Storage Analyzer** is a browser-based storage-auditing tool built directly into SharePoint Online as a web part. It helps site owners find storage that's safe to archive or clean up — browsing libraries and folders visually, and scanning a whole site for stale files — entirely from the browser, with no PowerShell, no third-party tools, and no IT assistance required.

SharePoint's default interface makes it hard to see *where* storage is actually going. Libraries show a single total, folders don't show their own size, and there's no built-in way to ask "what's old and safe to remove?" Smart Storage Analyzer answers both questions from a single interface: an interactive treemap for visual exploration, and a full-site scan for a complete, exportable report.

### What You Can Do

| Tool | Purpose |
|------|---------|
| **Explorer** | Visually browse a document library as a size-weighted treemap (or a sortable list), drilling into any folder to see what's using space |
| **Storage Report** | Scan an entire site (optionally including subsites) and produce a full report of every file, tagged by how stale it is |

---

## Who Is This For?

- **Site Owners** who need to understand where a site's storage is going and what's safe to archive
- **IT Administrators** doing periodic storage cleanup or preparing for a storage-quota conversation
- **Anyone asked "why is this site so big?"** who needs a fast, visual answer without exporting anything first

> **Note:** The web part runs as the currently signed-in user and only shows content, and total sizes for, sites that user can access. See [Getting Started](#getting-started) for the permission level required.

---

## Getting Started

### Prerequisites

You must be a **Site Owner** on the site (or otherwise hold the Manage Web / Manage Permissions right). Reading a folder or library's rolled-up storage size relies on the same underlying right as classic **Site Settings → Storage Metrics**. Anyone without it — Members, Visitors, and guests — will see the web part, but a warning banner instead of data.

### Accessing the Web Part

The web part is added to a SharePoint page by a site owner or administrator. Once added, navigate to the page where it has been placed — the web part automatically connects to that site; no setup or configuration is required to get started.

By default, it opens on the **Explorer** view, showing a treemap of the site's default document library. (A site administrator can change this — see [Web Part Configuration](#web-part-configuration).)

![The Explorer opening directly into a treemap of the default document library](docs/screenshots/01_explorer_treemap.png)

### If You Don't Have Access

If your account lacks the required permission, a banner appears below the header:

> **Site Owner access required —** Storage totals rely on the same right as classic Site Settings → Storage Metrics. Contact a site owner if you have questions.

Contact a site owner to request access, or have them run the tool on your behalf.

---

## Explorer

### What It Does

The Explorer opens directly into a **treemap** of the site's default document library — no picker screen first. Each rectangle is a folder or file, sized proportionally to how much storage it uses, so the biggest space users are visually obvious at a glance. Click any folder to zoom into it and see what's inside.

### The Treemap View

- **Square size = storage weight.** Bigger rectangles use more space. Folders are shown in blue; files are colored by their [archival status](#understanding-archival-status).
- **Click a folder** to drill into it — the treemap updates to show that folder's own contents.
- **Folder cells show their size and file count** (e.g. "24.6 MB · 138 files") when the cell is large enough to fit the text; smaller cells still show at least the name, and hovering any cell shows full detail in a tooltip.
- The smallest items are floored to a minimum on-screen size so they stay easy to click even on a site with many small files — treemap proportions favor clickability slightly over strict mathematical precision for the tiniest items.
- While a folder's sizes are loading, a progress bar with a live **"N of M folders"** counter is shown instead of the treemap; the treemap appears once every size has resolved.

### The List View

Click the **List** tab to switch to a sortable table of the same folder's contents — folders and files together, largest first by default. Click any column header to sort by it; click again to reverse. This view is often more convenient than the treemap when you want to scan names and exact sizes rather than compare visually.

![The List view showing the same folder as a sortable table, with Export Excel and Export CSV buttons above it](docs/screenshots/02_explorer_list.png)

Switching between Treemap and List keeps you in the same folder — you don't lose your place.

### Library Switcher and Breadcrumbs

- The button row at the top switches between every document library on the current site without leaving the Explorer.
- The breadcrumb trail below it shows your current path from the library root; click any earlier segment to jump back to it.

### Exporting

From the **List** view, use **Export Excel** or **Export CSV** to download the current folder's listing — name, size, item count, modified date, and archival status — for the folder you're currently viewing (not the whole site; use [Storage Report](#storage-report) for that).

---

## Storage Report

### What It Does

The Storage Report scans an entire site — and optionally every subsite beneath it — and classifies every file it finds as Active, Stale, or Very stale based on how long it's been since it was last modified. Where the Explorer is for visual browsing, the Storage Report is for producing a complete, exportable, and comparable record.

### Running a Scan

1. Switch to the **Storage Report** tab.
2. Optionally check **Include subsites** to also walk every subsite beneath the current site.
3. Optionally check **Include hidden/system libraries** to also scan Style Library, Form Templates, and other libraries normally hidden from default views.
4. Click **Run scan**.

![Storage Report before a scan has been run, showing the scope checkboxes and Run scan button](docs/screenshots/03_report_config.png)

While the scan runs, a progress bar tracks how many libraries have been processed, alongside a live file count and an elapsed timer.

![A scan in progress, showing the progress bar, current library, file count, and elapsed time](docs/screenshots/04_report_running.png)

### Browsing the Results

Once the scan completes, four summary tiles show the total size scanned, total files, and the count and size of Stale and Very-stale files. Below that, the full results table lists every file — library, name, size, modified date, author, and archival status — sortable by any column.

![Completed scan results, showing the summary tiles and the sortable file-level results table](docs/screenshots/05_report_results.png)

Check **Show archival candidates only** to filter the table (and the exports) down to Stale and Very-stale files, hiding Active ones.

### Exporting

- **Export Excel** — a color-coded `.xlsx` workbook with a Summary sheet and a full file-level Details sheet.
- **Export CSV** — a plain-text alternative for scripted processing or import into another tool.

### Scan History

Every completed scan is saved automatically to a history log in your browser (the 10 most recent scans are kept; older ones are dropped automatically). Scroll down to **Scan history** to see past scans by date, with their total size and stale-file count. Click **View** on any entry to reload those results without re-scanning, or the delete icon to remove a saved scan.

### Comparing Reports

Check the boxes next to any **two** saved scans to see a **Comparison** panel showing what changed between them: size change, new archival candidates, resolved (no-longer-stale) items, and the file-count delta. This is the easiest way to track whether a cleanup effort is actually working over time — run a scan on a schedule (e.g. monthly) and compare each new one against the last.

![Scan history with two scans selected and the resulting comparison panel showing size change and new archival candidates](docs/screenshots/06_report_history.png)

---

## Understanding Archival Status

Every file is classified into one of three tiers based on its **last modified date only** — SharePoint does not expose a reliable last-*accessed* signal at scale via REST, so modification date is the most reliable proxy available.

| Status | Meaning |
|--------|---------|
| **Active** | Modified within the "stale" threshold — business as usual |
| **Stale** | Not modified in at least the "stale" threshold (default: 180 days) — worth a look |
| **Very stale** | Not modified in at least the "very stale" threshold (default: 365 days) — a strong archival candidate |

Both thresholds are configurable in [Settings](#settings), and the legend shown on every screen always reflects your current configured values — not the defaults — so what you see on screen is always accurate to your setup.

> **The tool never deletes, moves, or archives anything.** It only reports and helps you browse — what you do with a Stale or Very-stale file (archive it, delete it, or leave it) is entirely up to you.

---

## Settings

Open **Settings** from the gear icon in the top-right corner of the header on any screen.

![The full-page Settings screen showing Scope, Archival thresholds, Performance, and Default view sections](docs/screenshots/07_settings.png)

### Scope

| Setting | Default | Description |
|---------|---------|-------------|
| **Include subsites in Storage Report scans** | Off | Also walks every subsite beneath the current site. Only affects the Storage Report — the Explorer is always scoped to the current site. |
| **Include system and hidden libraries** | Off | Includes Style Library, Form Templates, and other libraries normally hidden from default views. Applies to both tools. |

### Archival Thresholds

| Setting | Default | Description |
|---------|---------|-------------|
| **Stale after (days)** | 180 | Files not modified in this many days are flagged "Stale." |
| **Very stale after (days)** | 365 | Files not modified in this many days are flagged "Very stale." |

### Performance

| Setting | Default | Description |
|---------|---------|-------------|
| **Concurrent API requests** | 6 | How many SharePoint API requests run in parallel during scans and folder loads (1–15). SharePoint's throttling threshold is dynamic, not a fixed number Microsoft publishes — the app retries automatically on throttling (HTTP 429) with a short escalating backoff, but very high values can still net out slower than a moderate one. |

If you see scans getting *slower* rather than faster as you raise this, that's throttling — turn it back down rather than waiting it out.

### Default View on Load

The Settings page also shows step-by-step instructions for changing which screen the web part opens on by default — see [Web Part Configuration](#web-part-configuration) below for the full walkthrough.

---

## Web Part Configuration

Site administrators can configure the web part's default behavior through the SharePoint property pane.

### Setting a Default View

By default, the web part opens on the **Explorer**. You can change this so it opens directly on the **Storage Report** instead — useful if the web part is placed on a page dedicated to periodic reporting.

**To change the default view:**

1. Navigate to the SharePoint page where the web part is installed.
2. Put the page into **Edit** mode.
3. Click the web part, then click its **pencil (edit)** icon.
4. In the property pane, under **General**, use the **Default view on open** dropdown to choose **Explorer** or **Storage Report**.
5. **Republish** the page to save the change.

---

## Security & Privacy

### How It Works

Smart Storage Analyzer runs entirely inside your browser as a SharePoint web part. It makes direct calls to the **SharePoint REST API** using your signed-in credentials — the same API SharePoint itself uses.

### Key Security Properties

| Property | Detail |
|----------|--------|
| **No elevated permissions** | The tool uses only your existing access rights. It cannot see anything you couldn't already see. |
| **Read-only** | The tool never creates, modifies, moves, archives, or deletes any SharePoint content. It only reads and reports. |
| **No external services** | All data stays within your Microsoft 365 tenant. Nothing is sent to any external server or third-party service. |
| **Local-only history** | Saved Storage Report scans are stored in your browser's IndexedDB, not on any server. Clearing browser data removes them. |
| **Standard authentication** | Authentication is handled entirely by SharePoint and Microsoft 365. The web part never handles passwords or tokens directly. |

### What the Tool Can See

The tool can only access sites, libraries, folders, and files that **your account** already has permission to view.

---

## Frequently Asked Questions

**Q: Do I need any special permissions to use this tool?**
A: Yes — you must be a **Site Owner** (or hold the Manage Web / Manage Permissions right). Storage totals rely on the same right as classic Site Settings → Storage Metrics.

---

**Q: Will using this tool change, move, or delete anything?**
A: No. The tool is entirely read-only. It reports and helps you browse storage; it never modifies SharePoint content or settings.

---

**Q: A folder's size shows as an "Estimate" — what does that mean?**
A: SharePoint computes a folder's rolled-up size via a background job (the same one behind classic Site Settings → Storage Metrics), which can lag behind recent activity. When that rollup isn't available or looks stale, the tool falls back to computing the size live by walking the folder's contents directly — accurate, just computed differently.

---

**Q: Why does loading a folder with many subfolders take a moment?**
A: Each subfolder's size is a separate lookup. Sizes are cached for a short time after loading, so revisiting the same folder — even after a page refresh — is fast the second time. Lowering **Concurrent API requests** in Settings can help if you're seeing throttling; raising it can speed up first-time loads on fast tenants.

---

**Q: What's the difference between the Explorer and the Storage Report?**
A: The Explorer is for visual, interactive browsing of one library at a time — great for spotting what's big at a glance. The Storage Report is for a complete, exportable, and comparable record of every file across a whole site (and optionally its subsites), classified by staleness.

---

**Q: How is "stale" determined?**
A: By last-modified date only, compared against the two thresholds you can configure in Settings (default: 180 days for Stale, 365 for Very stale). SharePoint doesn't expose a reliable last-*accessed* signal at scale, so last-modified is the most practical proxy.

---

**Q: Can I compare two scans to see what changed?**
A: Yes. In the Storage Report, check the boxes next to any two saved scans in **Scan history** to see a comparison: size change, new archival candidates, resolved items, and file-count change.

---

**Q: Is my data secure when I use this tool?**
A: Yes. The tool talks only to your own SharePoint environment via the standard Microsoft 365 REST API. No data is sent anywhere else. See [Security & Privacy](#security--privacy).

---

## Troubleshooting

### "Site Owner access required" banner instead of data

- Confirm you have Site Owner access (or the Manage Web / Manage Permissions right) on the site.
- Ask a site owner to grant access, or have them run the scan/export on your behalf.

### A library or folder's size shows as "Estimate"

- This is expected when SharePoint's Storage Metrics rollup isn't yet available for that content. The tool falls back to a live, on-the-fly computation — the number is still accurate, just labeled differently.

### The Storage Report scan takes a long time

- Scan time scales with file count and, if enabled, the number of subsites. Try disabling **Include subsites** or **Include hidden/system libraries** in Settings to narrow the scope, or lower **Concurrent API requests** if you're seeing throttling (HTTP 429) errors in the browser console.

### Export to Excel doesn't seem to do anything

- Try again — a transient network issue loading the export component can occasionally cause a single failed attempt. If it keeps failing, check the browser console for an error message, or try **Export CSV** instead, which has no external dependency.

### The web part shows no data at all in local development

- The local workbench (`localhost:4321`) does not have SharePoint REST API access. Use the **hosted workbench** (your actual SharePoint site's `/_layouts/15/workbench.aspx`) during development instead.

---

*Smart Storage Analyzer is a browser-based utility that runs within your Microsoft 365 environment. It does not store, transmit, or log any data outside of your browser session.*
