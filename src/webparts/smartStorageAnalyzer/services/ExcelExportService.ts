// exceljs is the largest dependency in the bundle and is only needed when the
// user actually exports to .xlsx — load it on demand as a separate webpack
// chunk. Only types are imported statically (erased at compile time).
import type * as ExcelJS from 'exceljs';
import { CandidateTier, FileEntry, FolderListRow, StorageReportSummary } from '../models/models';
import { TIER_COLORS } from '../components/shared/tierBadge';
import { formatBytes } from '../components/shared/formatBytes';

let excelModulePromise: Promise<typeof ExcelJS> | undefined;
function loadExcelJS(): Promise<typeof ExcelJS> {
  if (!excelModulePromise) {
    excelModulePromise = import(/* webpackChunkName: 'exceljs' */ 'exceljs')
      .then((m: any) => (m.default ?? m) as typeof ExcelJS)
      .catch((err) => {
        // A failed dynamic import (e.g. a transient network blip fetching the
        // chunk) would otherwise be cached forever — every future export
        // attempt would immediately re-fail with no way to recover short of
        // reloading the page. Clear the cache so the next call retries.
        excelModulePromise = undefined;
        throw err;
      });
  }
  return excelModulePromise;
}

const COLOR = {
  headerFill: 'FF0078D4',
  headerFont: 'FFFFFFFF',
  titleFont: 'FF0078D4',
};

// Fixed status colors shared with the in-app tier badges (see tierBadge.ts) —
// converted from #rrggbb to Excel's ARGB (alpha-first) hex.
function toArgb(hex: string): string {
  return `FF${hex.replace('#', '').toUpperCase()}`;
}
const TIER_FILL: Record<CandidateTier, string> = {
  [CandidateTier.Active]: toArgb(TIER_COLORS[CandidateTier.Active].light),
  [CandidateTier.Stale]: toArgb(TIER_COLORS[CandidateTier.Stale].light),
  [CandidateTier.VeryStale]: toArgb(TIER_COLORS[CandidateTier.VeryStale].light),
};

function argbFill(hex: string): ExcelJS.Fill {
  return { type: 'pattern', pattern: 'solid', fgColor: { argb: hex } };
}

// Writes a real typed Date (with a date number format) rather than a
// locale-formatted string, so Excel can sort/filter these columns
// chronologically instead of alphabetically. Falls back to an empty cell for
// an unparseable timestamp — exceljs serializes Date values via
// toISOString() internally, which throws on an Invalid Date.
function setExcelDate(cell: ExcelJS.Cell, iso: string): void {
  const d = new Date(iso);
  if (isFinite(d.getTime())) {
    cell.value = d;
    cell.numFmt = 'yyyy-mm-dd';
  } else {
    cell.value = '';
  }
}

// ISO (yyyy-mm-dd) rather than toLocaleDateString() for CSV — sorts
// correctly as plain text and isn't ambiguous across locales (dd/mm vs
// mm/dd). Empty string for an unparseable timestamp.
function isoDate(iso: string): string {
  const d = new Date(iso);
  return isFinite(d.getTime()) ? d.toISOString().slice(0, 10) : '';
}

function timestampSuffix(): string {
  return new Date().toISOString().replace(/[-:T]/g, '').substring(0, 15).replace('.', '');
}

function sanitize(s: string): string {
  return s.replace(/[^a-z0-9]+/gi, '_').replace(/^_+|_+$/g, '').substring(0, 60);
}

// Short, filesystem-safe site identifier for export filenames — e.g.
// "https://tenant.sharepoint.com/sites/IR_BD" -> "IR_BD". Without this, every
// Storage Report export was named identically regardless of which site it
// came from, so telling two exports apart (or which site a downloaded file
// was even for) meant opening it. Falls back to the full URL, sanitized, if
// it isn't a parseable absolute URL — same fallback StorageReportView's own
// siteLabel() uses, for the same reason: a raw string beats nothing.
function siteNameForFilename(siteUrl: string): string {
  try {
    const path = new URL(siteUrl).pathname.replace(/\/+$/, '');
    const segments = path.split('/').filter(Boolean);
    return sanitize(segments[segments.length - 1] || 'Site') || 'Site';
  } catch {
    return sanitize(siteUrl) || 'Site';
  }
}

function downloadBlob(blob: Blob, filename: string): void {
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

export class ExcelExportService {
  async export(entries: FileEntry[], summary: StorageReportSummary, siteUrl: string): Promise<void> {
    const Excel = await loadExcelJS();
    const wb = new Excel.Workbook();
    this.addSummarySheet(wb, summary, siteUrl, entries.length);
    this.addDetailsSheet(wb, entries, !!summary.versionHistoryIncluded);

    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer as ArrayBuffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    downloadBlob(blob, `SP_StorageReport_${siteNameForFilename(siteUrl)}_${timestampSuffix()}.xlsx`);
  }

  private addSummarySheet(
    wb: ExcelJS.Workbook,
    summary: StorageReportSummary,
    siteUrl: string,
    totalEntries: number,
  ): void {
    const ws = wb.addWorksheet('Summary');

    const title = ws.getCell('A1');
    title.value = 'SharePoint Storage Report';
    title.font = { bold: true, size: 16, color: { argb: COLOR.titleFont } };

    const data: [string, string | number][] = [
      ['Site URL', siteUrl],
      ['Generated', new Date().toLocaleString()],
      ['Scan duration (s)', summary.durationSeconds.toFixed(1)],
      ['Files scanned', totalEntries],
      ['Total size', formatBytes(summary.totalSizeBytes)],
      ['Stale files (Stale)', summary.staleCount],
      ['Stale size', formatBytes(summary.staleSizeBytes)],
      ['Very stale files (strong archival candidates)', summary.veryStaleCount],
      ['Very stale size', formatBytes(summary.veryStaleSizeBytes)],
      // Additive to Total size, not a subset of it — see the in-app info
      // tooltip on this same figure (StorageReportView). Only added when the
      // scan actually measured it; totalVersionSizeBytes === 0 is otherwise
      // ambiguous between "no old versions exist" and "wasn't measured".
      ...(summary.versionHistoryIncluded
        ? ([['Version History Size (additive to Total size)', formatBytes(summary.totalVersionSizeBytes ?? 0)]] as [string, string][])
        : []),
    ];

    data.forEach(([label, value], i) => {
      const row = i + 3;
      const labelCell = ws.getCell(row, 1);
      labelCell.value = label;
      labelCell.font = { bold: true };
      ws.getCell(row, 2).value = value;
    });

    ws.getColumn(1).width = 34;
    ws.getColumn(2).width = 50;
  }

  private addDetailsSheet(wb: ExcelJS.Workbook, entries: FileEntry[], includeVersionHistory: boolean): void {
    const ws = wb.addWorksheet('Files');

    // Version history column only added when the scan actually measured it —
    // otherwise every entry's versionSizeBytes is undefined and the column
    // would just be blank, misleadingly implying "confirmed zero".
    const headers = ['Library', 'Path', 'Name', 'Size', 'Size (bytes)'];
    if (includeVersionHistory) headers.push('Version History Size', 'Version History Size (bytes)', 'Version Count (est.)');
    headers.push('Created', 'Modified', 'Age (days)', 'Author', 'Tier');
    const headerRow = ws.getRow(1);
    headers.forEach((h, i) => {
      const cell = headerRow.getCell(i + 1);
      cell.value = h;
      cell.fill = argbFill(COLOR.headerFill);
      cell.font = { bold: true, color: { argb: COLOR.headerFont } };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
    });
    headerRow.commit();
    ws.views = [{ state: 'frozen', xSplit: 0, ySplit: 1 }];

    entries.forEach((entry, idx) => {
      const row = ws.getRow(idx + 2);
      let col = 1;
      row.getCell(col++).value = entry.libraryTitle;
      row.getCell(col++).value = entry.serverRelativeUrl;
      row.getCell(col++).value = entry.name;
      row.getCell(col++).value = formatBytes(entry.sizeBytes);
      row.getCell(col++).value = entry.sizeBytes;
      if (includeVersionHistory) {
        row.getCell(col++).value = entry.versionSizeBytes != null ? formatBytes(entry.versionSizeBytes) : '';
        row.getCell(col++).value = entry.versionSizeBytes ?? '';
        row.getCell(col++).value = entry.versionCount ?? '';
      }
      setExcelDate(row.getCell(col++), entry.timeCreated);
      setExcelDate(row.getCell(col++), entry.timeLastModified);
      row.getCell(col++).value = entry.ageDays;
      row.getCell(col++).value = entry.authorDisplayName ?? '';

      const tierCell = row.getCell(col++);
      tierCell.value = entry.tier;
      tierCell.fill = argbFill(TIER_FILL[entry.tier]);
      tierCell.alignment = { horizontal: 'center', vertical: 'middle' };

      row.commit();
    });

    const widths = [24, 60, 30, 12, 14];
    if (includeVersionHistory) widths.push(14, 16, 12);
    widths.push(14, 14, 12, 24, 12);
    widths.forEach((w, i) => { ws.getColumn(i + 1).width = w; });
    const lastCol = String.fromCharCode('A'.charCodeAt(0) + headers.length - 1);
    ws.autoFilter = { from: 'A1', to: `${lastCol}1` };
  }

  // ── CSV export ────────────────────────────────────────────────────────────

  private csvEscape(v: string | number): string {
    let s = String(v);
    // CSV formula injection guard (OWASP-recommended mitigation): a leading
    // =, +, -, @, tab, or CR makes Excel/Sheets interpret the cell as a
    // formula when the file is opened — file/author names are
    // attacker-influenceable in a shared tenant. Prefixing with a single
    // quote forces text interpretation; Excel does not display the quote
    // itself. Applied unconditionally rather than only to "text" columns —
    // no current column emits a leading '-' from a legitimate value (sizes/
    // ages are always >= 0), so there's nothing to special-case.
    if (/^[=+\-@\t\r]/.test(s)) {
      s = `'${s}`;
    }
    return s.includes(',') || s.includes('"') || s.includes('\n')
      ? `"${s.replace(/"/g, '""')}"`
      : s;
  }

  private downloadCsv(rows: string[][], filename: string): void {
    const content = rows.map((r) => r.map((c) => this.csvEscape(c)).join(',')).join('\r\n');
    const blob = new Blob(['﻿' + content], { type: 'text/csv;charset=utf-8;' });
    downloadBlob(blob, filename);
  }

  exportCsv(entries: FileEntry[], includeVersionHistory = false, siteUrl = ''): void {
    const header = ['Library', 'Path', 'Name', 'Size (bytes)'];
    if (includeVersionHistory) header.push('Version History Size (bytes)', 'Version Count (est.)');
    header.push('Created', 'Modified', 'Age (days)', 'Author', 'Tier');
    const rows: string[][] = [header];
    for (const e of entries) {
      const row = [e.libraryTitle, e.serverRelativeUrl, e.name, String(e.sizeBytes)];
      if (includeVersionHistory) {
        row.push(e.versionSizeBytes != null ? String(e.versionSizeBytes) : '', e.versionCount != null ? String(e.versionCount) : '');
      }
      row.push(isoDate(e.timeCreated), isoDate(e.timeLastModified), String(e.ageDays), e.authorDisplayName ?? '', e.tier);
      rows.push(row);
    }
    this.downloadCsv(rows, `SP_StorageReport_${siteNameForFilename(siteUrl)}_${timestampSuffix()}.csv`);
  }

  // ── Explorer List View export ───────────────────────────────────────────────

  async exportFolderListing(rows: FolderListRow[], contextLabel: string, includeVersionHistory = false): Promise<void> {
    const Excel = await loadExcelJS();
    const wb = new Excel.Workbook();
    const ws = wb.addWorksheet('Listing');

    const title = ws.getCell('A1');
    title.value = `Folder Listing — ${contextLabel}`;
    title.font = { bold: true, size: 14, color: { argb: COLOR.titleFont } };

    // Folders never get a real value here (no recursive version-history
    // rollup exists — see FolderListRow.versionSizeBytes) and render blank,
    // same as the in-app List view's '—'.
    const headers = ['Type', 'Name', 'Size', 'Size (bytes)'];
    if (includeVersionHistory) headers.push('Version History Size', 'Version History Size (bytes)', 'Version Count (est.)');
    headers.push('Items', 'Modified', 'Age (days)', 'Author', 'Status');
    const headerRow = ws.getRow(3);
    headers.forEach((h, i) => {
      const cell = headerRow.getCell(i + 1);
      cell.value = h;
      cell.fill = argbFill(COLOR.headerFill);
      cell.font = { bold: true, color: { argb: COLOR.headerFont } };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
    });
    headerRow.commit();
    ws.views = [{ state: 'frozen', xSplit: 0, ySplit: 3 }];

    rows.forEach((r, idx) => {
      const row = ws.getRow(idx + 4);
      let col = 1;
      row.getCell(col++).value = r.kind === 'folder' ? 'Folder' : 'File';
      row.getCell(col++).value = r.name;
      row.getCell(col++).value = formatBytes(r.sizeBytes);
      row.getCell(col++).value = r.sizeBytes;
      if (includeVersionHistory) {
        const hasVersion = r.kind === 'file' && r.versionSizeBytes != null;
        row.getCell(col++).value = hasVersion ? formatBytes(r.versionSizeBytes!) : '';
        row.getCell(col++).value = hasVersion ? r.versionSizeBytes! : '';
        row.getCell(col++).value = r.kind === 'file' && r.versionCount != null ? r.versionCount : '';
      }
      row.getCell(col++).value = r.itemCount ?? '';
      const modifiedCell = row.getCell(col++);
      if (r.lastModified) setExcelDate(modifiedCell, r.lastModified); else modifiedCell.value = '';
      row.getCell(col++).value = r.ageDays ?? '';
      row.getCell(col++).value = r.authorDisplayName ?? '';
      if (r.kind === 'file' && r.tier) {
        const tierCell = row.getCell(col++);
        tierCell.value = r.tier;
        tierCell.fill = argbFill(TIER_FILL[r.tier]);
        tierCell.alignment = { horizontal: 'center', vertical: 'middle' };
      }
      row.commit();
    });

    const widths = [10, 34, 12, 14];
    if (includeVersionHistory) widths.push(14, 16, 12);
    widths.push(10, 14, 12, 24, 12);
    widths.forEach((w, i) => { ws.getColumn(i + 1).width = w; });
    const lastCol = String.fromCharCode('A'.charCodeAt(0) + headers.length - 1);
    ws.autoFilter = { from: 'A3', to: `${lastCol}3` };

    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer as ArrayBuffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    downloadBlob(blob, `SP_FolderListing_${sanitize(contextLabel)}_${timestampSuffix()}.xlsx`);
  }

  exportFolderListingCsv(rows: FolderListRow[], contextLabel: string, includeVersionHistory = false): void {
    const header = ['Type', 'Name', 'Size (bytes)'];
    if (includeVersionHistory) header.push('Version History Size (bytes)', 'Version Count (est.)');
    header.push('Items', 'Modified', 'Age (days)', 'Author', 'Status');
    const out: string[][] = [header];
    for (const r of rows) {
      const row = [r.kind === 'folder' ? 'Folder' : 'File', r.name, String(r.sizeBytes)];
      if (includeVersionHistory) {
        row.push(
          r.kind === 'file' && r.versionSizeBytes != null ? String(r.versionSizeBytes) : '',
          r.kind === 'file' && r.versionCount != null ? String(r.versionCount) : '',
        );
      }
      row.push(
        r.itemCount != null ? String(r.itemCount) : '',
        r.lastModified ? isoDate(r.lastModified) : '',
        r.ageDays != null ? String(r.ageDays) : '',
        r.authorDisplayName ?? '',
        r.tier ?? '',
      );
      out.push(row);
    }
    this.downloadCsv(out, `SP_FolderListing_${sanitize(contextLabel)}_${timestampSuffix()}.csv`);
  }
}
