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
    this.addDetailsSheet(wb, entries);

    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer as ArrayBuffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    downloadBlob(blob, `SP_StorageReport_${timestampSuffix()}.xlsx`);
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

  private addDetailsSheet(wb: ExcelJS.Workbook, entries: FileEntry[]): void {
    const ws = wb.addWorksheet('Files');

    const headers = ['Library', 'Path', 'Name', 'Size', 'Size (bytes)', 'Created', 'Modified', 'Age (days)', 'Author', 'Tier'];
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
      row.getCell(1).value = entry.libraryTitle;
      row.getCell(2).value = entry.serverRelativeUrl;
      row.getCell(3).value = entry.name;
      row.getCell(4).value = formatBytes(entry.sizeBytes);
      row.getCell(5).value = entry.sizeBytes;
      setExcelDate(row.getCell(6), entry.timeCreated);
      setExcelDate(row.getCell(7), entry.timeLastModified);
      row.getCell(8).value = entry.ageDays;
      row.getCell(9).value = entry.authorDisplayName ?? '';

      const tierCell = row.getCell(10);
      tierCell.value = entry.tier;
      tierCell.fill = argbFill(TIER_FILL[entry.tier]);
      tierCell.alignment = { horizontal: 'center', vertical: 'middle' };

      row.commit();
    });

    ws.getColumn(1).width = 24;
    ws.getColumn(2).width = 60;
    ws.getColumn(3).width = 30;
    ws.getColumn(4).width = 12;
    ws.getColumn(5).width = 14;
    ws.getColumn(6).width = 14;
    ws.getColumn(7).width = 14;
    ws.getColumn(8).width = 12;
    ws.getColumn(9).width = 24;
    ws.getColumn(10).width = 12;
    ws.autoFilter = { from: 'A1', to: 'J1' };
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

  exportCsv(entries: FileEntry[]): void {
    const rows: string[][] = [
      ['Library', 'Path', 'Name', 'Size (bytes)', 'Created', 'Modified', 'Age (days)', 'Author', 'Tier'],
    ];
    for (const e of entries) {
      rows.push([
        e.libraryTitle,
        e.serverRelativeUrl,
        e.name,
        String(e.sizeBytes),
        isoDate(e.timeCreated),
        isoDate(e.timeLastModified),
        String(e.ageDays),
        e.authorDisplayName ?? '',
        e.tier,
      ]);
    }
    this.downloadCsv(rows, `SP_StorageReport_${timestampSuffix()}.csv`);
  }

  // ── Explorer List View export ───────────────────────────────────────────────

  async exportFolderListing(rows: FolderListRow[], contextLabel: string): Promise<void> {
    const Excel = await loadExcelJS();
    const wb = new Excel.Workbook();
    const ws = wb.addWorksheet('Listing');

    const title = ws.getCell('A1');
    title.value = `Folder Listing — ${contextLabel}`;
    title.font = { bold: true, size: 14, color: { argb: COLOR.titleFont } };

    const headers = ['Type', 'Name', 'Size', 'Size (bytes)', 'Items', 'Modified', 'Age (days)', 'Author', 'Status'];
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
      row.getCell(1).value = r.kind === 'folder' ? 'Folder' : 'File';
      row.getCell(2).value = r.name;
      row.getCell(3).value = formatBytes(r.sizeBytes);
      row.getCell(4).value = r.sizeBytes;
      row.getCell(5).value = r.itemCount ?? '';
      if (r.lastModified) setExcelDate(row.getCell(6), r.lastModified); else row.getCell(6).value = '';
      row.getCell(7).value = r.ageDays ?? '';
      row.getCell(8).value = r.authorDisplayName ?? '';
      if (r.kind === 'file' && r.tier) {
        const tierCell = row.getCell(9);
        tierCell.value = r.tier;
        tierCell.fill = argbFill(TIER_FILL[r.tier]);
        tierCell.alignment = { horizontal: 'center', vertical: 'middle' };
      }
      row.commit();
    });

    ws.getColumn(1).width = 10;
    ws.getColumn(2).width = 34;
    ws.getColumn(3).width = 12;
    ws.getColumn(4).width = 14;
    ws.getColumn(5).width = 10;
    ws.getColumn(6).width = 14;
    ws.getColumn(7).width = 12;
    ws.getColumn(8).width = 24;
    ws.getColumn(9).width = 12;
    ws.autoFilter = { from: 'A3', to: 'I3' };

    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer as ArrayBuffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    downloadBlob(blob, `SP_FolderListing_${sanitize(contextLabel)}_${timestampSuffix()}.xlsx`);
  }

  exportFolderListingCsv(rows: FolderListRow[], contextLabel: string): void {
    const out: string[][] = [
      ['Type', 'Name', 'Size (bytes)', 'Items', 'Modified', 'Age (days)', 'Author', 'Status'],
    ];
    for (const r of rows) {
      out.push([
        r.kind === 'folder' ? 'Folder' : 'File',
        r.name,
        String(r.sizeBytes),
        r.itemCount != null ? String(r.itemCount) : '',
        r.lastModified ? isoDate(r.lastModified) : '',
        r.ageDays != null ? String(r.ageDays) : '',
        r.authorDisplayName ?? '',
        r.tier ?? '',
      ]);
    }
    this.downloadCsv(out, `SP_FolderListing_${sanitize(contextLabel)}_${timestampSuffix()}.csv`);
  }
}
