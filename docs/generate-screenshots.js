'use strict';
const puppeteer = require('puppeteer-core');
const path = require('path');
const fs = require('fs');

// Chrome path — fall back to Edge if Chrome not present
const CHROME = fs.existsSync('C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe')
  ? 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe'
  : 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe';

const OUT = path.join(__dirname, 'screenshots');
if (!fs.existsSync(OUT)) fs.mkdirSync(OUT, { recursive: true });

// ── Shared visual language (matches the real Fluent UI v9 app) ────────────────
const BRAND = '#0078d4';
const FOLDER_COLOR = '#2a78d6';
const TIER = { active: '#0ca30c', stale: '#fab219', veryStale: '#d03b3b' };
const NEUTRAL = {
  bg: '#faf9f8',
  surface: '#ffffff',
  border: '#e0e0e0',
  text1: '#242424',
  text2: '#616161',
  text3: '#8a8886',
  tileBg: '#f5f5f5',
};
const FONT = "'Segoe UI', -apple-system, Arial, sans-serif";

const HARDDRIVE_SVG = '<path d="M5.42 5.27A2.25 2.25 0 0 1 7.44 4h9.12c.86 0 1.65.5 2.02 1.27l3.1 6.39c.2.44.32.92.32 1.41v3.68c0 1.24-1 2.25-2.25 2.25H4.25C3.01 19 2 18 2 16.75v-3.68c0-.49.11-.97.32-1.41l3.1-6.39Zm11.81.65a.75.75 0 0 0-.67-.42H7.44a.75.75 0 0 0-.67.42L4.3 11h15.38l-2.46-5.08ZM3.5 13.25v3.5c0 .41.34.75.75.75h15.5c.41 0 .75-.34.75-.75v-3.5a.75.75 0 0 0-.75-.75H4.25a.75.75 0 0 0-.75.75ZM18 16a1 1 0 1 0 0-2 1 1 0 0 0 0 2Z"/>';
const GEAR_SVG = '<path d="M12.01 2.25c.74 0 1.47.1 2.18.25.32.07.55.33.59.65l.17 1.53a1.38 1.38 0 0 0 1.92 1.11l1.4-.61c.3-.13.64-.06.85.17a9.8 9.8 0 0 1 2.2 3.8c.1.3 0 .63-.26.82l-1.25.92a1.38 1.38 0 0 0 0 2.22l1.25.92c.26.19.36.52.27.82a9.8 9.8 0 0 1-2.2 3.8.75.75 0 0 1-.85.17l-1.4-.62a1.38 1.38 0 0 0-1.93 1.12l-.17 1.52a.75.75 0 0 1-.58.65 9.52 9.52 0 0 1-4.4 0 .75.75 0 0 1-.57-.65l-.17-1.52a1.38 1.38 0 0 0-1.93-1.11l-1.4.62a.75.75 0 0 1-.85-.18 9.8 9.8 0 0 1-2.2-3.8c-.1-.3 0-.63.26-.82l1.25-.92a1.38 1.38 0 0 0 0-2.22l-1.24-.92a.75.75 0 0 1-.28-.82 9.8 9.8 0 0 1 2.2-3.8c.23-.23.57-.3.86-.17l1.4.62c.4.17.86.15 1.25-.08.38-.22.63-.6.68-1.04l.17-1.53a.75.75 0 0 1 .58-.65c.72-.16 1.45-.24 2.2-.25Zm0 1.5c-.45 0-.9.04-1.35.12l-.11.97a2.89 2.89 0 0 1-4.03 2.33l-.9-.4A8.3 8.3 0 0 0 4.29 9.1l.8.59a2.88 2.88 0 0 1 0 4.64l-.8.59a8.3 8.3 0 0 0 1.35 2.32l.9-.4a2.88 2.88 0 0 1 4.02 2.32l.1.99c.9.15 1.8.15 2.7 0l.1-.99a2.88 2.88 0 0 1 4.02-2.32l.9.4a8.3 8.3 0 0 0 1.35-2.32l-.8-.59a2.88 2.88 0 0 1 0-4.64l.8-.59a8.3 8.3 0 0 0-1.35-2.32l-.9.4a2.88 2.88 0 0 1-4.02-2.32l-.1-.98c-.45-.08-.9-.11-1.34-.12ZM12 8.25a3.75 3.75 0 1 1 0 7.5 3.75 3.75 0 0 1 0-7.5Zm0 1.5a2.25 2.25 0 1 0 0 4.5 2.25 2.25 0 0 0 0-4.5Z"/>';
const CHEVRON_SVG = '<path d="M8.7 5.3a1 1 0 0 0 0 1.4L13 11l-4.3 4.3a1 1 0 1 0 1.4 1.4l5-5a1 1 0 0 0 0-1.4l-5-5a1 1 0 0 0-1.4 0Z"/>';
const DOWNLOAD_SVG = '<path d="M12 3a1 1 0 0 1 1 1v9.6l3.3-3.3a1 1 0 1 1 1.4 1.4l-5 5a1 1 0 0 1-1.4 0l-5-5a1 1 0 1 1 1.4-1.4L11 13.6V4a1 1 0 0 1 1-1ZM5 19a1 1 0 0 0 0 2h14a1 1 0 0 0 0-2H5Z"/>';
const FOLDER_SVG = '<path d="M4.5 5A1.5 1.5 0 0 0 3 6.5v11A1.5 1.5 0 0 0 4.5 19h15a1.5 1.5 0 0 0 1.5-1.5V8.5A1.5 1.5 0 0 0 19.5 7h-7.4l-1.5-1.5A1.5 1.5 0 0 0 9.5 5h-5Z"/>';
const DOC_SVG = '<path d="M7 2.5A1.5 1.5 0 0 0 5.5 4v16A1.5 1.5 0 0 0 7 21.5h10a1.5 1.5 0 0 0 1.5-1.5V8.6a1.5 1.5 0 0 0-.44-1.06l-5.1-5.1A1.5 1.5 0 0 0 11.9 2H7ZM13 3.4l4.6 4.6H13.75a.75.75 0 0 1-.75-.75V3.4Z"/>';
const HISTORY_SVG = '<path d="M12 4a8 8 0 1 0 6.93 4H16.9a6 6 0 1 1-1.5-2.6l-1.9 1.9H19V2l-2.1 2.1A7.96 7.96 0 0 0 12 4Zm-.75 3.5v4.44l3.02 1.75a.75.75 0 1 0 .75-1.3l-2.27-1.31V7.5a.75.75 0 0 0-1.5 0Z"/>';
const DELETE_SVG = '<path d="M9.5 3a1 1 0 0 0-1 1v1H5a1 1 0 0 0 0 2h.5l.9 12.1A2 2 0 0 0 8.4 21h7.2a2 2 0 0 0 2-1.9L18.5 7H19a1 1 0 1 0 0-2h-3.5V4a1 1 0 0 0-1-1h-5Zm.5 6a.75.75 0 0 1 1.5 0v8a.75.75 0 0 1-1.5 0V9Zm4.5 0a.75.75 0 0 1 1.5 0v8a.75.75 0 0 1-1.5 0V9Z"/>';
const INFO_SVG = '<path d="M12 2a10 10 0 1 0 0 20 10 10 0 0 0 0-20Zm0 6.5a1 1 0 1 1 0 2 1 1 0 0 1 0-2ZM11 11h1.5a.5.5 0 0 1 .5.5V17h1a1 1 0 1 1 0 2h-3.5a1 1 0 1 1 0-2h1v-4h-.5a1 1 0 1 1 0-2Z"/>';
const BACK_SVG = '<path d="M15.7 5.3a1 1 0 0 1 0 1.4L10.4 12l5.3 5.3a1 1 0 0 1-1.4 1.4l-6-6a1 1 0 0 1 0-1.4l6-6a1 1 0 0 1 1.4 0Z"/>';

function icon(svg, size, color) {
  return `<svg width="${size}" height="${size}" viewBox="0 0 24 24" fill="${color}" style="flex-shrink:0;">${svg}</svg>`;
}

function header(activeTab) {
  const tabBtn = (label, active) => `
    <button style="
      font-family:${FONT};font-size:13px;font-weight:600;padding:6px 12px;border-radius:4px;border:none;cursor:default;
      background:${active ? '#ffffff' : 'transparent'};color:${active ? BRAND : '#ffffff'};
    ">${label}</button>`;
  return `
  <div style="background:${BRAND};padding:10px 16px;display:flex;align-items:center;justify-content:space-between;gap:16px;">
    <div style="display:flex;align-items:center;gap:8px;">
      ${icon(HARDDRIVE_SVG, 20, '#ffffff')}
      <span style="color:#fff;font-family:${FONT};font-weight:600;font-size:14px;white-space:nowrap;">SharePoint Smart Storage Analyzer</span>
    </div>
    <div style="display:flex;align-items:center;gap:8px;">
      ${tabBtn('Explorer', activeTab === 'explorer')}
      ${tabBtn('Storage Report', activeTab === 'report')}
      <button style="background:transparent;border:none;padding:6px;cursor:default;">${icon(GEAR_SVG, 18, '#ffffff')}</button>
    </div>
  </div>`;
}

function pageShell(bodyHtml, opts) {
  opts = opts || {};
  return `<!DOCTYPE html><html><head><meta charset="utf-8">
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: ${FONT}; background: ${NEUTRAL.bg}; color: ${NEUTRAL.text1}; -webkit-font-smoothing: antialiased; }
    table { border-collapse: collapse; width: 100%; font-size: 13px; }
    th { text-align: left; padding: 6px 10px; border-bottom: 2px solid ${NEUTRAL.border}; font-weight: 600; color: ${NEUTRAL.text2}; font-size: 12px; }
    td { padding: 7px 10px; border-bottom: 1px solid #edebe9; vertical-align: middle; }
    .btn { font-family: ${FONT}; font-size: 13px; font-weight: 600; padding: 6px 14px; border-radius: 4px; border: 1px solid ${NEUTRAL.border}; background: #fff; color: ${NEUTRAL.text1}; display: inline-flex; align-items: center; gap: 6px; cursor: default; }
    .btn.primary { background: ${BRAND}; border-color: ${BRAND}; color: #fff; }
    .checkbox { display: inline-flex; align-items: center; gap: 6px; font-size: 13px; color: ${NEUTRAL.text1}; }
    .checkbox .box { width: 16px; height: 16px; border: 1.5px solid #8a8886; border-radius: 3px; display: inline-block; }
    .checkbox .box.checked { background: ${BRAND}; border-color: ${BRAND}; position: relative; }
  </style></head>
  <body>
    ${opts.noHeader ? '' : header(opts.tab)}
    <div style="max-width:${opts.maxWidth || '1200px'};margin:0 auto;padding:24px;">
      ${bodyHtml}
    </div>
  </body></html>`;
}

function tierLegend(showFolder) {
  const item = (color, label) => `
    <div style="display:flex;align-items:center;gap:5px;">
      <span style="width:10px;height:10px;border-radius:2px;background:${color};display:inline-block;"></span>
      <span style="font-size:12px;color:${NEUTRAL.text2};">${label}</span>
    </div>`;
  return `<div style="display:flex;gap:16px;flex-wrap:wrap;padding:6px 0 16px;">
    ${showFolder ? item(FOLDER_COLOR, 'Folder') : ''}
    ${item(TIER.active, 'Active (&lt; 180d)')}
    ${item(TIER.stale, 'Stale (180–364d)')}
    ${item(TIER.veryStale, 'Very stale (365d+)')}
  </div>`;
}

// ── 1. Explorer — Treemap view ────────────────────────────────────────────────
function explorerTreemapPage() {
  const cells = [
    { l: 0, t: 0, w: 44, h: 62, color: FOLDER_COLOR, name: 'Projects', sub: '4.2 GB · 1,204 files' },
    { l: 44, t: 0, w: 26, h: 62, color: FOLDER_COLOR, name: 'Archive', sub: '2.8 GB · 890 files' },
    { l: 70, t: 0, w: 30, h: 62, color: TIER.veryStale, name: 'Video-Raw-Footage.mp4', sub: '' },
    { l: 0, t: 62, w: 22, h: 38, color: FOLDER_COLOR, name: 'Old Reports', sub: '1.1 GB · 210 files' },
    { l: 22, t: 62, w: 18, h: 38, color: FOLDER_COLOR, name: 'Templates', sub: '340 MB · 156 files' },
    { l: 40, t: 62, w: 15, h: 38, color: TIER.stale, name: 'Q4-Budget.xlsx', sub: '' },
    { l: 55, t: 62, w: 22, h: 38, color: TIER.active, name: 'Brand-Guidelines.pdf', sub: '' },
    { l: 77, t: 62, w: 23, h: 38, color: TIER.stale, name: 'Vendor-Contracts.zip', sub: '' },
  ];
  const cellsHtml = cells.map((c) => `
    <div style="position:absolute;left:${c.l}%;top:${c.t}%;width:${c.w}%;height:${c.h}%;background:${c.color};border:1px solid ${NEUTRAL.bg};box-sizing:border-box;padding:5px 7px;overflow:hidden;">
      <div style="color:#fff;font-size:12px;font-weight:600;text-shadow:0 1px 2px rgba(0,0,0,0.5);white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${c.name}</div>
      ${c.sub ? `<div style="color:#fff;font-size:10.5px;opacity:0.9;text-shadow:0 1px 2px rgba(0,0,0,0.5);margin-top:2px;">${c.sub}</div>` : ''}
    </div>`).join('');

  const body = `
    <div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:16px;">
      <button class="btn primary">Documents</button>
      <button class="btn">Site Assets</button>
      <button class="btn">Campaign Archive</button>
    </div>
    <div style="display:flex;align-items:center;gap:2px;margin-bottom:12px;font-size:13px;color:${BRAND};">
      <span style="font-weight:600;">Documents</span>
    </div>
    <div style="display:flex;gap:0;border-bottom:2px solid ${BRAND};width:fit-content;margin-bottom:4px;">
      <div style="padding:8px 16px;font-size:13px;font-weight:600;color:${BRAND};border-bottom:2px solid ${BRAND};margin-bottom:-2px;">Treemap</div>
      <div style="padding:8px 16px;font-size:13px;color:${NEUTRAL.text2};">List</div>
    </div>
    ${tierLegend(true)}
    <div style="position:relative;width:100%;height:420px;background:${NEUTRAL.tileBg};">
      ${cellsHtml}
    </div>`;
  return pageShell(body, { tab: 'explorer' });
}

// ── 2. Explorer — List view ──────────────────────────────────────────────────
function explorerListPage() {
  const rows = [
    { icon: 'folder', name: 'Projects', size: '4.2 GB', items: '1,204', mod: '7/1/2026', status: null },
    { icon: 'folder', name: 'Archive', size: '2.8 GB', items: '890', mod: '6/2/2026', status: null },
    { icon: 'doc', name: 'Video-Raw-Footage.mp4', size: '1.9 GB', items: '', mod: '3/14/2025', status: 'veryStale' },
    { icon: 'folder', name: 'Old Reports', size: '1.1 GB', items: '210', mod: '1/9/2025', status: null },
    { icon: 'folder', name: 'Templates', size: '340 MB', items: '156', mod: '6/28/2026', status: null },
    { icon: 'doc', name: 'Vendor-Contracts.zip', size: '210 MB', items: '', mod: '9/2/2025', status: 'stale' },
    { icon: 'doc', name: 'Q4-Budget.xlsx', size: '18 MB', items: '', mod: '11/20/2025', status: 'stale' },
    { icon: 'doc', name: 'Brand-Guidelines.pdf', size: '9 MB', items: '', mod: '6/30/2026', status: 'active' },
  ];
  const statusDot = (s) => s ? `<span style="display:inline-flex;align-items:center;gap:5px;"><span style="width:8px;height:8px;border-radius:50%;background:${TIER[s]};display:inline-block;"></span>${s === 'veryStale' ? 'Very stale' : s === 'stale' ? 'Stale' : 'Active'}</span>` : '<span style="color:#c8c6c4;">—</span>';
  const rowsHtml = rows.map((r) => `
    <tr>
      <td><div style="display:flex;align-items:center;gap:6px;">${icon(r.icon === 'folder' ? FOLDER_SVG : DOC_SVG, 14, r.icon === 'folder' ? BRAND : NEUTRAL.text2)}<span>${r.name}</span></div></td>
      <td style="text-align:right;">${r.size}</td>
      <td style="text-align:right;">${r.items || ''}</td>
      <td>${r.mod}</td>
      <td>${statusDot(r.status)}</td>
    </tr>`).join('');

  const body = `
    <div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:16px;">
      <button class="btn primary">Documents</button>
      <button class="btn">Site Assets</button>
      <button class="btn">Campaign Archive</button>
    </div>
    <div style="display:flex;align-items:center;gap:2px;margin-bottom:12px;font-size:13px;">
      <span style="font-weight:600;color:${BRAND};">Documents</span>
    </div>
    <div style="display:flex;gap:0;border-bottom:1px solid ${NEUTRAL.border};width:fit-content;margin-bottom:16px;">
      <div style="padding:8px 16px;font-size:13px;color:${NEUTRAL.text2};">Treemap</div>
      <div style="padding:8px 16px;font-size:13px;font-weight:600;color:${BRAND};border-bottom:2px solid ${BRAND};margin-bottom:-1px;">List</div>
    </div>
    <div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px;margin-bottom:8px;">
      ${tierLegend(false)}
      <div style="display:flex;gap:8px;">
        <button class="btn">${icon(DOWNLOAD_SVG, 14, NEUTRAL.text1)} Export Excel</button>
        <button class="btn">${icon(DOWNLOAD_SVG, 14, NEUTRAL.text1)} Export CSV</button>
      </div>
    </div>
    <table style="background:#fff;">
      <thead><tr><th>Name</th><th style="text-align:right;">Size</th><th style="text-align:right;">Items</th><th>Modified</th><th>Status</th></tr></thead>
      <tbody>${rowsHtml}</tbody>
    </table>`;
  return pageShell(body, { tab: 'explorer' });
}

// ── 3. Storage Report — before running ───────────────────────────────────────
function reportConfigPage() {
  const body = `
    <div style="display:flex;align-items:center;gap:10px;margin-bottom:20px;">
      ${icon(BACK_SVG, 18, NEUTRAL.text1)}
      <span style="font-size:20px;font-weight:600;">Storage Report</span>
    </div>
    <div style="display:flex;align-items:center;gap:16px;flex-wrap:wrap;">
      <span class="checkbox"><span class="box"></span>Include subsites</span>
      <span class="checkbox"><span class="box"></span>Include hidden/system libraries</span>
      <button class="btn primary">Run scan</button>
    </div>`;
  return pageShell(body, { tab: 'report', maxWidth: '1100px' });
}

// ── 4. Storage Report — scanning in progress ─────────────────────────────────
function reportRunningPage() {
  const body = `
    <div style="display:flex;align-items:center;gap:10px;margin-bottom:20px;">
      ${icon(BACK_SVG, 18, NEUTRAL.text1)}
      <span style="font-size:20px;font-weight:600;">Storage Report</span>
    </div>
    <div style="display:flex;align-items:center;gap:16px;flex-wrap:wrap;margin-bottom:24px;">
      <span class="checkbox"><span class="box checked"></span>Include subsites</span>
      <span class="checkbox"><span class="box"></span>Include hidden/system libraries</span>
      <button class="btn primary" style="opacity:0.6;">Scanning…</button>
    </div>
    <div style="margin-bottom:24px;">
      <div style="height:4px;background:#edebe9;border-radius:2px;overflow:hidden;width:100%;">
        <div style="height:100%;width:64%;background:${BRAND};"></div>
      </div>
      <div style="font-size:12px;color:${NEUTRAL.text2};margin-top:6px;">
        Scanning Campaign Archive… — 8,412 files scanned — 34s elapsed
      </div>
    </div>`;
  return pageShell(body, { tab: 'report', maxWidth: '1100px' });
}

// ── 5. Storage Report — results ──────────────────────────────────────────────
function reportResultsPage() {
  const tile = (label, value) => `
    <div style="background:${NEUTRAL.tileBg};border-radius:4px;padding:14px 16px;">
      <div style="font-size:12px;color:${NEUTRAL.text2};">${label}</div>
      <div style="font-size:22px;font-weight:600;margin-top:2px;">${value}</div>
    </div>`;
  const rows = [
    ['Campaign Archive', 'Q3-Launch-Master.mov', '2.4 GB', '3/2/2025', 'A. Rivera', 'veryStale'],
    ['Documents', 'Video-Raw-Footage.mp4', '1.9 GB', '3/14/2025', 'S. Chen', 'veryStale'],
    ['Documents', 'Vendor-Contracts.zip', '210 MB', '9/2/2025', 'M. Patel', 'stale'],
    ['Site Assets', 'Homepage-Banner-v3.psd', '180 MB', '10/5/2025', 'A. Rivera', 'stale'],
    ['Documents', 'Q4-Budget.xlsx', '18 MB', '11/20/2025', 'S. Chen', 'stale'],
    ['Documents', 'Brand-Guidelines.pdf', '9 MB', '6/30/2026', 'M. Patel', 'active'],
  ];
  const statusChip = (s) => `<span style="display:inline-flex;align-items:center;gap:5px;"><span style="width:8px;height:8px;border-radius:50%;background:${TIER[s]};display:inline-block;"></span>${s === 'veryStale' ? 'Very stale' : s === 'stale' ? 'Stale' : 'Active'}</span>`;
  const rowsHtml = rows.map((r) => `
    <tr>
      <td>${r[0]}</td>
      <td><div style="display:flex;align-items:center;gap:6px;">${icon(DOC_SVG, 14, NEUTRAL.text2)}${r[1]}</div></td>
      <td style="text-align:right;">${r[2]}</td>
      <td>${r[3]}</td>
      <td>${r[4]}</td>
      <td>${statusChip(r[5])}</td>
    </tr>`).join('');

  const body = `
    <div style="display:flex;align-items:center;gap:10px;margin-bottom:20px;">
      ${icon(BACK_SVG, 18, NEUTRAL.text1)}
      <span style="font-size:20px;font-weight:600;">Storage Report</span>
    </div>
    <div style="display:flex;align-items:center;gap:16px;flex-wrap:wrap;margin-bottom:24px;">
      <span class="checkbox"><span class="box checked"></span>Include subsites</span>
      <span class="checkbox"><span class="box"></span>Include hidden/system libraries</span>
      <button class="btn primary">Run scan</button>
    </div>
    <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:16px;margin-bottom:24px;">
      ${tile('Total size', '18.6 GB')}
      ${tile('Files scanned', '8,412')}
      ${tile('Stale (2.1 GB)', '412')}
      ${tile('Very stale (4.3 GB)', '96')}
    </div>
    <div style="display:flex;align-items:center;gap:16px;flex-wrap:wrap;margin-bottom:8px;">
      <span class="checkbox"><span class="box checked"></span>Show archival candidates only</span>
      <button class="btn">${icon(DOWNLOAD_SVG, 14, NEUTRAL.text1)} Export Excel</button>
      <button class="btn">${icon(DOWNLOAD_SVG, 14, NEUTRAL.text1)} Export CSV</button>
    </div>
    ${tierLegend(false)}
    <table style="background:#fff;">
      <thead><tr><th>Library</th><th>Name</th><th style="text-align:right;">Size</th><th>Modified</th><th>Author</th><th>Status</th></tr></thead>
      <tbody>${rowsHtml}</tbody>
    </table>`;
  return pageShell(body, { tab: 'report', maxWidth: '1100px' });
}

// ── 6. Storage Report — scan history & compare ───────────────────────────────
function reportHistoryPage() {
  const scans = [
    { date: '7/1/2026, 9:14 AM', size: '18.6 GB', stale: 508, checked: true },
    { date: '6/3/2026, 9:02 AM', size: '17.1 GB', stale: 447, checked: true },
    { date: '5/1/2026, 8:55 AM', size: '15.9 GB', stale: 402, checked: false },
  ];
  const rowsHtml = scans.map((s) => `
    <div style="display:flex;align-items:center;gap:10px;padding:6px 0;">
      <span class="checkbox"><span class="box${s.checked ? ' checked' : ''}"></span></span>
      <span style="min-width:160px;font-size:13px;">${s.date}</span>
      <span style="background:#e8f2fc;color:${BRAND};font-size:11.5px;font-weight:600;padding:2px 8px;border-radius:10px;">${s.size}</span>
      <span style="background:#fdf1e0;color:#8a6100;font-size:11.5px;font-weight:600;padding:2px 8px;border-radius:10px;">${s.stale} stale</span>
      <button class="btn" style="padding:4px 10px;font-size:12px;">View</button>
      <span style="margin-left:auto;">${icon(DELETE_SVG, 16, NEUTRAL.text3)}</span>
    </div>`).join('');

  const body = `
    <div style="display:flex;align-items:center;gap:10px;margin-bottom:24px;">
      ${icon(BACK_SVG, 18, NEUTRAL.text1)}
      <span style="font-size:20px;font-weight:600;">Storage Report</span>
    </div>
    <div style="height:1px;background:${NEUTRAL.border};margin-bottom:24px;"></div>
    <div style="display:flex;align-items:center;gap:6px;margin-bottom:8px;">
      ${icon(HISTORY_SVG, 16, NEUTRAL.text1)}
      <span style="font-weight:600;font-size:14px;">Scan history</span>
    </div>
    <div style="font-size:12px;color:${NEUTRAL.text2};margin-bottom:12px;">Select two scans to compare growth in stale storage over time.</div>
    ${rowsHtml}
    <div style="background:${NEUTRAL.tileBg};border-radius:4px;padding:14px 16px;margin-top:16px;max-width:420px;">
      <div style="font-weight:600;font-size:13px;margin-bottom:8px;">Comparison</div>
      <div style="font-size:13px;line-height:1.8;">
        Size change: <strong style="color:#8a6100;">+1.5 GB</strong><br>
        New archival candidates: <strong>61</strong><br>
        Resolved (no longer stale): <strong>0</strong><br>
        File count change: <strong>+184</strong>
      </div>
    </div>`;
  return pageShell(body, { tab: 'report', maxWidth: '1100px' });
}

// ── 7. Settings ───────────────────────────────────────────────────────────────
function settingsPage() {
  const spin = (value) => `
    <span style="display:inline-flex;align-items:center;border:1px solid #8a8886;border-radius:4px;overflow:hidden;">
      <span style="padding:5px 10px;font-size:13px;min-width:34px;text-align:center;">${value}</span>
      <span style="display:flex;flex-direction:column;border-left:1px solid #8a8886;">
        <span style="padding:1px 6px;font-size:9px;border-bottom:1px solid #8a8886;">▲</span>
        <span style="padding:1px 6px;font-size:9px;">▼</span>
      </span>
    </span>`;
  const section = (title, content) => `
    <div style="margin-bottom:24px;">
      <div style="font-weight:600;font-size:14px;margin-bottom:10px;">${title}</div>
      ${content}
    </div>`;

  const body = `
    <div style="max-width:540px;margin:0 auto;">
      <div style="display:flex;align-items:center;gap:10px;margin-bottom:24px;">
        ${icon(BACK_SVG, 18, NEUTRAL.text1)}
        <span style="font-size:20px;font-weight:600;">Settings</span>
      </div>
      ${section('Scope', `
        <div style="display:flex;align-items:center;gap:6px;margin-bottom:8px;">
          <span class="checkbox"><span class="box"></span>Include subsites in Storage Report scans</span>
          ${icon(INFO_SVG, 14, NEUTRAL.text3)}
        </div>
        <div style="display:flex;align-items:center;gap:6px;">
          <span class="checkbox"><span class="box"></span>Include system and hidden libraries</span>
          ${icon(INFO_SVG, 14, NEUTRAL.text3)}
        </div>`)}
      <div style="height:1px;background:${NEUTRAL.border};margin:20px 0;"></div>
      ${section('Archival thresholds', `
        <div style="font-size:12px;color:${NEUTRAL.text2};margin-bottom:12px;">Based on last modified date only — SharePoint does not expose a reliable last-accessed signal at scale.</div>
        <div style="display:flex;align-items:center;gap:10px;margin-bottom:4px;">
          <span style="font-size:13px;">Stale after (days):</span> ${spin(180)}
        </div>
        <div style="font-size:12px;color:${NEUTRAL.text2};margin-bottom:14px;">Files not modified in this many days are flagged "Stale" — a candidate for review. Default: 180.</div>
        <div style="display:flex;align-items:center;gap:10px;margin-bottom:4px;">
          <span style="font-size:13px;">Very stale after (days):</span> ${spin(365)}
        </div>
        <div style="font-size:12px;color:${NEUTRAL.text2};">Files not modified in this many days are flagged "Very stale" — a strong archival candidate. Default: 365.</div>`)}
      <div style="height:1px;background:${NEUTRAL.border};margin:20px 0;"></div>
      ${section('Performance', `
        <div style="display:flex;align-items:center;gap:10px;margin-bottom:4px;">
          <span style="font-size:13px;">Concurrent API requests:</span> ${spin(4)}
          ${icon(INFO_SVG, 14, NEUTRAL.text3)}
        </div>
        <div style="font-size:12px;color:${NEUTRAL.text2};">Higher values scan faster but may trigger SharePoint throttling. Recommended: 3–5.</div>`)}
      <div style="height:1px;background:${NEUTRAL.border};margin:20px 0;"></div>
      ${section('Default view on load', `
        <div style="font-size:13px;color:${NEUTRAL.text2};margin-bottom:8px;">To change which screen opens when the web part first loads, edit the web part properties:</div>
        <ol style="padding-left:20px;font-size:13px;color:${NEUTRAL.text2};line-height:1.9;">
          <li>Put the SharePoint page into <strong>Edit</strong> mode.</li>
          <li>Click the <strong>pencil (edit)</strong> icon on the Smart Storage Analyzer web part.</li>
          <li>In the property panel, choose a view from the <strong>Default view on open</strong> dropdown.</li>
          <li><strong>Republish</strong> the page to save the change.</li>
        </ol>`)}
    </div>`;
  // The real app hides the brand header entirely on the Settings screen
  // (App.tsx only renders it when view !== 'settings') — just the back
  // button, title, and sections on the plain background.
  return pageShell(body, { noHeader: true, maxWidth: '1200px' });
}

// ── Screenshot runner ─────────────────────────────────────────────────────────
async function main() {
  const browser = await puppeteer.launch({
    executablePath: CHROME,
    headless: 'new',
    args: ['--no-sandbox', '--disable-setuid-sandbox', '--font-render-hinting=none'],
  });

  const shots = [
    ['01_explorer_treemap.png', explorerTreemapPage, { width: 1280, height: 720 }],
    ['02_explorer_list.png', explorerListPage, { width: 1280, height: 560 }],
    ['03_report_config.png', reportConfigPage, { width: 1160, height: 220 }],
    ['04_report_running.png', reportRunningPage, { width: 1160, height: 280 }],
    ['05_report_results.png', reportResultsPage, { width: 1160, height: 640 }],
    ['06_report_history.png', reportHistoryPage, { width: 1160, height: 420 }],
    ['07_settings.png', settingsPage, { width: 1280, height: 760 }],
  ];

  for (const [filename, htmlFn, vp] of shots) {
    const pg = await browser.newPage();
    await pg.setViewport(vp);
    await pg.setContent(htmlFn(), { waitUntil: 'load' });
    // fullPage rather than trusting a guessed viewport height — content
    // length varies per mockup and clipping is worse than a taller image.
    await pg.screenshot({ path: path.join(OUT, filename), fullPage: true });
    await pg.close();
    console.log('✓', filename);
  }

  await browser.close();
  console.log('\nAll screenshots saved to docs/screenshots/');
}

main().catch((e) => { console.error(e); process.exit(1); });
