import * as React from 'react';
import {
  Button,
  Checkbox,
  Title3,
  Body1,
  Text,
  Badge,
  ProgressBar,
  MessageBar,
  MessageBarBody,
  Divider,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import {
  ArrowLeft24Regular,
  DocumentArrowDown24Regular,
  History24Regular,
  Delete24Regular,
  Document24Regular,
} from '@fluentui/react-icons';

import { StorageAnalyzerService } from '../services/StorageAnalyzerService';
import { ExcelExportService } from '../services/ExcelExportService';
import { ReportHistoryService } from '../services/ReportHistoryService';
import { CandidateTier, FileEntry, ScanProgress, StorageReportSummary, StoredReport } from '../models/models';
import { StorageTable, StorageTableColumn } from './shared/StorageTable';
import { TierLegend } from './shared/Treemap';
import { formatBytes, formatAge } from './shared/formatBytes';
import { tierColor, tierLabel } from './shared/tierBadge';
import { diffReports } from '../utils/reportDiff';

const useStyles = makeStyles({
  root: {
    padding: tokens.spacingVerticalL,
    maxWidth: '1100px',
    margin: '0 auto',
    minHeight: '500px',
  },
  header: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    marginBottom: tokens.spacingVerticalL,
  },
  row: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    flexWrap: 'wrap',
    marginBottom: tokens.spacingVerticalM,
  },
  summaryGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fit, minmax(140px, 1fr))',
    gap: tokens.spacingHorizontalM,
    marginBottom: tokens.spacingVerticalL,
  },
  statTile: {
    padding: tokens.spacingVerticalM,
    background: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
  },
  historyRow: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    padding: `${tokens.spacingVerticalXS} 0`,
  },
});

const reportHistory = new ReportHistoryService();

// A report whose entries would otherwise strain IndexedDB quota (10 large
// scans at, say, 200k rows each) is stored with only its stale/very-stale
// rows — diffReports only ever reads stale-tier paths plus summary fields,
// so comparisons stay fully accurate; only the "View" file listing for that
// one report is incomplete (flagged via entriesTruncated).
const MAX_STORED_ENTRIES = 50000;

// Site-relative path, used to tell scans from different (sub)sites apart in
// the shared IndexedDB history — falls back to the raw string if it isn't a
// parseable absolute URL.
function siteLabel(url: string): string {
  try {
    const path = new URL(url).pathname.replace(/\/+$/, '');
    return path || '/';
  } catch {
    return url;
  }
}

export interface StorageReportViewProps {
  sp: StorageAnalyzerService;
  excel: ExcelExportService;
  siteUrl: string;
  includeSubsites: boolean;
  includeHidden: boolean;
  staleDays: number;
  veryStaleDays: number;
  onBack: () => void;
}

export const StorageReportView: React.FC<StorageReportViewProps> = ({
  sp, excel, siteUrl, includeSubsites, includeHidden, staleDays, veryStaleDays, onBack,
}) => {
  const styles = useStyles();

  const [subsites, setSubsites] = React.useState(includeSubsites);
  const [hidden, setHidden] = React.useState(includeHidden);
  const [scanning, setScanning] = React.useState(false);
  const [cancelRequested, setCancelRequested] = React.useState(false);
  const [canceledNotice, setCanceledNotice] = React.useState(false);
  const [progress, setProgress] = React.useState<ScanProgress | null>(null);
  const [elapsed, setElapsed] = React.useState(0);
  const [entries, setEntries] = React.useState<FileEntry[]>([]);
  const [summary, setSummary] = React.useState<StorageReportSummary | null>(null);
  const [error, setError] = React.useState('');
  const [warning, setWarning] = React.useState('');
  const [staleOnly, setStaleOnly] = React.useState(false);

  // The report currently on screen, when it was loaded from history rather
  // than just produced by a live scan — null means "the results below are
  // from the scan just run with the current Settings thresholds". Needed so
  // the legend/labels reflect the thresholds THAT report was classified
  // with (StoredReport.options), not whatever Settings currently holds.
  const [viewedReport, setViewedReport] = React.useState<StoredReport | null>(null);

  const [history, setHistory] = React.useState<StoredReport[]>([]);
  const [compareIds, setCompareIds] = React.useState<string[]>([]);
  const [showAllSites, setShowAllSites] = React.useState(false);

  const abortControllerRef = React.useRef<AbortController | null>(null);
  // Per-file counter fed by onEntry (fires potentially thousands of times
  // per scan) — flushed into `progress` state on the existing elapsed-timer
  // tick rather than on every file, so a single-library scan shows live
  // movement without re-rendering per file.
  const scannedRef = React.useRef(0);

  const loadHistory = React.useCallback((): void => {
    reportHistory.getAll().then(setHistory).catch(() => setHistory([]));
  }, []);

  React.useEffect(() => { loadHistory(); }, [loadHistory]);

  const handleScan = async (): Promise<void> => {
    setScanning(true);
    setCancelRequested(false);
    setCanceledNotice(false);
    setError('');
    setWarning('');
    setViewedReport(null);
    setEntries([]);
    setSummary(null);
    setElapsed(0);
    scannedRef.current = 0;
    setProgress({ message: 'Starting scan…', scanned: 0, libsDone: 0, libsTotal: 0 });
    const start = Date.now();
    const timer = setInterval(() => {
      setElapsed((Date.now() - start) / 1000);
      setProgress((p) => (p ? { ...p, scanned: scannedRef.current } : p));
    }, 500);

    const abortController = new AbortController();
    abortControllerRef.current = abortController;

    try {
      const result = await sp.scanSite(
        {
          siteUrl, includeSubsites: subsites, includeHidden: hidden, staleDays, veryStaleDays,
          scanConcurrency: sp.scanConcurrency, signal: abortController.signal,
        },
        (p) => { scannedRef.current = p.scanned; setProgress(p); },
        () => { scannedRef.current++; },
      );
      setEntries(result.entries);
      setSummary(result.summary);

      if (result.canceled) {
        // Cooperative cancellation only stops the walk between queued
        // folders/libraries/sites, so entries collected up to that point are
        // real (if partial) — worth keeping on screen, just not worth
        // saving to history as if it were a complete scan.
        setCanceledNotice(true);
      } else {
        const truncate = result.entries.length > MAX_STORED_ENTRIES;
        const stored: StoredReport = {
          id: `${Date.now()}`,
          timestamp: Date.now(),
          siteUrl,
          options: { includeSubsites: subsites, staleDays, veryStaleDays },
          summary: result.summary,
          entries: truncate ? result.entries.filter((e) => e.tier !== CandidateTier.Active) : result.entries,
          entriesTruncated: truncate,
        };
        try {
          await reportHistory.add(stored);
          loadHistory();
        } catch (err: any) {
          // The scan itself succeeded and its results are already on
          // screen — only the save to history failed (e.g. IndexedDB quota
          // on a very large report). Reporting this as "Scan failed" would
          // be actively misleading.
          setWarning(`Scan completed, but the report could not be saved to history: ${err?.message ?? String(err)}`);
        }
      }
    } catch (err: any) {
      setError(`Scan failed: ${err?.message ?? String(err)}`);
    } finally {
      clearInterval(timer);
      setScanning(false);
      setCancelRequested(false);
      setProgress(null);
      abortControllerRef.current = null;
    }
  };

  const handleCancel = (): void => {
    setCancelRequested(true);
    abortControllerRef.current?.abort();
  };

  const filteredEntries = React.useMemo(
    () => (staleOnly ? entries.filter((e) => e.tier !== CandidateTier.Active) : entries),
    [entries, staleOnly],
  );

  // Excel export lazy-loads a lib chunk (exceljs) and writes a workbook, so
  // unlike CSV it has real ways to fail (chunk fetch, out-of-memory on a
  // huge report). Left un-caught, a failure here was an invisible unhandled
  // promise rejection — the button just did nothing with no way to tell why.
  const handleExportExcel = async (): Promise<void> => {
    if (!summary) return;
    try {
      setError('');
      await excel.export(filteredEntries, summary, siteUrl);
    } catch (err: any) {
      setError(`Excel export failed: ${err?.message ?? String(err)}`);
    }
  };

  const columns: StorageTableColumn<FileEntry>[] = [
    { key: 'library', header: 'Library', sortValue: (e) => e.libraryTitle, render: (e) => <span>{e.libraryTitle}</span> },
    {
      key: 'name',
      header: 'Name',
      sortValue: (e) => e.name,
      render: (e) => (
        <span style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
          <Document24Regular style={{ fontSize: '14px', flexShrink: 0 }} />
          <span title={e.serverRelativeUrl}>{e.name}</span>
        </span>
      ),
    },
    { key: 'size', header: 'Size', align: 'right', sortValue: (e) => e.sizeBytes, render: (e) => <span>{formatBytes(e.sizeBytes)}</span> },
    {
      key: 'modified',
      header: 'Modified',
      sortValue: (e) => e.timeLastModified,
      render: (e) => <span>{new Date(e.timeLastModified).toLocaleDateString()} ({formatAge(e.ageDays)} ago)</span>,
    },
    { key: 'author', header: 'Author', sortValue: (e) => e.authorDisplayName ?? '', render: (e) => <span>{e.authorDisplayName ?? '—'}</span> },
    {
      key: 'tier',
      header: 'Status',
      sortValue: (e) => e.ageDays,
      render: (e) => (
        <span style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
          <span style={{ width: '8px', height: '8px', borderRadius: '50%', background: tierColor(e.tier), flexShrink: 0 }} />
          <span>{tierLabel(e.tier)}</span>
        </span>
      ),
    },
  ];

  const toggleCompare = (id: string): void => {
    setCompareIds((prev) => {
      if (prev.indexOf(id) !== -1) return prev.filter((x) => x !== id);
      if (prev.length >= 2) return [prev[1], id];
      return [...prev, id];
    });
  };

  const comparison = React.useMemo(() => {
    if (compareIds.length !== 2) return null;
    const [a, b] = compareIds.map((id) => history.find((h) => h.id === id)).filter((x): x is StoredReport => !!x);
    if (!a || !b) return null;
    const [older, newer] = a.timestamp <= b.timestamp ? [a, b] : [b, a];
    return { diff: diffReports(older, newer), crossSite: a.siteUrl !== b.siteUrl };
  }, [compareIds, history]);

  // IndexedDB is per-origin, so every site's/subsite's reports land in the
  // same store — filtered to the current site by default so a diff or the
  // history list doesn't casually mix unrelated sites, with an explicit
  // opt-in to see the rest.
  const currentSiteHistory = React.useMemo(
    () => history.filter((h) => h.siteUrl === siteUrl),
    [history, siteUrl],
  );
  const visibleHistory = showAllSites ? history : currentSiteHistory;
  const otherSiteCount = history.length - currentSiteHistory.length;

  const effectiveStaleDays = viewedReport?.options.staleDays ?? staleDays;
  const effectiveVeryStaleDays = viewedReport?.options.veryStaleDays ?? veryStaleDays;

  const partialWarnings: string[] = [];
  if (summary) {
    const skippedFolders = summary.skippedFolders ?? 0;
    const skippedSites = summary.skippedSites ?? 0;
    if (skippedFolders > 0) {
      partialWarnings.push(`${skippedFolders} folder${skippedFolders === 1 ? '' : 's'} could not be read and ${skippedFolders === 1 ? 'was' : 'were'} skipped`);
    }
    if (skippedSites > 0) {
      partialWarnings.push(`${skippedSites} subsite${skippedSites === 1 ? '' : 's'} could not be accessed and ${skippedSites === 1 ? 'was' : 'were'} skipped`);
    }
  }

  const viewReport = (h: StoredReport): void => {
    setViewedReport(h);
    setEntries(h.entries);
    setSummary(h.summary);
    setError('');
    setWarning('');
    setCanceledNotice(false);
  };

  return (
    <div className={styles.root}>
      <div className={styles.header}>
        <Button appearance="subtle" icon={<ArrowLeft24Regular />} onClick={onBack} aria-label="Back to explorer" />
        <Title3>Storage Report</Title3>
      </div>

      {error && (
        <MessageBar intent="error" style={{ marginBottom: tokens.spacingVerticalM }}>
          <MessageBarBody>{error}</MessageBarBody>
        </MessageBar>
      )}
      {warning && (
        <MessageBar intent="warning" style={{ marginBottom: tokens.spacingVerticalM }}>
          <MessageBarBody>{warning}</MessageBarBody>
        </MessageBar>
      )}
      {(canceledNotice || partialWarnings.length > 0) && (
        <MessageBar intent="warning" style={{ marginBottom: tokens.spacingVerticalM }}>
          <MessageBarBody>
            {canceledNotice && 'Scan canceled — showing partial results collected before cancellation. '}
            {partialWarnings.length > 0 && `Results are partial: ${partialWarnings.join('; ')}.`}
          </MessageBarBody>
        </MessageBar>
      )}

      <div className={styles.row}>
        <Checkbox label="Include subsites" checked={subsites} onChange={(_, d) => setSubsites(!!d.checked)} disabled={scanning} />
        <Checkbox label="Include hidden/system libraries" checked={hidden} onChange={(_, d) => setHidden(!!d.checked)} disabled={scanning} />
        <Button appearance="primary" onClick={handleScan} disabled={scanning}>
          {scanning ? 'Scanning…' : 'Run scan'}
        </Button>
        {scanning && (
          <Button appearance="secondary" onClick={handleCancel} disabled={cancelRequested}>
            {cancelRequested ? 'Canceling…' : 'Cancel'}
          </Button>
        )}
      </div>

      {scanning && progress && (
        <div style={{ marginBottom: tokens.spacingVerticalL }}>
          <ProgressBar value={progress.libsTotal > 0 ? progress.libsDone / progress.libsTotal : undefined} />
          <Text style={{ display: 'block', marginTop: tokens.spacingVerticalXS, color: tokens.colorNeutralForeground3, fontSize: tokens.fontSizeBase200 }}>
            {progress.message} — {progress.scanned} files scanned — {elapsed.toFixed(0)}s elapsed
          </Text>
        </div>
      )}

      {summary && (
        <>
          {viewedReport && (
            <Text style={{ display: 'block', color: tokens.colorNeutralForeground3, marginBottom: tokens.spacingVerticalS }}>
              Viewing saved scan from {new Date(viewedReport.timestamp).toLocaleString()} — {siteLabel(viewedReport.siteUrl)}
              {viewedReport.entriesTruncated ? ' (only stale/very-stale files were saved — this report exceeded the row limit)' : ''}
            </Text>
          )}

          <div className={styles.summaryGrid}>
            <div className={styles.statTile}>
              <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3 }}>Total size</Text>
              <Text weight="semibold" style={{ display: 'block', fontSize: tokens.fontSizeBase500 }}>{formatBytes(summary.totalSizeBytes)}</Text>
            </div>
            <div className={styles.statTile}>
              <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3 }}>Files scanned</Text>
              <Text weight="semibold" style={{ display: 'block', fontSize: tokens.fontSizeBase500 }}>{summary.totalFiles}</Text>
            </div>
            <div className={styles.statTile}>
              <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3 }}>Stale ({formatBytes(summary.staleSizeBytes)})</Text>
              <Text weight="semibold" style={{ display: 'block', fontSize: tokens.fontSizeBase500 }}>{summary.staleCount}</Text>
            </div>
            <div className={styles.statTile}>
              <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3 }}>Very stale ({formatBytes(summary.veryStaleSizeBytes)})</Text>
              <Text weight="semibold" style={{ display: 'block', fontSize: tokens.fontSizeBase500 }}>{summary.veryStaleCount}</Text>
            </div>
          </div>

          <div className={styles.row}>
            <Checkbox label="Show archival candidates only" checked={staleOnly} onChange={(_, d) => setStaleOnly(!!d.checked)} />
            <Button icon={<DocumentArrowDown24Regular />} onClick={handleExportExcel}>
              Export Excel
            </Button>
            <Button icon={<DocumentArrowDown24Regular />} appearance="secondary" onClick={() => excel.exportCsv(filteredEntries)}>
              Export CSV
            </Button>
          </div>

          <TierLegend staleDays={effectiveStaleDays} veryStaleDays={effectiveVeryStaleDays} showFolder={false} />

          <StorageTable rows={filteredEntries} columns={columns} getRowKey={(e) => e.serverRelativeUrl} defaultSortKey="size" />
        </>
      )}

      <Divider style={{ margin: `${tokens.spacingVerticalXL} 0` }} />

      <Text weight="semibold" style={{ display: 'flex', alignItems: 'center', gap: '6px', marginBottom: tokens.spacingVerticalS }}>
        <History24Regular style={{ fontSize: '16px' }} /> Scan history
      </Text>
      {history.length === 0 ? (
        <Body1 style={{ color: tokens.colorNeutralForeground3 }}>No saved scans yet — run a scan to save one.</Body1>
      ) : (
        <>
          <Body1 style={{ color: tokens.colorNeutralForeground3, fontSize: tokens.fontSizeBase200 }}>
            Select two scans to compare growth in stale storage over time.
          </Body1>
          {otherSiteCount > 0 && (
            <Button
              appearance="transparent"
              size="small"
              onClick={() => setShowAllSites((v) => !v)}
              style={{ padding: 0, minWidth: 'unset', height: 'auto', marginBottom: tokens.spacingVerticalXS }}
            >
              {showAllSites ? 'Show only this site' : `Show reports from other sites (${otherSiteCount} hidden)`}
            </Button>
          )}
          {visibleHistory.length === 0 ? (
            <Body1 style={{ color: tokens.colorNeutralForeground3 }}>No saved scans for this site yet.</Body1>
          ) : visibleHistory.map((h) => (
            <div key={h.id} className={styles.historyRow}>
              <Checkbox checked={compareIds.indexOf(h.id) !== -1} onChange={() => toggleCompare(h.id)} />
              <Text style={{ minWidth: '160px' }}>{new Date(h.timestamp).toLocaleString()}</Text>
              <Badge appearance="tint" title={h.siteUrl}>{siteLabel(h.siteUrl)}</Badge>
              <Badge appearance="tint">{formatBytes(h.summary.totalSizeBytes)}</Badge>
              <Badge appearance="tint" color="warning">{h.summary.staleCount + h.summary.veryStaleCount} stale</Badge>
              {h.entriesTruncated && (
                <Badge appearance="tint" color="informative" title="Report exceeded the row limit — only stale/very-stale files were saved">
                  Partial
                </Badge>
              )}
              <Button appearance="subtle" size="small" onClick={() => viewReport(h)}>
                View
              </Button>
              <Button
                appearance="subtle"
                size="small"
                icon={<Delete24Regular />}
                aria-label="Delete scan"
                onClick={() => reportHistory.delete(h.id).then(loadHistory)}
              />
            </div>
          ))}
        </>
      )}

      {comparison && (
        <div className={styles.statTile} style={{ marginTop: tokens.spacingVerticalM }}>
          <Text weight="semibold" style={{ display: 'block', marginBottom: tokens.spacingVerticalXS }}>Comparison</Text>
          {comparison.crossSite && (
            <Text style={{ display: 'block', color: tokens.colorPaletteYellowForeground1 }}>
              ⚠ These two reports are from different sites — the comparison may not be meaningful.
            </Text>
          )}
          <Text style={{ display: 'block' }}>
            Size change: {comparison.diff.sizeDeltaBytes >= 0 ? '+' : ''}{formatBytes(Math.abs(comparison.diff.sizeDeltaBytes))}
          </Text>
          <Text style={{ display: 'block' }}>New archival candidates: {comparison.diff.newStaleCount}</Text>
          <Text style={{ display: 'block' }}>Resolved (no longer stale): {comparison.diff.resolvedStaleCount}</Text>
          <Text style={{ display: 'block' }}>File count change: {comparison.diff.totalFilesDelta >= 0 ? '+' : ''}{comparison.diff.totalFilesDelta}</Text>
        </div>
      )}
    </div>
  );
};
