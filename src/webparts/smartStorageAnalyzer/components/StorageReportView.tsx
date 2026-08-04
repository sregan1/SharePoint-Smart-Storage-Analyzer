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
  Tooltip,
  Spinner,
  Radio,
  RadioGroup,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import {
  ArrowLeft24Regular,
  DocumentArrowDown24Regular,
  History24Regular,
  Delete24Regular,
  Document24Regular,
  Info16Regular,
} from '@fluentui/react-icons';

import { StorageAnalyzerService } from '../services/StorageAnalyzerService';
import { ExcelExportService } from '../services/ExcelExportService';
import { ReportHistoryService } from '../services/ReportHistoryService';
import {
  CandidateTier, FileEntry, ScanProgress, StorageReportSummary, StoredReport, VersionScanMode,
} from '../models/models';
import { QUICK_VERSION_FILE_LIMIT } from '../utils/settingsBounds';
import { StorageTable, StorageTableColumn } from './shared/StorageTable';
import { TierLegend } from './shared/Treemap';
import { formatBytes, formatAge, formatDuration, formatElapsed } from './shared/formatBytes';
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

// Explains the Quick/Full version-history control.
//
// Built as block-level JSX rather than a string: Fluent renders tooltip content
// as HTML, so "\n\n" in a plain string collapses to a single space and the whole
// thing arrives as one dense wall of text. Real elements are the only way to get
// paragraphs, and the term/description pairs below are what make Quick vs Full
// scannable instead of buried mid-sentence.
const VersionScanTooltip: React.FC<{ quickLimit: number }> = ({ quickLimit }) => (
  <div style={{ maxWidth: '320px', display: 'grid', gap: tokens.spacingVerticalS }}>
    <div>
      Most libraries return version-history size in the same bulk read as everything else, at no
      extra scan time. A few lists don&apos;t expose it in bulk at all — those have to be measured
      one file at a time, and this setting only affects those.
    </div>
    <div>
      <strong>Quick</strong> — the largest {quickLimit.toLocaleString()} files per library, where
      nearly all version-history storage is.
    </div>
    <div>
      <strong>Full</strong> — every file. Exact, but can take hours on a large library.
    </div>
    <div>
      A file that couldn&apos;t be measured is always reported as unmeasured, never counted as zero.
    </div>
    <div>
      Version <em>size</em> is exact. Version <em>count</em> is estimated from the file&apos;s current
      version number, so it can run slightly high where a retention limit is configured — never low.
    </div>
  </div>
);

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
  const [includeVersions, setIncludeVersions] = React.useState(false);
  // Only consulted if a library has to escalate all the way to per-file
  // version measurement — see versionSizes.ts. Quick by default because Full
  // on a large library where no bulk mechanism works is genuinely an hours-long
  // operation, and that should be a deliberate choice rather than a surprise.
  const [versionScanMode, setVersionScanMode] = React.useState<VersionScanMode>('quick');
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
  const [showSkippedDetails, setShowSkippedDetails] = React.useState(false);

  // The report currently on screen, when it was loaded from history rather
  // than just produced by a live scan — null means "the results below are
  // from the scan just run with the current Settings thresholds". Needed so
  // the legend/labels reflect the thresholds THAT report was classified
  // with (StoredReport.options), not whatever Settings currently holds.
  const [viewedReport, setViewedReport] = React.useState<StoredReport | null>(null);

  const [history, setHistory] = React.useState<StoredReport[]>([]);
  const [compareIds, setCompareIds] = React.useState<string[]>([]);
  const [showAllSites, setShowAllSites] = React.useState(false);

  // Which history entry's View was just clicked, and which export is running —
  // purely so the button can show a spinner immediately. Both View (setting
  // entries into a many-thousand-row table) and Export (building a workbook or
  // CSV in memory) are synchronous, CPU-bound work on the main thread once
  // started; without a state flip BEFORE that work begins, the click has no
  // visible effect until it's all done, which is what reads as unresponsive on
  // a large report. See handleView/handleExportExcel/handleExportCsv for how
  // the flag is given a chance to paint before the heavy work starts.
  const [pendingViewId, setPendingViewId] = React.useState<string | null>(null);
  const [excelExporting, setExcelExporting] = React.useState(false);
  const [csvExporting, setCsvExporting] = React.useState(false);

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
    setShowSkippedDetails(false);
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
          includeVersionHistory: includeVersions,
          versionScanMode,
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
          options: {
            includeSubsites: subsites, staleDays, veryStaleDays,
            includeVersionHistory: includeVersions, versionScanMode,
          },
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

  // Double rAF: the first fires after the browser has accepted the current
  // frame's updates, the second after it has actually painted them — so the
  // spinner set just before calling this is guaranteed on screen before the
  // synchronous, CPU-bound work in `fn` starts blocking the main thread.
  // (A single rAF, or setTimeout, is not reliably late enough in every
  // browser to guarantee the intervening paint happened first.)
  const afterPaint = (fn: () => void): void => {
    requestAnimationFrame(() => requestAnimationFrame(fn));
  };

  const filteredEntries = React.useMemo(
    () => (staleOnly ? entries.filter((e) => e.tier !== CandidateTier.Active) : entries),
    [entries, staleOnly],
  );

  // Excel export lazy-loads a lib chunk (exceljs) and writes a workbook, so
  // unlike CSV it has real ways to fail (chunk fetch, out-of-memory on a
  // huge report). Left un-caught, a failure here was an invisible unhandled
  // promise rejection — the button just did nothing with no way to tell why.
  const handleExportExcel = (): void => {
    if (!summary || excelExporting) return;
    setError('');
    setExcelExporting(true);
    afterPaint(() => {
      excel.export(filteredEntries, summary, effectiveSiteUrl)
        .catch((err: any) => setError(`Excel export failed: ${err?.message ?? String(err)}`))
        .finally(() => setExcelExporting(false));
    });
  };

  const handleExportCsv = (): void => {
    if (csvExporting) return;
    setError('');
    setCsvExporting(true);
    afterPaint(() => {
      try {
        excel.exportCsv(filteredEntries, !!summary?.versionHistoryIncluded, effectiveSiteUrl);
      } catch (err: any) {
        setError(`CSV export failed: ${err?.message ?? String(err)}`);
      } finally {
        setCsvExporting(false);
      }
    });
  };

  // A report loaded from history shows/hides the version column based on how
  // THAT scan was run, independent of the current (possibly different)
  // checkbox state — mirrors effectiveStaleDays/effectiveVeryStaleDays below.
  const effectiveVersionHistoryIncluded = viewedReport
    ? !!viewedReport.options.includeVersionHistory
    : includeVersions;

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
    ...(effectiveVersionHistoryIncluded ? [{
      key: 'versionSize',
      header: 'Version history',
      align: 'right' as const,
      sortValue: (e: FileEntry) => e.versionSizeBytes ?? -1,
      render: (e: FileEntry) => <span>{e.versionSizeBytes !== undefined ? formatBytes(e.versionSizeBytes) : '—'}</span>,
    }, {
      key: 'versionCount',
      header: 'Version Count (est.)',
      align: 'right' as const,
      sortValue: (e: FileEntry) => e.versionCount ?? -1,
      render: (e: FileEntry) => <span>{e.versionCount !== undefined ? e.versionCount : '—'}</span>,
    }] : []),
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
  // A viewed report can be from a DIFFERENT site (the cross-site history
  // toggle) — exports must reflect that report's own site, not whatever site
  // the page happens to be on right now, same reasoning as the thresholds
  // above. Used for both the export filename and the Summary sheet's own
  // "Site URL" row (see excel.export's siteUrl param).
  const effectiveSiteUrl = viewedReport?.siteUrl ?? siteUrl;

  // ── Phase-scoped progress + ETA ─────────────────────────────────────────
  // A scan has two phases that measure DIFFERENT things: 'items' counts list
  // items read, 'versions' counts files individually measured. They are not
  // interchangeable and must never share a rate.
  //
  // The original ETA projected itemsFetched/elapsed against totalItemsHint
  // unconditionally. That reads correctly during the items phase and is
  // actively misleading during a version phase: itemsFetched is frozen while
  // elapsed keeps climbing, so the estimate decays toward "nearly done"
  // exactly while the slowest part of the scan runs. On a real scan it showed
  // "~0s remaining" with over an hour left.
  //
  // So the phase's own start time is tracked, and each phase is projected
  // using only its own unit over its own elapsed time.
  const phaseStartRef = React.useRef<number>(0);
  const phaseKey = progress?.phase === 'versions'
    // A new library entering the version phase restarts the estimate: its file
    // count (and the tenant's throttle state) bear no relation to the last
    // library's, so carrying the old rate over would be worse than restarting.
    ? `versions:${progress.versionsTotal ?? 0}`
    : 'items';
  React.useEffect(() => {
    phaseStartRef.current = Date.now();
  }, [phaseKey]);

  const versionsPhase = progress?.phase === 'versions' && !!progress.versionsTotal;

  // Determinate fraction for the bar, in whatever unit the current phase
  // measures. Undefined leaves the bar indeterminate, which is the honest
  // rendering when there is no denominator worth trusting.
  const progressValue = React.useMemo(() => {
    if (!progress) return undefined;
    if (versionsPhase) {
      return Math.min(1, (progress.versionsDone ?? 0) / progress.versionsTotal!);
    }
    if (progress.totalItemsHint) {
      return Math.min(1, (progress.itemsFetched ?? 0) / progress.totalItemsHint);
    }
    return progress.libsTotal > 0 ? progress.libsDone / progress.libsTotal : undefined;
  }, [progress, versionsPhase]);

  // Held back until there's enough signal to not be wildly wrong: under a few
  // seconds into a phase, or before anything has completed in it, the rate is
  // either undefined or dominated by fixed per-request overhead rather than
  // real throughput.
  const etaSeconds = React.useMemo(() => {
    if (!progress) return null;
    const phaseElapsed = (Date.now() - phaseStartRef.current) / 1000;
    if (versionsPhase) {
      const done = progress.versionsDone ?? 0;
      if (done <= 0 || phaseElapsed < 3) return null;
      const rate = done / phaseElapsed;
      if (rate <= 0) return null;
      return Math.max(0, progress.versionsTotal! - done) / rate;
    }
    if (!progress.totalItemsHint || elapsed < 3) return null;
    const done = progress.itemsFetched ?? 0;
    if (done <= 0) return null;
    const rate = done / elapsed;
    if (rate <= 0) return null;
    return Math.max(0, progress.totalItemsHint - done) / rate;
    // `elapsed` is a dependency because it ticks every 500ms and is what makes
    // the items-phase estimate refresh; the version phase reads the clock
    // directly off phaseStartRef but still needs that same tick to re-run.
  }, [progress, elapsed, versionsPhase]);

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
    // Version-history gaps, reported as TWO separate conditions because they
    // have two different remedies. Without this, a run that hit either one
    // would silently show 0 B / "—" for the affected files with no indication
    // anything went wrong — indistinguishable from "these files genuinely have
    // no version history".
    const skippedVersions = summary.skippedVersions ?? 0;
    const unmeasuredVersions = summary.unmeasuredVersions ?? 0;
    if (summary.versionHistoryIncluded && skippedVersions > 0) {
      // Attempted and failed — retrying, or lowering concurrency, may fix it.
      partialWarnings.push(
        `version history could not be measured for ${skippedVersions.toLocaleString()} file${skippedVersions === 1 ? '' : 's'} ` +
        '(excluded from the version history totals below, not counted as zero)',
      );
    }
    if (summary.versionHistoryIncluded && unmeasuredVersions > 0) {
      // Never attempted. Retrying changes nothing, so the message names the
      // thing that would actually change the outcome instead.
      partialWarnings.push(
        `${unmeasuredVersions.toLocaleString()} smaller file${unmeasuredVersions === 1 ? '' : 's'} were not measured for version history`
        + (summary.versionScanMode === 'quick'
          ? ` (Quick mode measures the largest ${QUICK_VERSION_FILE_LIMIT.toLocaleString()} per library — re-run with Full to measure all of them)`
          : ' (excluded from the version history totals below, not counted as zero)'),
      );
    }
  }

  const viewReport = (h: StoredReport): void => {
    if (pendingViewId) return;
    setPendingViewId(h.id);
    // setEntries below hands a many-thousand-row array to StorageTable, which
    // renders one real <tr> per row (no virtualization) and sorts the full
    // array up front — on a large report that's real, synchronous work on
    // the main thread. Deferred past a paint so the spinner set above is
    // actually visible first, rather than the click looking dead until the
    // whole re-render finishes.
    afterPaint(() => {
      setViewedReport(h);
      setEntries(h.entries);
      setSummary(h.summary);
      setError('');
      setWarning('');
      setCanceledNotice(false);
      setShowSkippedDetails(false);
      setPendingViewId(null);
    });
  };

  return (
    <div className={styles.root}>
      <div className={styles.header}>
        <Button appearance="subtle" icon={<ArrowLeft24Regular />} onClick={onBack} aria-label="Back to home" />
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
            {!!summary?.skippedFolderDetails?.length && (
              <>
                {' '}
                <Button
                  appearance="transparent"
                  size="small"
                  onClick={() => setShowSkippedDetails((v) => !v)}
                  style={{ padding: 0, minWidth: 'unset', height: 'auto', verticalAlign: 'baseline' }}
                >
                  {showSkippedDetails ? 'Hide details' : 'Show details'}
                </Button>
              </>
            )}
          </MessageBarBody>
        </MessageBar>
      )}

      {showSkippedDetails && !!summary?.skippedFolderDetails?.length && (
        <div className={styles.statTile} style={{ marginBottom: tokens.spacingVerticalM }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: tokens.spacingVerticalXS }}>
            <Text weight="semibold">
              Skipped folders{(summary.skippedFolders ?? 0) > summary.skippedFolderDetails.length
                ? ` (showing first ${summary.skippedFolderDetails.length} of ${summary.skippedFolders})`
                : ''}
            </Text>
            <Button
              size="small"
              appearance="secondary"
              onClick={() => {
                const text = summary.skippedFolderDetails!.map((d) => `${d.url}\t${d.error}`).join('\n');
                navigator.clipboard?.writeText(text).catch(() => { /* clipboard unavailable — ignore */ });
              }}
            >
              Copy to clipboard
            </Button>
          </div>
          <div style={{ maxHeight: '240px', overflowY: 'auto', fontSize: tokens.fontSizeBase200 }}>
            {summary.skippedFolderDetails.map((d, i) => (
              <div key={`${d.url}-${i}`} style={{ padding: `${tokens.spacingVerticalXS} 0`, borderBottom: `1px solid ${tokens.colorNeutralStroke2}` }}>
                <Text style={{ display: 'block', fontFamily: 'monospace' }} title={d.url}>{d.url}</Text>
                <Text style={{ display: 'block', color: tokens.colorNeutralForeground3 }}>{d.error}</Text>
              </div>
            ))}
          </div>
        </div>
      )}

      <div className={styles.row}>
        <Checkbox label="Include subsites" checked={subsites} onChange={(_, d) => setSubsites(!!d.checked)} disabled={scanning} />
        <Checkbox label="Include hidden/system libraries" checked={hidden} onChange={(_, d) => setHidden(!!d.checked)} disabled={scanning} />
        <Checkbox
          label="Include version history size"
          checked={includeVersions}
          onChange={(_, d) => setIncludeVersions(!!d.checked)}
          disabled={scanning}
        />
        <Button appearance="primary" onClick={handleScan} disabled={scanning}>
          {scanning ? 'Scanning…' : 'Run scan'}
        </Button>
        {scanning && (
          <Button appearance="secondary" onClick={handleCancel} disabled={cancelRequested}>
            {cancelRequested ? 'Canceling…' : 'Cancel'}
          </Button>
        )}
      </div>
      {includeVersions && (
        // One line, not four paragraphs. The detail hasn't been deleted — it
        // moved into the same Info16Regular + Tooltip pattern the stat tiles
        // use, so it's available on demand instead of crowding the controls.
        <div
          className={styles.row}
          style={{ marginTop: `-${tokens.spacingVerticalS}`, gap: tokens.spacingHorizontalS }}
        >
          <Text style={{ color: tokens.colorNeutralForeground3, fontSize: tokens.fontSizeBase200 }}>
            If measured per file{' '}
            <Tooltip
              relationship="label"
              content={{
                children: <VersionScanTooltip quickLimit={QUICK_VERSION_FILE_LIMIT} />,
              }}
            >
              <Info16Regular style={{ verticalAlign: 'middle', cursor: 'help' }} />
            </Tooltip>
          </Text>
          <RadioGroup
            layout="horizontal"
            value={versionScanMode}
            onChange={(_, d) => setVersionScanMode(d.value as VersionScanMode)}
            disabled={scanning}
          >
            <Radio value="quick" label={`Quick (largest ${QUICK_VERSION_FILE_LIMIT.toLocaleString()})`} />
            <Radio value="full" label="Full (all files, slow)" />
          </RadioGroup>
        </div>
      )}

      {scanning && progress && (
        <div style={{ marginBottom: tokens.spacingVerticalL }}>
          <ProgressBar value={progressValue} />
          <Text style={{ display: 'block', marginTop: tokens.spacingVerticalXS, color: tokens.colorNeutralForeground3, fontSize: tokens.fontSizeBase200 }}>
            {progress.message}
            {/* The counter shown has to match what the bar is measuring, or
                the two disagree on screen during the version phase. */}
            {versionsPhase
              ? ` (${(progress.versionsDone ?? 0).toLocaleString()} of ${progress.versionsTotal!.toLocaleString()} files measured)`
              : ` — ${progress.scanned} files scanned`}
            {' — '}{formatElapsed(elapsed)} elapsed
            {etaSeconds != null && ` — ~${formatDuration(etaSeconds)} remaining`}
            {/* Why this library is slow, and that stopping is safe — kept to an
                icon so a long scan doesn't grow a paragraph underneath it. */}
            {versionsPhase && (
              <>
                {' '}
                <Tooltip
                  relationship="label"
                  content={
                    'This library doesn\'t expose version-history size in bulk, so each file is '
                    + 'being measured individually. Cancel is safe — everything collected so far is '
                    + 'kept, and unmeasured files are reported as unmeasured rather than zero.'
                  }
                >
                  <Info16Regular style={{ verticalAlign: 'middle', cursor: 'help' }} />
                </Tooltip>
              </>
            )}
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
              <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3 }}>
                Total size{' '}
                <Tooltip
                  content="Sum of every scanned file's current content only — not an estimate, and never includes version history, whether or not that option was enabled for this scan. See Version history size (when shown) for that additional storage."
                  relationship="label"
                >
                  <Info16Regular style={{ verticalAlign: 'middle', cursor: 'help' }} />
                </Tooltip>
              </Text>
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
            {summary.versionHistoryIncluded && (
              <div className={styles.statTile}>
                <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3 }}>
                  Version history size{' '}
                  <Tooltip
                    content="Storage used by older, retained versions of files (SharePoint's version history), on top of the current file content already counted in Total size. This is additional storage consumed in the library."
                    relationship="label"
                  >
                    <Info16Regular style={{ verticalAlign: 'middle', cursor: 'help' }} />
                  </Tooltip>
                </Text>
                <Text weight="semibold" style={{ display: 'block', fontSize: tokens.fontSizeBase500 }}>{formatBytes(summary.totalVersionSizeBytes ?? 0)}</Text>
                {/* How this number was obtained, when it wasn't the free bulk
                    field. Worth surfacing because the difference between the
                    bulk side channel and the per-file pass is the difference
                    between a complete figure and a capped sample — and nothing
                    else on screen would tell the user which one they got.

                    Keyed on unmeasuredVersions, NOT on the scan mode. Quick
                    only samples when there were more candidate files than its
                    budget; when the version labels proved most files had no
                    retained versions, the budget never binds and Quick measured
                    every file that could have contributed. Saying "largest
                    files only" there would understate a complete answer. */}
                {summary.versionSizeStrategy === 'per-file' && (
                  <Text style={{ display: 'block', fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3 }}>
                    {(summary.unmeasuredVersions ?? 0) > 0
                      ? `measured per file — largest ${QUICK_VERSION_FILE_LIMIT.toLocaleString()} only, so this is a floor`
                      : 'measured per file — all files with versions'}
                  </Text>
                )}
                {summary.versionSizeStrategy === 'none' && (
                  <Text style={{ display: 'block', fontSize: tokens.fontSizeBase200, color: tokens.colorPaletteMarigoldForeground1 }}>
                    incomplete — no version-history source available for at least one library
                  </Text>
                )}
              </div>
            )}
          </div>

          <div className={styles.row}>
            <Checkbox label="Show archival candidates only" checked={staleOnly} onChange={(_, d) => setStaleOnly(!!d.checked)} />
            <Button
              icon={excelExporting ? <Spinner size="tiny" /> : <DocumentArrowDown24Regular />}
              onClick={handleExportExcel}
              disabled={excelExporting}
            >
              {excelExporting ? 'Exporting…' : 'Export Excel'}
            </Button>
            <Button
              icon={csvExporting ? <Spinner size="tiny" /> : <DocumentArrowDown24Regular />}
              appearance="secondary"
              onClick={handleExportCsv}
              disabled={csvExporting}
            >
              {csvExporting ? 'Exporting…' : 'Export CSV'}
            </Button>
          </div>

          <TierLegend staleDays={effectiveStaleDays} veryStaleDays={effectiveVeryStaleDays} showFolder={false} />

          <StorageTable
            rows={filteredEntries}
            columns={columns}
            getRowKey={(e) => e.serverRelativeUrl}
            defaultSortKey="size"
            // A full-site scan can return hundreds of thousands of rows —
            // rendering all of them at once is what makes the results table
            // slow to paint and heavy to scroll. Export (Excel/CSV) is
            // unaffected — both already export the full filteredEntries
            // array regardless of what's paged into view here.
            pageSize={200}
          />
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
              {h.summary.versionHistoryIncluded && (
                <Badge appearance="tint" title="Version history size">{formatBytes(h.summary.totalVersionSizeBytes ?? 0)} versions</Badge>
              )}
              {h.entriesTruncated && (
                <Badge appearance="tint" color="informative" title="Report exceeded the row limit — only stale/very-stale files were saved">
                  Partial
                </Badge>
              )}
              <Button
                appearance="subtle"
                size="small"
                icon={pendingViewId === h.id ? <Spinner size="tiny" /> : undefined}
                onClick={() => viewReport(h)}
                disabled={!!pendingViewId}
              >
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
