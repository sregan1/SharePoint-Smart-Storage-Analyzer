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
  CandidateTier, FileEntry, ReportDiff, ScanProgress, StorageReportSummary, StoredReport,
  StoredReportMeta,
} from '../models/models';
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

// What a stage is actually doing, for the status block's info hover. Detail that
// is diagnostic rather than at-a-glance belongs here, not on screen — this
// display has twice grown into a wall of text by explaining itself inline.
//
// Block-level JSX rather than a string, because Fluent renders tooltip content as
// HTML: "\n\n" collapses to a single space and the whole thing arrives as one
// dense paragraph.
const STAGE_NOTES: Record<ScanProgress['stage'], string> = {
  'discovering': 'Reading the site structure and the list of libraries to scan.',
  'items': 'One flat read per library — about one request per 5,000 items, regardless of how the '
    + 'folders are arranged.',
  'version-probe': 'Working out how this list will report version-history size, before spending '
    + 'anything on measuring it.',
  'version-bulk': 'Reading version-history size in bulk. Each page covers 5,000 items.',
  'version-validate': 'Checking the bulk figures against each file\'s version number, so an '
    + 'unpopulated field can\'t be reported as a confident zero.',
  'version-per-file': 'This list doesn\'t report version-history size in bulk, so each file is '
    + 'measured individually. Cancel is safe — everything collected so far is kept, and anything '
    + 'unmeasured is reported as unmeasured rather than zero.',
};

const ScanStatusTooltip: React.FC<{
  progress: ScanProgress;
  requestCount: number;
}> = ({ progress, requestCount }) => (
  <div style={{ maxWidth: '340px', display: 'grid', gap: tokens.spacingVerticalS }}>
    <div>{STAGE_NOTES[progress.stage]}</div>
    <div>
      {/* The liveness number. It is the only counter that moves during a stage
          with no denominator, so it is what proves the scan isn't hung. */}
      {requestCount.toLocaleString()} SharePoint requests so far
      {progress.itemsFetched != null && (
        <>
          <br />
          {progress.itemsFetched.toLocaleString()} items read
          {progress.totalItemsHint ? ` of roughly ${progress.totalItemsHint.toLocaleString()} expected` : ''}
        </>
      )}
      <br />
      {progress.scanned.toLocaleString()} of {progress.filesSeen.toLocaleString()} files added to the report
    </div>
    {(progress.skippedVersions > 0 || progress.unmeasuredVersions > 0) && (
      <div>
        {progress.skippedVersions > 0 && (
          <>{progress.skippedVersions.toLocaleString()} version lookups failed<br /></>
        )}
        {progress.unmeasuredVersions > 0 && (
          <>{progress.unmeasuredVersions.toLocaleString()} files left unmeasured</>
        )}
      </div>
    )}
  </div>
);

// There was a MAX_STORED_ENTRIES row cap here, which filtered Active-tier rows
// out of any report over 50,000 files. It was solving the wrong problem: the
// pressure came from ReportHistoryService.getAll() loading every saved report's
// entire listing just to render the history list, not from disk. Now that
// listings are stored separately and loaded on demand, saved reports are
// complete — and the cap being a tier FILTER rather than a slice meant it
// dropped 12% of a real 184,860-row report while saving almost nothing.

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
  // Pulled off the client on the scan ticker rather than pushed through
  // ScanProgress: both are properties of the API client, not of the scan, and
  // ExplorerView already surfaces throttling this way.
  const [throttleWait, setThrottleWait] = React.useState(0);
  const [requestCount, setRequestCount] = React.useState(0);
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
  const [viewedReport, setViewedReport] = React.useState<StoredReportMeta | null>(null);

  // Summaries only. The file listings live in a separate store and are fetched
  // on demand — loading all of them here is what previously forced saved reports
  // to be trimmed.
  const [history, setHistory] = React.useState<StoredReportMeta[]>([]);
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
    setProgress({
      stage: 'discovering',
      stageLabel: 'Starting scan',
      stageKey: 'start',
      libsDone: 0,
      libsTotal: 0,
      filesSeen: 0,
      scanned: 0,
      skippedVersions: 0,
      unmeasuredVersions: 0,
      skippedLibraries: 0,
    });
    const start = Date.now();
    // Also what keeps the display alive during a tenant-imposed pause: nothing
    // completes then, so no data-layer callback can fire, and it is this tick
    // plus sp.throttleRemainingSeconds that turns "frozen" into "waiting 44s".
    const timer = setInterval(() => {
      setElapsed((Date.now() - start) / 1000);
      setProgress((p) => (p ? { ...p, scanned: scannedRef.current } : p));
      setThrottleWait(sp.isThrottled ? sp.throttleRemainingSeconds : 0);
      setRequestCount(sp.requestCount);
    }, 500);

    const abortController = new AbortController();
    abortControllerRef.current = abortController;

    try {
      const result = await sp.scanSite(
        {
          siteUrl, includeSubsites: subsites, includeHidden: hidden, staleDays, veryStaleDays,
          scanConcurrency: sp.scanConcurrency, signal: abortController.signal,
          includeVersionHistory: includeVersions,
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
        const stored: StoredReport = {
          id: `${Date.now()}`,
          timestamp: Date.now(),
          siteUrl,
          options: {
            includeSubsites: subsites, staleDays, veryStaleDays,
            includeVersionHistory: includeVersions,
          },
          summary: result.summary,
          // Every row. No filtering — see the note where MAX_STORED_ENTRIES used
          // to be.
          entries: result.entries,
          entryCount: result.entries.length,
        };
        try {
          const { listingSaved } = await reportHistory.add(stored);
          if (!listingSaved) {
            // The numbers were saved; only the per-file listing didn't fit. Say
            // exactly that, because "could not be saved" would imply the whole
            // report was lost when the summary and every comparison still work.
            setWarning(
              'Scan completed and saved, but browser storage was full so this report\'s file '
              + 'listing was not kept — its totals and comparisons are unaffected. Delete an older '
              + 'report to free space.',
            );
          }
          loadHistory();
        } catch (err: any) {
          // The scan itself succeeded and its results are already on
          // screen — only the save to history failed. Reporting this as
          // "Scan failed" would be actively misleading.
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

  // Was a useMemo, but the stale-path half of a diff needs both reports' file
  // listings, which are now fetched. An effect rather than a memo, keyed on the
  // two ids only — the previous memo depended on `history`, whose identity
  // changes on every reload, so it recomputed far more often than the selection
  // actually changed.
  const [comparison, setComparison] = React.useState<
    { diff: ReportDiff; crossSite: boolean; listingMissing: boolean } | null
  >(null);
  const [comparing, setComparing] = React.useState(false);
  const compareKey = compareIds.join('|');

  React.useEffect(() => {
    if (compareIds.length !== 2) { setComparison(null); return; }
    const a = history.find((h) => h.id === compareIds[0]);
    const b = history.find((h) => h.id === compareIds[1]);
    if (!a || !b) { setComparison(null); return; }
    const [older, newer] = a.timestamp <= b.timestamp ? [a, b] : [b, a];

    // Guards against a slow load landing after the selection has moved on.
    let canceled = false;
    setComparing(true);
    Promise.all([reportHistory.getEntries(older.id), reportHistory.getEntries(newer.id)])
      .then(([olderEntries, newerEntries]) => {
        if (canceled) return;
        setComparison({
          diff: diffReports(older, newer, olderEntries, newerEntries),
          crossSite: a.siteUrl !== b.siteUrl,
          listingMissing: !olderEntries || !newerEntries,
        });
      })
      .catch(() => {
        if (!canceled) setComparison(null);
      })
      .finally(() => {
        if (!canceled) setComparing(false);
      });
    return () => { canceled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [compareKey, history]);

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

  // ── Stage-scoped progress + ETA ─────────────────────────────────────────
  // Every stage measures a DIFFERENT unit (items, pages, files) and must never
  // share a rate window with another. The rule for both the bar and the ETA is
  // the same: use a denominator the producer still stands behind, or show
  // nothing at all.
  //
  // The start time is compared DURING RENDER, not in an effect. An effect
  // commits after the render that already read the ref, so the first frame of a
  // new stage projected the new stage's counter against the old stage's start —
  // and with the previous initial value of 0 that meant dividing by ~1.75e9
  // seconds, which could render an ETA in five figures of hours. Writing a ref
  // during render is safe here: it isn't state, nothing else reads it this pass,
  // and re-running the comparison for the same key is idempotent.
  const stageStartRef = React.useRef(Date.now());
  const stageKeyRef = React.useRef('');
  if (progress && progress.stageKey !== stageKeyRef.current) {
    stageKeyRef.current = progress.stageKey;
    stageStartRef.current = Date.now();
  }

  const throttled = throttleWait > 0;

  // Determinate fraction for the bar, in whatever unit the current stage
  // measures. Undefined leaves it indeterminate, which is the honest rendering
  // when there is no denominator worth trusting.
  //
  // No Math.min clamp, and no libsDone/libsTotal fallback: the clamp is what
  // pegged the bar at 100% while work continued, and the fallback jumped
  // BACKWARDS whenever a library moved from its items sweep into version work
  // (libsDone hasn't moved, but the new stage restarts from zero).
  const progressValue = React.useMemo(() => {
    if (!progress) return undefined;
    const { stageDone, stageTotal } = progress;
    if (stageTotal == null || stageTotal <= 0) return undefined;
    if (stageDone == null || stageDone > stageTotal) return undefined;
    return stageDone / stageTotal;
  }, [progress]);

  const etaSeconds = React.useMemo(() => {
    if (!progress) return null;
    const { stageDone, stageTotal } = progress;
    if (stageTotal == null || stageDone == null) return null;
    // While the tenant holds the gate, the rate measures a PAUSE, not
    // throughput. The status line already says "paused"; a decaying estimate
    // beside it would be two lies for the price of one.
    if (throttled) return null;
    const remaining = stageTotal - stageDone;
    // NOT Math.max(0, …). That clamp is exactly why the line read
    // "~0s remaining" for eleven minutes: a stale estimate gets reached and then
    // exceeded, the subtraction goes negative, the clamp turns it into a
    // confident zero, and the render guard happily prints it. A spent
    // denominator means NO estimate.
    if (remaining <= 0) return null;
    const stageElapsed = (Date.now() - stageStartRef.current) / 1000;
    // 5s rather than 3: on a throttled tenant 3 seconds is often less than a
    // single round trip, so the first sample is pure fixed overhead.
    if (stageElapsed < 5 || stageDone <= 0) return null;
    return remaining / (stageDone / stageElapsed);
    // `elapsed` ticks every 500ms and is what makes this re-run.
  }, [progress, elapsed, throttled]);

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
      // Never attempted. Retrying changes nothing, so the message must not
      // suggest it — except for a report saved by a build that HAD a Quick mode,
      // where re-running genuinely is the fix. That is the one thing the legacy
      // versionScanMode field is still read for.
      partialWarnings.push(
        `${unmeasuredVersions.toLocaleString()} file${unmeasuredVersions === 1 ? '' : 's'} were not measured for version history`
        + (summary.versionScanMode === 'quick'
          ? ' — this saved report ran in the old Quick mode, which measured only the largest files per library; re-run to measure them all'
          : ' (reported as unmeasured, never counted as zero)'),
      );
    }
  }

  const viewReport = async (h: StoredReportMeta): Promise<void> => {
    if (pendingViewId) return;
    setPendingViewId(h.id);
    setError('');
    setWarning('');
    setCanceledNotice(false);
    setShowSkippedDetails(false);

    // The listing is no longer carried with the metadata, so it has to be read
    // now. The spinner is already showing because pendingViewId was set above.
    let loaded: FileEntry[] | null = null;
    try {
      loaded = await reportHistory.getEntries(h.id);
    } catch (err: any) {
      setError(`Could not load this report's file listing: ${err?.message ?? String(err)}`);
      setPendingViewId(null);
      return;
    }

    // setEntries below hands a many-thousand-row array to StorageTable, which
    // renders one real <tr> per row (no virtualization) and sorts the full
    // array up front — on a large report that's real, synchronous work on
    // the main thread. Deferred past a paint so the spinner is actually visible
    // first, rather than the click looking dead until the re-render finishes.
    afterPaint(() => {
      setViewedReport(h);
      setEntries(loaded ?? []);
      setSummary(h.summary);
      if (!loaded) {
        // Summary and comparisons are intact; only the per-file rows are gone.
        // Said plainly, because an empty table with no explanation reads as a bug.
        setWarning(
          'This report\'s file listing is no longer stored — it was removed to make room for newer '
          + 'reports. Its totals below are complete and unaffected; re-run the scan to get the file '
          + 'list back.',
        );
      }
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
      {scanning && progress && (
        // Three fixed lines, deliberately. Line 1 is what/where, line 2 is the
        // stage's own progress, line 3 is scan-wide and always ticking. Anything
        // diagnostic lives behind the single info hover — this display has grown
        // into a wall of text twice by explaining itself inline instead.
        <div style={{ marginBottom: tokens.spacingVerticalL }}>
          <ProgressBar value={progressValue} />

          <Text style={{ display: 'block', marginTop: tokens.spacingVerticalXS }}>
            {progress.siteLabel ? `${siteLabel(progress.siteLabel)} · ` : ''}
            {progress.stageLabel}
            {/* Em dash reserved for the stage/target relation, so the hierarchy
                reads without being read. */}
            {progress.libraryTitle ? ` — ${progress.libraryTitle}` : ''}
          </Text>

          {/* Line 2 is REPLACED, not appended to, while paused. During a
              tenant-imposed wait the stage detail is stale by definition, and
              showing both invites averaging two contradictory signals. */}
          <Text
            style={{
              display: 'block',
              fontSize: tokens.fontSizeBase200,
              color: cancelRequested || throttled
                ? tokens.colorPaletteMarigoldForeground1
                : tokens.colorNeutralForeground3,
            }}
          >
            {cancelRequested
              ? 'Stopping — finishing the request already in flight…'
              : throttled
                ? `Paused — SharePoint is throttling requests. Waiting for it to allow more… (${throttleWait}s)`
                : (progress.detail ?? ' ')}
          </Text>

          <Text style={{ display: 'block', fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3 }}>
            {progress.libsTotal > 0 && `${progress.libsDone} of ${progress.libsTotal} libraries · `}
            {/* filesSeen, not `scanned` — files are identified the moment each
                page lands, whereas entries can only be published once their
                version sizes are final, so `scanned` sits at 0 for a whole
                library and reads as a stalled scan. The committed count is in
                the hover. */}
            {progress.filesSeen.toLocaleString()} files found
            {' · '}{formatElapsed(elapsed)}
            {etaSeconds != null && ` · ~${formatDuration(etaSeconds)} left`}
            {' '}
            <Tooltip
              relationship="label"
              content={{ children: <ScanStatusTooltip progress={progress} requestCount={requestCount} /> }}
            >
              <Info16Regular style={{ verticalAlign: 'middle', cursor: 'help' }} />
            </Tooltip>
          </Text>
        </div>
      )}

      {summary && (
        <>
          {viewedReport && (
            <Text style={{ display: 'block', color: tokens.colorNeutralForeground3, marginBottom: tokens.spacingVerticalS }}>
              Viewing saved scan from {new Date(viewedReport.timestamp).toLocaleString()} — {siteLabel(viewedReport.siteUrl)}
              {viewedReport.listingEvicted ? ' (file listing no longer stored — totals are complete)' : ''}
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
                <Text weight="semibold" style={{ display: 'block', fontSize: tokens.fontSizeBase500, color: tokens.colorBrandForeground1 }}>{formatBytes(summary.totalVersionSizeBytes ?? 0)}</Text>
                {/* Count is the same per-file number Excel's Details sheet
                    already had; this is just its sum, same treatment as the
                    size total above. */}
                <Text style={{ display: 'block', fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3 }}>
                  {(summary.totalVersionCount ?? 0).toLocaleString()} retained version{summary.totalVersionCount === 1 ? '' : 's'}
                </Text>
                {/* How this number was obtained, when it wasn't the free bulk
                    field. Worth surfacing because the difference between the
                    bulk side channel and the per-file pass is the difference
                    between a complete figure and a capped sample — and nothing
                    else on screen would tell the user which one they got.

                    Keyed on unmeasuredVersions, not on how the scan was
                    configured: every file that can have retained versions is
                    measured now, so anything left over means the pass was
                    interrupted or a list had no usable source — either way the
                    total is a floor and must say so. */}
                {summary.versionSizeStrategy === 'per-file' && (
                  <Text style={{ display: 'block', fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3 }}>
                    {(summary.unmeasuredVersions ?? 0) > 0
                      ? 'measured per file — incomplete, so this is a floor'
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
              {/* Only when the file listing genuinely isn't retrievable. This was
                  a "Partial" badge shown for any report over 50,000 rows, which
                  described a saved-listing detail in a word that reads as "your
                  numbers are incomplete" — they never were. */}
              {h.listingEvicted && (
                <Badge
                  appearance="tint"
                  color="informative"
                  title="Browser storage was full, so this report's file listing was removed to make room for newer reports. Its totals and comparisons are unaffected."
                >
                  No file list
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

      {(comparison || comparing) && (
        <div className={styles.statTile} style={{ marginTop: tokens.spacingVerticalM }}>
          <Text weight="semibold" style={{ display: 'block', marginBottom: tokens.spacingVerticalXS }}>Comparison</Text>
          {comparing && !comparison && (
            <Text style={{ display: 'block', color: tokens.colorNeutralForeground3 }}>Comparing…</Text>
          )}
          {comparison && (
            <>
              {comparison.crossSite && (
                <Text style={{ display: 'block', color: tokens.colorPaletteYellowForeground1 }}>
                  ⚠ These two reports are from different sites — the comparison may not be meaningful.
                </Text>
              )}
              <Text style={{ display: 'block' }}>
                Size change: {comparison.diff.sizeDeltaBytes >= 0 ? '+' : ''}{formatBytes(Math.abs(comparison.diff.sizeDeltaBytes))}
              </Text>
              {/* Undefined, not 0, when either file listing is missing — showing
                  "0 new" for a report whose listing was evicted would be
                  indistinguishable from "nothing changed". */}
              {comparison.diff.newStaleCount != null && (
                <Text style={{ display: 'block' }}>New archival candidates: {comparison.diff.newStaleCount.toLocaleString()}</Text>
              )}
              {comparison.diff.resolvedStaleCount != null && (
                <Text style={{ display: 'block' }}>Resolved (no longer stale): {comparison.diff.resolvedStaleCount.toLocaleString()}</Text>
              )}
              <Text style={{ display: 'block' }}>File count change: {comparison.diff.totalFilesDelta >= 0 ? '+' : ''}{comparison.diff.totalFilesDelta.toLocaleString()}</Text>
              {comparison.listingMissing && (
                <Text style={{ display: 'block', marginTop: tokens.spacingVerticalXS, color: tokens.colorNeutralForeground3, fontSize: tokens.fontSizeBase200 }}>
                  Per-file changes aren&apos;t shown because one of these reports no longer has its
                  file listing stored. The size and count changes above are from the saved totals and
                  are accurate.
                </Text>
              )}
            </>
          )}
        </div>
      )}
    </div>
  );
};
