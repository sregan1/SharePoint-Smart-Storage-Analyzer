import * as React from 'react';
import {
  Button,
  Body1,
  Checkbox,
  Spinner,
  ProgressBar,
  Text,
  Tab,
  TabList,
  Tooltip,
  MessageBar,
  MessageBarBody,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import { Document24Regular, Folder24Regular, ChevronRight16Regular, DocumentArrowDown24Regular, Info16Regular, Warning16Regular } from '@fluentui/react-icons';

import { StorageAnalyzerService } from '../services/StorageAnalyzerService';
import { ExcelExportService } from '../services/ExcelExportService';
import { CandidateTier, FolderListRow, FolderStorageNode, LibraryInfo, TreemapItem } from '../models/models';
import { pickDefaultLibrary } from '../services/sp/libraryStats';
import { Treemap, TierLegend } from './shared/Treemap';
import { StorageTable, StorageTableColumn } from './shared/StorageTable';
import { SizeBar } from './shared/SizeBar';
import { formatBytes, formatAge } from './shared/formatBytes';
import { tierColor, tierLabel } from './shared/tierBadge';
import { ageInDays, classify } from '../utils/archivalClassification';

const useStyles = makeStyles({
  root: {
    padding: tokens.spacingVerticalL,
    maxWidth: '1200px',
    margin: '0 auto',
    minHeight: '500px',
  },
  libraryRow: {
    display: 'flex',
    gap: tokens.spacingHorizontalS,
    flexWrap: 'wrap',
    marginBottom: tokens.spacingVerticalM,
  },
  breadcrumb: {
    display: 'flex',
    alignItems: 'center',
    flexWrap: 'wrap',
    gap: '2px',
    marginBottom: tokens.spacingVerticalS,
  },
});

interface FolderFileRow {
  name: string;
  serverRelativeUrl: string;
  sizeBytes: number;
  timeLastModified: string;
  authorDisplayName?: string;
  ageDays: number;
  tier: CandidateTier;
  versionSizeBytes?: number;
}

// Raw shape cached per folder — NOT the derived FolderFileRow. Caching the
// derived tier/ageDays would freeze a folder's file list at whatever
// staleDays/veryStaleDays were in effect the first time it loaded; changing
// the archival thresholds in Settings and revisiting a previously-viewed
// folder would then silently show stale (pun intended) classifications that
// disagree with the legend. Deriving at read time keeps it always current.
interface RawFolderFile {
  name: string;
  serverRelativeUrl: string;
  sizeBytes: number;
  timeLastModified: string;
  authorDisplayName?: string;
  versionSizeBytes?: number;
}

export interface ExplorerViewProps {
  sp: StorageAnalyzerService;
  excel: ExcelExportService;
  siteUrl: string;
  staleDays: number;
  veryStaleDays: number;
}

export const ExplorerView: React.FC<ExplorerViewProps> = ({
  sp, excel, siteUrl, staleDays, veryStaleDays,
}) => {
  const styles = useStyles();

  const [libraries, setLibraries] = React.useState<LibraryInfo[]>([]);
  const [librariesLoading, setLibrariesLoading] = React.useState(true);
  const [libraryUrl, setLibraryUrl] = React.useState<string | undefined>();
  const [error, setError] = React.useState('');
  const [viewMode, setViewMode] = React.useState<'treemap' | 'list'>('treemap');

  const [selectedUrl, setSelectedUrl] = React.useState<string>('');
  const [childrenLoading, setChildrenLoading] = React.useState(false);
  const [childrenProgress, setChildrenProgress] = React.useState<{ done: number; total: number } | null>(null);
  const [cacheVersion, setCacheVersion] = React.useState(0);
  const childrenCache = React.useRef<Map<string, FolderStorageNode[]>>(new Map());
  const parentOf = React.useRef<Map<string, { name: string; url: string }>>(new Map());

  const [selectedFiles, setSelectedFiles] = React.useState<FolderFileRow[]>([]);
  const [filesLoading, setFilesLoading] = React.useState(false);
  const filesCache = React.useRef<Map<string, RawFolderFile[]>>(new Map());
  const [includeVersions, setIncludeVersions] = React.useState(false);

  // Load the library list once per site and seed the default document library.
  React.useEffect(() => {
    let cancelled = false;
    setLibrariesLoading(true);
    sp.getLibrariesWithStats(siteUrl, false)
      .then((libs) => {
        if (cancelled) return;
        setLibraries(libs);
        setLibraryUrl(pickDefaultLibrary(libs)?.serverRelativeUrl);
      })
      .catch((err: any) => { if (!cancelled) setError(`Failed to load libraries: ${err?.message ?? String(err)}`); })
      .finally(() => { if (!cancelled) setLibrariesLoading(false); });
    return () => { cancelled = true; };
  }, [siteUrl]);

  // Reset drill-down state when the selected library changes.
  React.useEffect(() => {
    if (!libraryUrl) return;
    childrenCache.current.clear();
    parentOf.current.clear();
    filesCache.current.clear();
    setSelectedUrl(libraryUrl);
  }, [libraryUrl]);

  const currentLibrary = libraries.find((l) => l.serverRelativeUrl === libraryUrl);

  const loadChildren = React.useCallback((url: string): void => {
    if (childrenCache.current.has(url)) return;
    setChildrenLoading(true);
    setChildrenProgress(null);
    // The Treemap/List only render once the whole batch is in (see
    // loadingIndicator below), so the per-child callback here only drives
    // the progress counter — a cache hit skips it entirely, going straight
    // to the .then() with the complete result.
    sp.getFolderChildren(siteUrl, url, (done, total) => { setChildrenProgress({ done, total }); })
      .then((children) => {
        setError('');
        childrenCache.current.set(url, children);
        children.forEach((c) => parentOf.current.set(c.serverRelativeUrl, { name: c.name, url }));
      })
      .catch((err: any) => {
        // Still cache [] so this folder doesn't re-fire the load on every
        // render, but surface the failure — silently treating a fetch
        // failure as "this folder is empty" would misreport a permissions
        // error or throttling exhaustion as an actual empty folder.
        childrenCache.current.set(url, []);
        setError(`Failed to load folder contents: ${err?.message ?? String(err)}`);
      })
      .finally(() => {
        // Order matters here: React 17 does not batch state updates made
        // from a promise callback, so each of these setState calls commits
        // as its own separate render. selectedChildren is a useMemo keyed on
        // cacheVersion (see below), not on the childrenCache ref mutation
        // above — so bumping cacheVersion AFTER setChildrenLoading(false)
        // would open the loading gate (childrenLoading || filesLoading) one
        // render before selectedChildren actually reflects the new folder,
        // showing files (already loaded) with the previous folder's
        // subfolders, or none, until the next render catches up. Bumping
        // cacheVersion first means selectedChildren is already current by
        // the time the render that clears childrenLoading actually happens.
        setChildrenProgress(null);
        setCacheVersion((v) => v + 1);
        setChildrenLoading(false);
      });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [siteUrl, libraryUrl]);

  // Fetch immediate files whenever the selected folder changes, or the
  // version-history toggle changes. Cache key includes the toggle state
  // since raw files fetched with it off never have versionSizeBytes.
  React.useEffect(() => {
    if (!selectedUrl) return;
    const toRows = (files: RawFolderFile[]): FolderFileRow[] => files.map((f) => {
      const ageDays = ageInDays(f.timeLastModified);
      return {
        name: f.name,
        serverRelativeUrl: f.serverRelativeUrl,
        sizeBytes: f.sizeBytes,
        timeLastModified: f.timeLastModified,
        authorDisplayName: f.authorDisplayName,
        ageDays,
        tier: classify(ageDays, staleDays, veryStaleDays),
        versionSizeBytes: f.versionSizeBytes,
      };
    });
    const cacheKey = `${selectedUrl}::${includeVersions}`;
    const cached = filesCache.current.get(cacheKey);
    if (cached) { setSelectedFiles(toRows(cached)); return; }
    // Clear the previous folder's rows before setting filesLoading — this
    // effect runs one render after selectedUrl actually changes, and without
    // this, that render would still hold the OLD folder's files (loading
    // flags haven't flipped true yet either) and could briefly paint stale
    // files against the new (still-empty) folder list.
    setSelectedFiles([]);
    setFilesLoading(true);
    sp.getFolderFiles(siteUrl, selectedUrl, includeVersions)
      .then((files) => {
        setError('');
        filesCache.current.set(cacheKey, files);
        setSelectedFiles(toRows(files));
      })
      .catch((err: any) => {
        // Same reasoning as loadChildren's catch: silently showing an empty
        // file list here would misreport a fetch failure (throttling, a
        // transient error) as "this folder is genuinely empty" — surface it
        // instead. Not cached, so revisiting the folder retries the fetch.
        setSelectedFiles([]);
        setError(`Failed to load folder files: ${err?.message ?? String(err)}`);
      })
      .finally(() => setFilesLoading(false));
  }, [selectedUrl, siteUrl, staleDays, veryStaleDays, includeVersions]);

  // Ensure the currently selected folder's children are loaded (needed for
  // both the treemap and the list view).
  React.useEffect(() => {
    if (selectedUrl) loadChildren(selectedUrl);
  }, [selectedUrl, loadChildren]);

  // Breadcrumb: walk parentOf from the selected folder back to the library root.
  const breadcrumb = React.useMemo(() => {
    const trail: { name: string; url: string }[] = [];
    let cursor = selectedUrl;
    let guard = 0;
    while (cursor && cursor !== libraryUrl && guard++ < 50) {
      const p = parentOf.current.get(cursor);
      if (!p) break;
      trail.unshift({ name: findName(cursor), url: cursor });
      cursor = p.url;
    }
    if (currentLibrary) trail.unshift({ name: currentLibrary.title, url: libraryUrl ?? '' });
    return trail;

    function findName(url: string): string {
      for (const list of childrenCache.current.values()) {
        const match = list.find((n) => n.serverRelativeUrl === url);
        if (match) return match.name;
      }
      return url.split('/').pop() ?? url;
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedUrl, libraryUrl, currentLibrary]);

  // Keyed on cacheVersion (bumped only when childrenCache actually changes)
  // rather than recomputed unconditionally, so an unrelated re-render (e.g.
  // filesLoading toggling) doesn't hand the Treemap a new array reference
  // and force it to redo the full squarify layout for unchanged data.
  const selectedChildren = React.useMemo(
    () => (selectedUrl ? childrenCache.current.get(selectedUrl) ?? [] : []),
    [selectedUrl, cacheVersion],
  );

  // File cells are sized by file + version history combined (when the
  // toggle is on and that file's version fetch succeeded) so the square
  // reflects real total storage weight, not just current content — Treemap
  // then re-derives the breakdown from sizeBytes/versionSizeBytes to label
  // and tooltip each cell. Folders can't join in: StorageMetrics doesn't
  // report version-history totals, so a folder's own recursive rollup never
  // includes it (see getFolderChildren/getStorageMetrics) — folder cells
  // keep showing file-content size only.
  const treemapItems: TreemapItem[] = React.useMemo(() => [
    ...selectedChildren.map((c) => ({
      id: c.serverRelativeUrl,
      label: c.name,
      sizeBytes: c.totalSizeBytes,
      kind: 'folder' as const,
      lastModified: c.lastModified,
      itemCount: c.fileCount,
      sizeUnknown: c.sizeSource === 'error',
    })),
    ...selectedFiles.map((f) => ({
      id: f.serverRelativeUrl,
      label: f.name,
      sizeBytes: f.sizeBytes + (includeVersions && f.versionSizeBytes != null ? f.versionSizeBytes : 0),
      kind: 'file' as const,
      tier: f.tier,
      lastModified: f.timeLastModified,
      versionSizeBytes: includeVersions ? f.versionSizeBytes : undefined,
    })),
  ], [selectedChildren, selectedFiles, includeVersions]);

  // Count of immediate subfolders whose size couldn't be determined (usually
  // throttling exhausted mid-walk) — surfaced as a warning rather than left
  // silent, since those folders otherwise look identical to genuinely empty
  // ones (see getFolderChildren's 'error' sizeSource).
  const unknownSizeFolderCount = React.useMemo(
    () => selectedChildren.filter((c) => c.sizeSource === 'error').length,
    [selectedChildren],
  );

  const listRows: FolderListRow[] = React.useMemo(() => [
    ...selectedChildren.map((c) => ({
      kind: 'folder' as const,
      name: c.name,
      serverRelativeUrl: c.serverRelativeUrl,
      sizeBytes: c.totalSizeBytes,
      lastModified: c.lastModified,
      itemCount: c.fileCount,
      sizeUnknown: c.sizeSource === 'error',
    })),
    ...selectedFiles.map((f) => ({
      kind: 'file' as const,
      name: f.name,
      serverRelativeUrl: f.serverRelativeUrl,
      sizeBytes: f.sizeBytes,
      lastModified: f.timeLastModified,
      ageDays: f.ageDays,
      tier: f.tier,
      authorDisplayName: f.authorDisplayName,
      versionSizeBytes: f.versionSizeBytes,
    })),
  ], [selectedChildren, selectedFiles]);

  // Folders have no versionSizeBytes — a recursive rollup isn't available
  // without a full walk (StorageMetrics doesn't report version-history
  // totals), so only leaf files in the current folder ever get a real
  // value; folder rows render '—' for this column.
  const versionsTotalBytes = React.useMemo(
    () => selectedFiles.reduce((sum, f) => sum + (f.versionSizeBytes ?? 0), 0),
    [selectedFiles],
  );
  // reduce, not Math.max(1, ...listRows.map(...)) — spreading tens of
  // thousands of arguments into Math.max risks a RangeError on very large
  // folders.
  const maxListSize = React.useMemo(
    () => listRows.reduce((max, r) => Math.max(max, r.sizeBytes), 1),
    [listRows],
  );

  const listColumns: StorageTableColumn<FolderListRow>[] = [
    {
      key: 'name',
      header: 'Name',
      sortValue: (r) => r.name,
      render: (r) => (r.kind === 'folder' ? (
        <Button
          appearance="transparent"
          size="small"
          icon={<Folder24Regular style={{ fontSize: '14px' }} />}
          onClick={() => setSelectedUrl(r.serverRelativeUrl)}
          style={{ padding: 0, minWidth: 0, justifyContent: 'flex-start' }}
        >
          {r.name}
        </Button>
      ) : (
        <span style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
          <Document24Regular style={{ fontSize: '14px', flexShrink: 0 }} />
          <span>{r.name}</span>
        </span>
      )),
    },
    {
      key: 'size',
      header: 'Size',
      align: 'right',
      // Sort unknown-size folders as smallest rather than 0 — a confirmed
      // empty folder and an unmeasurable one shouldn't be indistinguishable
      // in sort order either.
      sortValue: (r) => (r.sizeUnknown ? -1 : r.sizeBytes),
      render: (r) => (r.sizeUnknown ? (
        <Tooltip content="Couldn't determine this folder's size — likely SharePoint throttling or a transient error, not a confirmed empty folder. Try again in a moment." relationship="label">
          <span style={{ display: 'inline-flex', alignItems: 'center', gap: '4px', color: tokens.colorPaletteMarigoldForeground1, cursor: 'help' }}>
            <Warning16Regular /> Unknown
          </span>
        </Tooltip>
      ) : (
        <div style={{ display: 'flex', alignItems: 'center', gap: '8px', justifyContent: 'flex-end' }}>
          <span style={{ minWidth: '64px', textAlign: 'right' }}>{formatBytes(r.sizeBytes)}</span>
          <div style={{ width: '60px' }}><SizeBar value={r.sizeBytes} max={maxListSize} /></div>
        </div>
      )),
    },
    {
      key: 'items',
      header: 'Items',
      align: 'right',
      sortValue: (r) => r.itemCount ?? -1,
      render: (r) => <span>{r.kind === 'folder' ? (r.sizeUnknown ? '—' : r.itemCount ?? '—') : ''}</span>,
    },
    ...(includeVersions ? [{
      key: 'versionSize',
      // The Size column above is content size only — SharePoint's folder
      // rollup (StorageMetrics / live walk) has no recursive version-history
      // total, so a folder's Size never includes it. Only individual files
      // get a real number here; folders always show '—'.
      header: 'Version history (files only)',
      align: 'right' as const,
      sortValue: (r: FolderListRow) => r.versionSizeBytes ?? -1,
      render: (r: FolderListRow) => <span>{r.kind === 'file' && r.versionSizeBytes !== undefined ? formatBytes(r.versionSizeBytes) : '—'}</span>,
    }] : []),
    {
      key: 'modified',
      header: 'Modified',
      sortValue: (r) => r.lastModified ?? '',
      render: (r) => (r.lastModified ? (
        <span>
          {new Date(r.lastModified).toLocaleDateString()}
          {r.kind === 'file' && r.ageDays != null ? ` (${formatAge(r.ageDays)} ago)` : ''}
        </span>
      ) : <span>—</span>),
    },
    {
      key: 'tier',
      header: 'Status',
      sortValue: (r) => r.ageDays ?? -1,
      render: (r) => (r.kind === 'file' && r.tier ? (
        <span style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
          <span style={{ width: '8px', height: '8px', borderRadius: '50%', background: tierColor(r.tier), flexShrink: 0 }} />
          <span>{tierLabel(r.tier)}</span>
        </span>
      ) : <span>—</span>),
    },
  ];

  const contextLabel = `${currentLibrary?.title ?? ''}${breadcrumb.length > 1 ? ' / ' + breadcrumb.slice(1).map((b) => b.name).join(' / ') : ''}`;
  // Excel export lazy-loads a chunk and writes a workbook, so unlike CSV it
  // has real ways to fail — left un-caught (as `void`-fired before), a
  // failure was an invisible unhandled promise rejection with no user-facing
  // signal that anything went wrong.
  const handleExportExcel = async (): Promise<void> => {
    try {
      setError('');
      await excel.exportFolderListing(listRows, contextLabel);
    } catch (err: any) {
      setError(`Excel export failed: ${err?.message ?? String(err)}`);
    }
  };
  const handleExportCsv = (): void => { excel.exportFolderListingCsv(listRows, contextLabel); };

  // Gates the Treemap/List while a folder is loading — a real progress bar
  // with a live count once the folder total is known (right after the fast
  // listing call resolves), a plain spinner before that or while the
  // separate immediate-files fetch is still pending.
  const loadingIndicator = childrenProgress && childrenProgress.total > 0 ? (
    <div style={{ marginTop: tokens.spacingVerticalM }}>
      <ProgressBar value={childrenProgress.done / childrenProgress.total} />
      <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3 }}>
        Loading folder sizes… {childrenProgress.done} of {childrenProgress.total}
      </Text>
    </div>
  ) : (
    <Spinner size="small" label="Loading folder contents…" style={{ marginTop: tokens.spacingVerticalM }} />
  );

  return (
    <div className={styles.root}>
      {error && (
        <MessageBar intent="error" style={{ marginBottom: tokens.spacingVerticalM }}>
          <MessageBarBody>{error}</MessageBarBody>
        </MessageBar>
      )}

      {!childrenLoading && !filesLoading && unknownSizeFolderCount > 0 && (
        <MessageBar intent="warning" style={{ marginBottom: tokens.spacingVerticalM }}>
          <MessageBarBody>
            {unknownSizeFolderCount} folder{unknownSizeFolderCount === 1 ? '' : 's'} in this view could not be measured
            (likely SharePoint throttling) and show as "Unknown" rather than a real size — they are not confirmed empty.
            Try again in a moment, or lower Concurrent API requests in Settings.
          </MessageBarBody>
        </MessageBar>
      )}

      {librariesLoading ? (
        <Spinner size="small" label="Loading libraries…" />
      ) : libraries.length === 0 ? (
        <Body1 style={{ color: tokens.colorNeutralForeground3 }}>No document libraries found on this site.</Body1>
      ) : (
        <>
          <div className={styles.libraryRow}>
            {libraries.map((l) => (
              <Button
                key={l.serverRelativeUrl}
                appearance={l.serverRelativeUrl === libraryUrl ? 'primary' : 'secondary'}
                size="small"
                onClick={() => setLibraryUrl(l.serverRelativeUrl)}
              >
                {l.title}
              </Button>
            ))}
          </div>

          <div className={styles.breadcrumb}>
            {breadcrumb.map((b, i) => (
              <React.Fragment key={b.url}>
                {i > 0 && <ChevronRight16Regular style={{ color: tokens.colorNeutralForeground3 }} />}
                <Button appearance="transparent" size="small" onClick={() => setSelectedUrl(b.url)}>
                  {b.name}
                </Button>
              </React.Fragment>
            ))}
          </div>

          <TabList selectedValue={viewMode} onTabSelect={(_, d) => setViewMode(d.value as 'treemap' | 'list')} style={{ marginBottom: tokens.spacingVerticalM }}>
            <Tab value="treemap">Treemap</Tab>
            <Tab value="list">List</Tab>
          </TabList>

          <div style={{ marginBottom: tokens.spacingVerticalS, display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalXS }}>
            <Checkbox
              label="Include version history size (slower)"
              checked={includeVersions}
              onChange={(_, d) => setIncludeVersions(!!d.checked)}
            />
            <Tooltip
              content="Only individual files get a real version-history number — SharePoint's folder size rollup has no recursive version-history total, so a folder's Size never includes its files' version history, on or off."
              relationship="label"
            >
              <Info16Regular style={{ cursor: 'help', color: tokens.colorNeutralForeground3 }} />
            </Tooltip>
          </div>

          {viewMode === 'treemap' ? (
            <div>
              <TierLegend staleDays={staleDays} veryStaleDays={veryStaleDays} />
              {includeVersions && (
                <Text style={{ display: 'block', color: tokens.colorNeutralForeground3, fontSize: tokens.fontSizeBase200, marginBottom: tokens.spacingVerticalS }}>
                  File squares are sized by file + version history combined ({formatBytes(versionsTotalBytes)} in version history across this folder's files); folders still show file size only — SharePoint has no recursive version-history rollup.
                </Text>
              )}
              {childrenLoading || filesLoading ? (
                loadingIndicator
              ) : (
                <>
                  <Treemap
                    items={treemapItems}
                    onFolderClick={(item) => setSelectedUrl(item.id)}
                  />
                  {treemapItems.length === 0 && !error && (
                    <Body1 style={{ color: tokens.colorNeutralForeground3, marginTop: tokens.spacingVerticalM }}>
                      This folder has no subfolders or files.
                    </Body1>
                  )}
                </>
              )}
            </div>
          ) : (
            <div>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: tokens.spacingHorizontalS, marginBottom: tokens.spacingVerticalS }}>
                <TierLegend staleDays={staleDays} veryStaleDays={veryStaleDays} />
                <div style={{ display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalS }}>
                  <Button icon={<DocumentArrowDown24Regular />} appearance="secondary" size="small" onClick={handleExportExcel} disabled={listRows.length === 0}>
                    Export Excel
                  </Button>
                  <Button icon={<DocumentArrowDown24Regular />} appearance="secondary" size="small" onClick={handleExportCsv} disabled={listRows.length === 0}>
                    Export CSV
                  </Button>
                </div>
              </div>
              {includeVersions && (
                <Text style={{ display: 'block', color: tokens.colorNeutralForeground3, fontSize: tokens.fontSizeBase200, marginBottom: tokens.spacingVerticalS }}>
                  This folder's files: {formatBytes(versionsTotalBytes)} in version history (files only — folders show '—', no recursive rollup).
                </Text>
              )}
              {childrenLoading || filesLoading ? (
                loadingIndicator
              ) : listRows.length === 0 && !error ? (
                <Body1 style={{ color: tokens.colorNeutralForeground3 }}>This folder is empty.</Body1>
              ) : (
                <StorageTable rows={listRows} columns={listColumns} getRowKey={(r) => r.serverRelativeUrl} defaultSortKey="size" />
              )}
            </div>
          )}
        </>
      )}
    </div>
  );
};
