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
import { Document24Regular, Folder24Regular, ChevronRight16Regular, DocumentArrowDown24Regular, Info16Regular, Warning16Regular, ArrowClockwise20Regular, ArrowLeft24Regular, Delete16Regular } from '@fluentui/react-icons';

import { StorageAnalyzerService } from '../services/StorageAnalyzerService';
import { ExcelExportService } from '../services/ExcelExportService';
import { CandidateTier, FolderListRow, FolderStorageNode, LibraryInfo, TreemapItem } from '../models/models';
import { MAX_TREEMAP_CELLS } from '../utils/treemapLayout';
import { LibraryRollup } from '../services/sp/libraryStats';
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
  versionCount?: number;
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
  versionCount?: number;
}

export interface ExplorerViewProps {
  sp: StorageAnalyzerService;
  excel: ExcelExportService;
  siteUrl: string;
  staleDays: number;
  veryStaleDays: number;
  // Which of the two internal tabs to open on — set from the Home screen card
  // (Tree View vs List View) the user picked, or the property pane's default
  // view. The tab itself remains switchable afterward either way.
  initialViewMode?: 'treemap' | 'list';
  // Present only when reached via the Home screen, so a defaultView of
  // 'tree'/'list' (which skips Home entirely) doesn't show a Back button with
  // nowhere sensible to return to.
  onBack?: () => void;
}

export const ExplorerView: React.FC<ExplorerViewProps> = ({
  sp, excel, siteUrl, staleDays, veryStaleDays, initialViewMode, onBack,
}) => {
  const styles = useStyles();

  const [libraries, setLibraries] = React.useState<LibraryInfo[]>([]);
  const [librariesLoading, setLibrariesLoading] = React.useState(true);
  // undefined = the site-root view: every library rendered as a treemap square
  // sized by its own rollup, which is where the Explorer now opens. Picking a
  // library (square, button row, or breadcrumb) sets this and drills in.
  const [libraryUrl, setLibraryUrl] = React.useState<string | undefined>();
  const [rollups, setRollups] = React.useState<LibraryRollup[]>([]);
  // Starts true, and only the effect's finally clears it. Initialising to false
  // would leave a render where the libraries have arrived but their rollups
  // haven't started resolving — the same "nothing is loading, so render an
  // empty treemap" gray-box gap fixed at folder level.
  const [rollupsLoading, setRollupsLoading] = React.useState(true);
  const atRoot = libraryUrl === undefined;
  const [error, setError] = React.useState('');
  const [viewMode, setViewMode] = React.useState<'treemap' | 'list'>(initialViewMode ?? 'treemap');

  const [selectedUrl, setSelectedUrl] = React.useState<string>('');
  const [childrenLoading, setChildrenLoading] = React.useState(false);
  const [cacheVersion, setCacheVersion] = React.useState(0);
  const childrenCache = React.useRef<Map<string, FolderStorageNode[]>>(new Map());
  const parentOf = React.useRef<Map<string, { name: string; url: string }>>(new Map());

  const [selectedFiles, setSelectedFiles] = React.useState<FolderFileRow[]>([]);
  const [filesLoading, setFilesLoading] = React.useState(false);
  const filesCache = React.useRef<Map<string, RawFolderFile[]>>(new Map());
  const [includeVersions, setIncludeVersions] = React.useState(false);
  // Which files request the current selectedFiles actually belong to. Compared
  // against the live key below to answer "are the files on screen the ones for
  // the folder now selected?" — a plain `filesLoading` flag can't, because it
  // is set by an effect that runs a render AFTER selectedUrl changes.
  const [loadedFilesKey, setLoadedFilesKey] = React.useState('');
  // Seconds spent on the current load, ticked purely so the loading UI always
  // has something visibly moving. The item counter advances once per
  // 5,000-item page, which on a throttled tenant can be a long gap — long
  // enough to read as a frozen app without a second-by-second signal.
  const [loadElapsed, setLoadElapsed] = React.useState(0);
  const [throttled, setThrottled] = React.useState(false);
  // Bumped by Refresh. Included in the load effects' deps because clearing the
  // caches alone can't re-trigger them — and loadChildren early-returns on a
  // cache hit, so a stale/failed result would otherwise stick for the whole
  // session no matter how many times the user re-navigated to it.
  const [refreshToken, setRefreshToken] = React.useState(0);

  // ── Cancellation / live scan progress ────────────────────────────────────
  // ONE controller for whichever measurement is in flight (root rollups or a
  // folder drill-down) — a single controller enforces "one measurement at a
  // time" and gives Cancel an unambiguous target without needing to know
  // which view is loading.
  const scanAbortRef = React.useRef<AbortController | null>(null);
  // Written by the data layer's throttled onWalkProgress (potentially
  // hundreds of times a second on a huge site); flushed into state by the
  // existing elapsed-time ticker rather than driving a render per page.
  const itemsReadRef = React.useRef(0);
  const [itemsRead, setItemsRead] = React.useState(0);
  const [cancelRequested, setCancelRequested] = React.useState(false);
  // Keyed by scan target ('' = site root, else a folder url) so a canceled
  // scan's banner follows that specific folder/root across navigation rather
  // than a lone boolean bleeding onto whatever is opened next.
  const [canceledKeys, setCanceledKeys] = React.useState<ReadonlySet<string>>(new Set());
  // Restarts the root rollup scan when the user returns to the site root
  // after abandoning it mid-flight by drilling into a library.
  const [rollupsToken, setRollupsToken] = React.useState(0);
  const [rollupsAbandoned, setRollupsAbandoned] = React.useState(false);

  // Starts a measurement, superseding any in flight (aborting it — not the
  // same as a user Cancel, see isCurrentScan below). Returns the controller;
  // callers compare scanAbortRef.current against it to answer "are my
  // results still wanted?"
  const beginScan = React.useCallback((key: string): AbortController => {
    if (scanAbortRef.current) scanAbortRef.current.abort();
    const ctrl = new AbortController();
    scanAbortRef.current = ctrl;
    itemsReadRef.current = 0;
    setItemsRead(0);
    setCancelRequested(false);
    setCanceledKeys((prev) => {
      if (!prev.has(key)) return prev; // same reference — no re-render
      const next = new Set(prev);
      next.delete(key);
      return next;
    });
    return ctrl;
  }, []);

  // True only while `ctrl` is still THE current scan. False means a newer
  // scan (or a navigation away) replaced it, so its results belong to a
  // screen the user has left and must not be written anywhere. This is what
  // distinguishes "normal completion" and "user canceled" (ref left pointing
  // at ctrl in both cases) from "superseded/navigated away" (ref points
  // elsewhere or is null).
  const isCurrentScan = (ctrl: AbortController): boolean => scanAbortRef.current === ctrl;

  const handleCancel = (): void => {
    if (!scanAbortRef.current || cancelRequested) return;
    setCancelRequested(true);
    // Deliberately does NOT null the ref — staying current is what tells the
    // settling promise these partials were asked for and should be kept.
    scanAbortRef.current.abort();
  };

  // Abort whatever is measuring when the user navigates — a different
  // folder, library, breadcrumb hop, the Site button, Refresh, or unmount.
  // Nulling the ref marks the in-flight scan non-current, so its
  // .then/.finally discard rather than writing into the screen the user has
  // moved to. Without this, an abandoned walk on a huge library would run
  // indefinitely (there's no deadline any more), competing for the same
  // request ceiling as the walk the user is actually waiting on.
  React.useEffect(() => () => {
    if (scanAbortRef.current) {
      scanAbortRef.current.abort();
      scanAbortRef.current = null;
    }
  }, [selectedUrl, libraryUrl, siteUrl, refreshToken]);

  // Load the library list once per site, then resolve each library's rollup for
  // the site-root treemap. No default library is auto-selected any more — the
  // root view IS the landing screen, and it's more informative than dropping
  // straight into one library (it answers "which library is the storage in?"
  // before you have to guess).
  // Navigation resets only when the SITE changes — deliberately separate from
  // the loader below, which also re-runs on Refresh. Folding this into that
  // effect would make Refresh yank the user back to the site root instead of
  // reloading the folder they're actually looking at.
  React.useEffect(() => { setLibraryUrl(undefined); }, [siteUrl]);

  // Libraries only — split from the rollup effect below so the rollup scan
  // can be independently re-triggered (e.g. restarted after being abandoned
  // mid-flight) without re-fetching the library list itself.
  React.useEffect(() => {
    let cancelled = false;
    setLibrariesLoading(true);
    // Plain list, not getLibrariesWithStats — sizes come from getLibraryRollups
    // below. getLibrariesWithStats would, on a site with no library named
    // "Documents", full-walk every library to pick a default by size, which
    // is the whole-site walk this view is designed to avoid.
    sp.getLibraries(siteUrl, false)
      .then((libs) => {
        if (cancelled) return;
        setLibraries(libs);
      })
      .catch((err: any) => { if (!cancelled) setError(`Failed to load libraries: ${err?.message ?? String(err)}`); })
      .finally(() => { if (!cancelled) setLibrariesLoading(false); });
    return () => { cancelled = true; };
  }, [siteUrl, refreshToken]);

  // Root rollups. Runs after the library list resolves; restarts on Refresh
  // or when rollupsToken is bumped (see the "restart on return" effect
  // below). No automatic time/count limit any more — the user cancels via
  // the Cancel button, and SpApiClient's own throttle governor is what
  // actually protects the tenant on a large site.
  React.useEffect(() => {
    if (librariesLoading) return;
    if (libraries.length === 0) { setRollupsLoading(false); return; }
    const ctrl = beginScan('');
    setRollupsLoading(true);
    sp.getLibraryRollups(siteUrl, libraries, {
      signal: ctrl.signal,
      onWalkProgress: (visited) => { if (visited > itemsReadRef.current) itemsReadRef.current = visited; },
    })
      .then((r) => {
        if (!isCurrentScan(ctrl)) { setRollupsAbandoned(true); return; }
        setRollups(r);
        if (ctrl.signal.aborted) setCanceledKeys((prev) => new Set(prev).add(''));
      })
      .catch((err: any) => {
        if (!isCurrentScan(ctrl)) return;
        setError(`Failed to load libraries: ${err?.message ?? String(err)}`);
      })
      .finally(() => {
        if (!isCurrentScan(ctrl)) return;
        scanAbortRef.current = null;
        setCancelRequested(false);
        setRollupsLoading(false);
      });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [siteUrl, refreshToken, rollupsToken, libraries, librariesLoading]);

  // Returning to the site root after the library scan was abandoned
  // mid-flight (the user drilled into a library before it finished) restarts
  // it — otherwise the root treemap would permanently show whatever partial
  // set existed at the moment of drilling in, since nothing else would ever
  // change this effect's deps again. Deliberately NOT triggered by a user
  // Cancel (that lands in canceledKeys instead, showing a banner rather than
  // silently restarting) — cancel means stopped, not "retries when I look
  // away and back".
  React.useEffect(() => {
    if (atRoot && rollupsAbandoned) {
      setRollupsAbandoned(false);
      setRollupsToken((t) => t + 1);
    }
  }, [atRoot, rollupsAbandoned]);

  // Reset drill-down state when the selected library changes.
  React.useEffect(() => {
    if (!libraryUrl) {
      // Back at the site root — drop the folder selection so the folder-level
      // loaders stay idle and the readiness gate doesn't think it's mid-load.
      setSelectedUrl('');
      return;
    }
    childrenCache.current.clear();
    parentOf.current.clear();
    filesCache.current.clear();
    setSelectedUrl(libraryUrl);
  }, [libraryUrl]);

  const currentLibrary = libraries.find((l) => l.serverRelativeUrl === libraryUrl);

  const loadChildren = React.useCallback((url: string): void => {
    if (childrenCache.current.has(url)) return;
    if (!currentLibrary) return; // no library selected — nothing to sweep
    const ctrl = beginScan(url);
    setChildrenLoading(true);
    // The Treemap/List only render once the whole batch is in (see
    // loadingIndicator below), so the per-child callback here only drives
    // the progress counter — a cache hit skips it entirely, going straight
    // to the .then() with the complete result.
    sp.getFolderChildren(
      siteUrl,
      currentLibrary,
      url,
      undefined,
      {
        signal: ctrl.signal,
        // Ref only — no setState here. This fires once per 5,000-item page;
        // the ticker below flushes it into state at a paint-friendly rate.
        onWalkProgress: (items) => { if (items > itemsReadRef.current) itemsReadRef.current = items; },
      },
    )
      .then((children) => {
        if (!isCurrentScan(ctrl)) return; // superseded/navigated away — drop it
        setError('');
        // Partials are cached in memory deliberately: the readiness gate
        // needs *something* in the cache or the loading indicator would spin
        // forever. canceledKeys is what marks this copy as incomplete rather
        // than exact (the underlying aggregate cache refuses to keep a
        // canceled sweep, so Refresh genuinely re-measures).
        childrenCache.current.set(url, children);
        children.forEach((c) => parentOf.current.set(c.serverRelativeUrl, { name: c.name, url }));
        if (ctrl.signal.aborted) setCanceledKeys((prev) => new Set(prev).add(url));
      })
      .catch((err: any) => {
        if (!isCurrentScan(ctrl)) return;
        // Still cache [] so this folder doesn't re-fire the load on every
        // render, but surface the failure — silently treating a fetch
        // failure as "this folder is empty" would misreport a permissions
        // error or throttling exhaustion as an actual empty folder.
        childrenCache.current.set(url, []);
        setError(`Failed to load folder contents: ${err?.message ?? String(err)}`);
      })
      .finally(() => {
        if (!isCurrentScan(ctrl)) return; // do not clobber a newer scan
        scanAbortRef.current = null;
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
        setCancelRequested(false);
        setCacheVersion((v) => v + 1);
        setChildrenLoading(false);
      });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [siteUrl, libraryUrl, currentLibrary, beginScan]);

  // Fetch immediate files whenever the selected folder changes, or the
  // version-history toggle changes. Cache key includes the toggle state
  // since raw files fetched with it off never have versionSizeBytes.
  React.useEffect(() => {
    if (!selectedUrl || !currentLibrary) return;
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
        versionCount: f.versionCount,
      };
    });
    // Version history (size and an approximate count) is always included in
    // the fetch now at no extra cost — see storageMetrics.getFolderFiles —
    // so the cache key and this effect no longer need to vary by
    // includeVersions at all; that toggle only controls DISPLAY below.
    const cacheKey = selectedUrl;
    const cached = filesCache.current.get(cacheKey);
    if (cached) { setSelectedFiles(toRows(cached)); setLoadedFilesKey(cacheKey); return; }
    // Clear the previous folder's rows before setting filesLoading — this
    // effect runs one render after selectedUrl actually changes, and without
    // this, that render would still hold the OLD folder's files (loading
    // flags haven't flipped true yet either) and could briefly paint stale
    // files against the new (still-empty) folder list.
    setSelectedFiles([]);
    setFilesLoading(true);
    sp.getFolderFiles(siteUrl, currentLibrary, selectedUrl)
      .then((files) => {
        setError('');
        filesCache.current.set(cacheKey, files);
        setSelectedFiles(toRows(files));
        setLoadedFilesKey(cacheKey);
      })
      .catch((err: any) => {
        // Same reasoning as loadChildren's catch: silently showing an empty
        // file list here would misreport a fetch failure (throttling, a
        // transient error) as "this folder is genuinely empty" — surface it
        // instead. Not cached, so revisiting the folder retries the fetch.
        setSelectedFiles([]);
        setError(`Failed to load folder files: ${err?.message ?? String(err)}`);
        // Mark resolved even on failure, or the readiness gate below would
        // hold the loading indicator forever and the error banner explaining
        // what happened would never be reachable.
        setLoadedFilesKey(cacheKey);
      })
      .finally(() => setFilesLoading(false));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedUrl, siteUrl, currentLibrary, staleDays, veryStaleDays, refreshToken]);

  // Ensure the currently selected folder's children are loaded (needed for
  // both the treemap and the list view).
  React.useEffect(() => {
    if (selectedUrl) loadChildren(selectedUrl);
  }, [selectedUrl, loadChildren, refreshToken]);

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
    // Site root, always first: the library-level treemap is a real level in
    // the drill-down now, so it needs a way back. Empty url is the sentinel
    // the click handler maps to "clear the library selection".
    trail.unshift({ name: 'All libraries', url: '' });
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
      sizeErrorMessage: c.sizeErrorMessage,
      sizeApproximate: c.sizeApproximate,
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
  const unknownSizeFolders = React.useMemo(
    () => selectedChildren.filter((c) => c.sizeSource === 'error'),
    [selectedChildren],
  );
  const unknownSizeFolderCount = unknownSizeFolders.length;

  // The distinct underlying causes, not a guess. This banner used to assert
  // "likely SharePoint throttling" for every failure and tell people to lower
  // their concurrency — unhelpful when the real cause was a 403 on one folder
  // or a name the REST API rejects, neither of which concurrency affects.
  const unknownSizeReasons = React.useMemo(() => {
    const reasons = new Set<string>();
    for (const f of unknownSizeFolders) {
      const message = f.sizeErrorMessage ?? '';
      // 406 is checked with the throttling group, NOT as a bad request: SPO
      // redirects throttled calls to its HTML throttle page, which fails
      // content negotiation as 406. Classifying it as "unusual folder name"
      // (as this once did) blames the data and tells users the one setting
      // that helps won't help.
      if (/HTTP 429|HTTP 503|HTTP 406|throttl/i.test(message)) reasons.add('SharePoint throttling');
      else if (/HTTP 40[13]/.test(message)) reasons.add('permission denied on that folder');
      else if (/threshold/i.test(message)) reasons.add('the folder exceeds SharePoint\'s list view threshold');
      else if (message) reasons.add(message.split(' — ')[0]);
      else reasons.add('an unrecorded error');
    }
    return [...reasons];
  }, [unknownSizeFolders]);

  const listRows: FolderListRow[] = React.useMemo(() => [
    ...selectedChildren.map((c) => ({
      kind: 'folder' as const,
      name: c.name,
      serverRelativeUrl: c.serverRelativeUrl,
      sizeBytes: c.totalSizeBytes,
      lastModified: c.lastModified,
      itemCount: c.fileCount,
      sizeUnknown: c.sizeSource === 'error',
      sizeErrorMessage: c.sizeErrorMessage,
      sizeApproximate: c.sizeApproximate,
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
      versionCount: f.versionCount,
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
  // How many of the smallest items got folded into the Treemap's single
  // "Other" cell (see MAX_TREEMAP_CELLS in treemapLayout.ts) — used to show a
  // persistent note rather than relying on the user to notice and hover the
  // cell. The List view has no such folding, so it's always where they are.
  // ── Site-root level: one square/row per library ─────────────────────────
  // Libraries are modelled as 'folder' items so they reuse the whole existing
  // folder presentation (blue fill, click-to-drill, the striped "size unknown"
  // treatment) rather than needing a parallel set of cases everywhere.
  const rootTreemapItems: TreemapItem[] = React.useMemo(
    () => rollups.map((r) => ({
      id: r.library.serverRelativeUrl,
      label: r.library.title,
      sizeBytes: r.totalSizeBytes,
      kind: 'folder' as const,
      lastModified: r.lastModified,
      itemCount: r.fileCount,
      sizeUnknown: r.sizeUnknown,
      sizeApproximate: r.sizeApproximate,
    })),
    [rollups],
  );

  const rootListRows: FolderListRow[] = React.useMemo(
    () => rollups.map((r) => ({
      kind: 'folder' as const,
      name: r.library.title,
      serverRelativeUrl: r.library.serverRelativeUrl,
      sizeBytes: r.totalSizeBytes,
      lastModified: r.lastModified,
      itemCount: r.fileCount,
      sizeUnknown: r.sizeUnknown,
      sizeApproximate: r.sizeApproximate,
    })),
    [rollups],
  );

  const activeTreemapItems = atRoot ? rootTreemapItems : treemapItems;
  const activeListRows = atRoot ? rootListRows : listRows;
  const unknownRootCount = React.useMemo(() => rollups.filter((r) => r.sizeUnknown).length, [rollups]);
  // Mirrors layoutTreemap's own folding rule (smallest beyond
  // MAX_TREEMAP_CELLS - 1 get folded into one "Other" cell) — recomputing the
  // count here rather than reading it off the rendered rects keeps this note
  // independent of the Treemap's own width/height, which aren't known until
  // after its first paint.
  const foldedIntoOtherCount = Math.max(0, activeTreemapItems.length - (MAX_TREEMAP_CELLS - 1));

  // reduce, not Math.max(1, ...listRows.map(...)) — spreading tens of
  // thousands of arguments into Math.max risks a RangeError on very large
  // folders.
  const maxListSize = React.useMemo(
    () => activeListRows.reduce((max, r) => Math.max(max, r.sizeBytes), 1),
    [activeListRows],
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
          onClick={() => (atRoot ? setLibraryUrl(r.serverRelativeUrl) : setSelectedUrl(r.serverRelativeUrl))}
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
        <Tooltip content={r.sizeErrorMessage ?? 'Size could not be determined, and no error detail was recorded. Not a confirmed empty folder.'} relationship="label">
          <span style={{ display: 'inline-flex', alignItems: 'center', gap: '4px', color: tokens.colorPaletteMarigoldForeground1, cursor: 'help' }}>
            <Warning16Regular /> Unknown
          </span>
        </Tooltip>
      ) : (
        <div style={{ display: 'flex', alignItems: 'center', gap: '8px', justifyContent: 'flex-end' }}>
          {r.sizeApproximate ? (
            <Tooltip content="At least this much. Measurement was stopped before this folder's subtree was fully counted — open the folder to measure it directly, or use Refresh to measure again." relationship="label">
              <span style={{ minWidth: '64px', textAlign: 'right', cursor: 'help' }}>≥ {formatBytes(r.sizeBytes)}</span>
            </Tooltip>
          ) : (
            <span style={{ minWidth: '64px', textAlign: 'right' }}>{formatBytes(r.sizeBytes)}</span>
          )}
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
    }, {
      key: 'versionCount',
      header: 'Version Count (est.)',
      align: 'right' as const,
      sortValue: (r: FolderListRow) => r.versionCount ?? -1,
      render: (r: FolderListRow) => <span>{r.kind === 'file' && r.versionCount !== undefined ? r.versionCount : '—'}</span>,
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

  // Derived straight from the breadcrumb, skipping its "All libraries" root.
  // (It used to prefix currentLibrary.title separately, which now duplicates
  // it — the library is itself a breadcrumb segment since the root level.)
  const contextLabel = breadcrumb.slice(1).map((b) => b.name).join(' / ') || 'All libraries';
  // Excel export lazy-loads a chunk and writes a workbook, so unlike CSV it
  // has real ways to fail — left un-caught (as `void`-fired before), a
  // failure was an invisible unhandled promise rejection with no user-facing
  // signal that anything went wrong.
  const handleExportExcel = async (): Promise<void> => {
    try {
      setError('');
      // activeListRows/atRoot, not listRows unconditionally — the root
      // (library) level has its own row set, and exporting the folder-level
      // one there would silently export stale data from whatever folder was
      // last visited instead of the libraries actually on screen.
      await excel.exportFolderListing(activeListRows, contextLabel, !atRoot && includeVersions);
    } catch (err: any) {
      setError(`Excel export failed: ${err?.message ?? String(err)}`);
    }
  };
  const handleExportCsv = (): void => {
    excel.exportFolderListingCsv(activeListRows, contextLabel, !atRoot && includeVersions);
  };

  // Re-measure the current view from SharePoint. All three caches must go:
  // the sessionStorage folder-size cache (survives a page reload, 10-min TTL)
  // plus the two in-memory maps. Clearing only the persistent one would leave
  // loadChildren short-circuiting on its in-memory copy — which is precisely
  // why an "Unknown" folder was previously unretryable for the whole session.
  const handleRefresh = (): void => {
    sp.clearFolderSizeCache(siteUrl);
    childrenCache.current.clear();
    filesCache.current.clear();
    setLoadedFilesKey('');
    setError('');
    setCanceledKeys(new Set());
    setRollupsAbandoned(false);
    setCacheVersion((v) => v + 1);
    setRefreshToken((t) => t + 1);
  };

  // Whether what's in state actually corresponds to the folder now selected.
  // Checked in addition to the loading flags because both flags are set by
  // effects that run a render AFTER selectedUrl changes — in that gap nothing
  // is "loading" yet and the caches are empty, so the old gate rendered an
  // empty Treemap: a 420px featureless gray box (plus "no subfolders or
  // files"), which is exactly what showed up when clicking a large library.
  const childrenReady = React.useMemo(
    () => (selectedUrl ? childrenCache.current.has(selectedUrl) : false),
    [selectedUrl, cacheVersion],
  );
  const filesReady = loadedFilesKey === selectedUrl;
  const showLoading = atRoot
    ? rollupsLoading
    : !!selectedUrl && (childrenLoading || filesLoading || !childrenReady || !filesReady);

  // Drives the elapsed counter, the throttle notice, and the live
  // folders-scanned counter while loading. These exist to guarantee the
  // indicator always has visible motion: without them, a large library sits
  // on "0 of 240" for minutes (each increment is one child's whole subtree)
  // and looks indistinguishable from a hang. 500ms rather than 1s so the
  // folder count visibly moves.
  React.useEffect(() => {
    if (!showLoading) {
      setLoadElapsed(0);
      setThrottled(false);
      setItemsRead(0);
      return;
    }
    const startedAt = Date.now();
    const id = setInterval(() => {
      setLoadElapsed(Math.floor((Date.now() - startedAt) / 1000));
      setThrottled(sp.isThrottled);
      setItemsRead(itemsReadRef.current);
    }, 500);
    return () => clearInterval(id);
  }, [showLoading, sp]);

  // Always indeterminate now. The sweep pages through a library's items and
  // SharePoint does not report a total up front, so there is no honest
  // denominator to draw a determinate bar against — and a determinate bar
  // guessing at one is worse than an animated bar that admits it doesn't
  // know. The item counter below carries the real "something is happening"
  // signal.
  const loadingStatus = cancelRequested
    ? 'Stopping — finishing the page already in flight…'
    : throttled
      ? 'Paused — SharePoint is throttling requests. Waiting for it to allow more…'
      : atRoot
        ? 'Measuring library sizes…'
        : 'Reading library contents…';

  // Items, not folders: one flat sweep reads files and folders together, so
  // the honest unit of progress is items read.
  const visitedText = itemsRead > 0
    ? ` — ${itemsRead.toLocaleString()} item${itemsRead === 1 ? '' : 's'} read`
    : '';

  const loadingIndicator = (
    <div style={{ marginTop: tokens.spacingVerticalM }}>
      <ProgressBar />
      <Text
        style={{
          display: 'block',
          marginTop: tokens.spacingVerticalXS,
          fontSize: tokens.fontSizeBase200,
          color: throttled ? tokens.colorPaletteMarigoldForeground1 : tokens.colorNeutralForeground3,
        }}
      >
        {loadingStatus}{visitedText} — {loadElapsed}s elapsed
      </Text>
      {loadElapsed >= 20 && !throttled && !cancelRequested && (
        <Text style={{ display: 'block', fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3 }}>
          This library is read once, in full — after that, every folder inside it opens instantly.
          There is no time limit; click Cancel to stop now and keep what has been measured so far.
        </Text>
      )}
    </div>
  );

  return (
    <div className={styles.root}>
      {onBack && (
        <div style={{ marginBottom: tokens.spacingVerticalM }}>
          <Button appearance="subtle" icon={<ArrowLeft24Regular />} onClick={onBack} aria-label="Back to home">
            Home
          </Button>
        </div>
      )}
      {error && (
        <MessageBar intent="error" style={{ marginBottom: tokens.spacingVerticalM }}>
          <MessageBarBody>{error}</MessageBarBody>
        </MessageBar>
      )}

      {!showLoading && canceledKeys.has(atRoot ? '' : selectedUrl) && (
        <MessageBar intent="warning" style={{ marginBottom: tokens.spacingVerticalM }}>
          <MessageBarBody>
            Measurement canceled — showing what was measured before you stopped. Folders that weren't
            finished show "≥ size" (a real floor) or "Unknown" (never reached). Use Refresh to measure
            this view again.
          </MessageBarBody>
        </MessageBar>
      )}

      {!showLoading && !atRoot && unknownSizeFolderCount > 0 && (
        <MessageBar intent="warning" style={{ marginBottom: tokens.spacingVerticalM }}>
          <MessageBarBody>
            {unknownSizeFolderCount} folder{unknownSizeFolderCount === 1 ? '' : 's'} could not be measured, even after
            retrying, and show as "Unknown" rather than a real size — they are not confirmed empty.
            {' '}Cause: {unknownSizeReasons.join('; ')}.
            {unknownSizeReasons.some((r) => r.includes('throttl'))
              ? ' Lowering Concurrent API requests in Settings usually helps.'
              : ' This is specific to those folders — changing Concurrent API requests will not affect it.'}
            {' '}Hover a folder for its exact error, or check the browser console.
          </MessageBarBody>
        </MessageBar>
      )}

      {/* Reaching this now means the live walk failed for a library, which is
          unusual — a stale SharePoint rollup alone no longer lands here (see
          getLibraryRollups), it produces a "≥" floor instead. */}
      {!showLoading && atRoot && unknownRootCount > 0 && (
        <MessageBar intent="info" style={{ marginBottom: tokens.spacingVerticalM }}>
          <MessageBarBody>
            {unknownRootCount} librar{unknownRootCount === 1 ? 'y' : 'ies'} could not be measured and
            show as "Unknown" — most often SharePoint throttling. Try Refresh in a moment, or open the
            library to measure it directly. Sizes shown as "≥" are real floors, measured as far as the
            scan got.
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
            <Button
              appearance={atRoot ? 'primary' : 'secondary'}
              size="small"
              onClick={() => setLibraryUrl(undefined)}
            >
              Site
            </Button>
            {libraries.map((l) => (
              l.isRecycleBin ? (
                <Tooltip
                  key={l.serverRelativeUrl}
                  content="First-stage Recycle Bin — items an end user can still restore. Items purged to the site collection Recycle Bin require Site Collection Administrator access and aren't included."
                  relationship="label"
                >
                  <Button
                    appearance={l.serverRelativeUrl === libraryUrl ? 'primary' : 'secondary'}
                    size="small"
                    icon={<Delete16Regular />}
                    onClick={() => setLibraryUrl(l.serverRelativeUrl)}
                  >
                    {l.title}
                  </Button>
                </Tooltip>
              ) : (
                <Button
                  key={l.serverRelativeUrl}
                  appearance={l.serverRelativeUrl === libraryUrl ? 'primary' : 'secondary'}
                  size="small"
                  onClick={() => setLibraryUrl(l.serverRelativeUrl)}
                >
                  {l.title}
                </Button>
              )
            ))}
          </div>

          <div className={styles.breadcrumb}>
            {breadcrumb.map((b, i) => (
              <React.Fragment key={b.url}>
                {i > 0 && <ChevronRight16Regular style={{ color: tokens.colorNeutralForeground3 }} />}
                <Button
                  appearance="transparent"
                  size="small"
                  onClick={() => (b.url === '' ? setLibraryUrl(undefined) : setSelectedUrl(b.url))}
                >
                  {b.name}
                </Button>
              </React.Fragment>
            ))}
          </div>

          <TabList selectedValue={viewMode} onTabSelect={(_, d) => setViewMode(d.value as 'treemap' | 'list')} style={{ marginBottom: tokens.spacingVerticalM }}>
            <Tab value="treemap">Treemap</Tab>
            <Tab value="list">List</Tab>
          </TabList>

          <div style={{ marginBottom: tokens.spacingVerticalS, display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalS }}>
            {/* Version history is hidden at the library level: it's a per-file
                measurement and there are no files at this level. */}
            {!atRoot && (
              <div style={{ display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalXS }}>
                <Checkbox
                  label="Include version history size"
                  checked={includeVersions}
                  onChange={(_, d) => setIncludeVersions(!!d.checked)}
                />
                <Tooltip
                  content="Only individual files get a real version-history number — SharePoint's folder size rollup has no recursive version-history total, so a folder's Size never includes its files' version history, on or off. Version History Size is exact; Version Count is an estimate based on the file's current version number, so it can run slightly high (never low) on a library with a configured version-retention limit."
                  relationship="label"
                >
                  <Info16Regular style={{ cursor: 'help', color: tokens.colorNeutralForeground3 }} />
                </Tooltip>
              </div>
            )}
            <div style={{ display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalS, marginLeft: atRoot ? 0 : 'auto' }}>
              <Tooltip
                content="Discard cached sizes and measure this view again. Use this to retry folders showing Unknown or ≥, or after content has changed — sizes are otherwise cached for about 10 minutes."
                relationship="label"
              >
                <Button
                  icon={<ArrowClockwise20Regular />}
                  appearance="subtle"
                  size="small"
                  onClick={handleRefresh}
                  disabled={showLoading}
                >
                  Refresh
                </Button>
              </Tooltip>
              {showLoading && (
                <Tooltip
                  content="Stop measuring now and keep what has been measured so far. Folders that weren't finished show “≥ size” or “Unknown”."
                  relationship="label"
                >
                  <Button appearance="secondary" size="small" onClick={handleCancel} disabled={cancelRequested}>
                    {cancelRequested ? 'Canceling…' : 'Cancel'}
                  </Button>
                </Tooltip>
              )}
            </div>
          </div>

          {viewMode === 'treemap' ? (
            <div>
              <TierLegend staleDays={staleDays} veryStaleDays={veryStaleDays} />
              {!atRoot && includeVersions && (
                <Text style={{ display: 'block', color: tokens.colorNeutralForeground3, fontSize: tokens.fontSizeBase200, marginBottom: tokens.spacingVerticalS }}>
                  File squares are sized by file + version history combined ({formatBytes(versionsTotalBytes)} in version history across this folder's files); folders still show file size only — SharePoint has no recursive version-history rollup.
                </Text>
              )}
              {foldedIntoOtherCount > 0 && (
                <Text style={{ display: 'block', color: tokens.colorNeutralForeground3, fontSize: tokens.fontSizeBase200, marginBottom: tokens.spacingVerticalS }}>
                  The {foldedIntoOtherCount} smallest item{foldedIntoOtherCount === 1 ? '' : 's'} here {foldedIntoOtherCount === 1 ? 'is' : 'are'} combined into the "Other" square — it isn't a real folder or file, so there's nothing to click into.{' '}
                  <Button appearance="transparent" size="small" onClick={() => setViewMode('list')} style={{ padding: 0, minWidth: 'unset', height: 'auto', verticalAlign: 'baseline' }}>
                    Switch to List view
                  </Button>
                  {' '}to see them individually.
                </Text>
              )}
              {showLoading ? (
                loadingIndicator
              ) : (
                <>
                  <Treemap
                    items={activeTreemapItems}
                    onFolderClick={(item) => (atRoot ? setLibraryUrl(item.id) : setSelectedUrl(item.id))}
                  />
                  {activeTreemapItems.length === 0 && !error && (
                    <Body1 style={{ color: tokens.colorNeutralForeground3, marginTop: tokens.spacingVerticalM }}>
                      {atRoot ? 'No libraries with measurable content on this site.' : 'This folder has no subfolders or files.'}
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
                  <Button icon={<DocumentArrowDown24Regular />} appearance="secondary" size="small" onClick={handleExportExcel} disabled={activeListRows.length === 0}>
                    Export Excel
                  </Button>
                  <Button icon={<DocumentArrowDown24Regular />} appearance="secondary" size="small" onClick={handleExportCsv} disabled={activeListRows.length === 0}>
                    Export CSV
                  </Button>
                </div>
              </div>
              {!atRoot && includeVersions && (
                <Text style={{ display: 'block', color: tokens.colorNeutralForeground3, fontSize: tokens.fontSizeBase200, marginBottom: tokens.spacingVerticalS }}>
                  This folder's files: {formatBytes(versionsTotalBytes)} in version history (files only — folders show '—', no recursive rollup).
                </Text>
              )}
              {showLoading ? (
                loadingIndicator
              ) : activeListRows.length === 0 && !error ? (
                <Body1 style={{ color: tokens.colorNeutralForeground3 }}>
                  {atRoot ? 'No libraries found on this site.' : 'This folder is empty.'}
                </Body1>
              ) : (
                <StorageTable rows={activeListRows} columns={listColumns} getRowKey={(r) => r.serverRelativeUrl} defaultSortKey="size" />
              )}
            </div>
          )}
        </>
      )}
    </div>
  );
};
