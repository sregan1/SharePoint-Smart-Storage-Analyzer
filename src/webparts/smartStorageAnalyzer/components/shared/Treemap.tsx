import * as React from 'react';
import { Spinner, Text, Tooltip, tokens } from '@fluentui/react-components';
import { CandidateTier, TreemapItem, TreemapRect } from '../../models/models';
import { layoutTreemap } from '../../utils/treemapLayout';
import { FOLDER_COLOR, OTHER_COLOR, TIER_COLORS, tierLabel } from './tierBadge';
import { formatBytes, formatAge } from './formatBytes';
import { readableTextOn } from './readableText';
import * as strings from 'SmartStorageAnalyzerWebPartStrings';
import { formatString } from '../localeUtils';

// Diagonal stripes (rather than a solid fill) for a folder whose size
// couldn't be determined — visually distinct at a glance from a genuinely
// small/empty folder, which would otherwise be an identically-colored,
// merely small rectangle.
const UNKNOWN_SIZE_PATTERN = 'repeating-linear-gradient(45deg, #a36200, #a36200 6px, #c77f00 6px, #c77f00 12px)';

function cellColor(item: TreemapRect): string {
  if (item.kind === 'folder') return item.sizeUnknown ? UNKNOWN_SIZE_PATTERN : FOLDER_COLOR.light;
  if (item.kind === 'other') return OTHER_COLOR.light;
  return TIER_COLORS[item.tier ?? CandidateTier.Active].light;
}

function cellTooltip(item: TreemapRect): string {
  // "Other" is a synthetic bucket folding the smallest items beyond
  // MAX_TREEMAP_CELLS into one cell (see treemapLayout.ts) — it isn't a real
  // folder or file, so there's nothing to drill into here. The List view has
  // no such folding (every row is listed individually), so that's where the
  // folded-in items actually are.
  if (item.kind === 'other') {
    // Broken down the same way an individual file's size is — otherwise the
    // folded total silently absorbs version-history bytes with no way to
    // tell it's in there, which is easy to miss on a site where most of the
    // version history sits in many small files rather than a few big ones.
    const sizeText = item.versionSizeBytes != null
      ? formatString(strings.Treemap_Other_SizeCombinedWithVersion, {
        total: formatBytes(item.sizeBytes),
        current: formatBytes(item.sizeBytes - item.versionSizeBytes),
        version: formatBytes(item.versionSizeBytes),
      })
      : formatString(strings.Treemap_Other_SizeCombined, { total: formatBytes(item.sizeBytes) });
    return formatString(strings.Treemap_Other_Tooltip, { label: item.label, sizeText });
  }
  if (item.kind === 'folder') {
    if (item.sizeUnknown) {
      // The real error when we have one — guessing at the cause here (it used
      // to always blame throttling) sends people to fix the wrong thing.
      return item.sizeErrorMessage
        ? formatString(strings.Treemap_Folder_UnknownWithError, { label: item.label, error: item.sizeErrorMessage })
        : formatString(strings.Treemap_Folder_UnknownNoError, { label: item.label });
    }
    const fileCount = item.itemCount != null
      ? formatString(item.itemCount === 1 ? strings.Treemap_Folder_FileCount_Singular : strings.Treemap_Folder_FileCount_Plural, { count: item.itemCount })
      : '';
    const size = item.sizeApproximate
      ? formatString(strings.Treemap_Folder_SizeApproximate, { bytes: formatBytes(item.sizeBytes) })
      : formatBytes(item.sizeBytes);
    // A folder never carries a version-history number — SharePoint has no
    // recursive rollup for one, and only individual files get a real,
    // measured figure (see the file branch below).
    const lastModified = item.lastModified
      ? formatString(strings.Treemap_Folder_LastModifiedSuffix, { date: new Date(item.lastModified).toLocaleDateString() })
      : '';
    return formatString(strings.Treemap_Folder_Tooltip, { label: item.label, size, fileCount, lastModified });
  }
  const ageText = item.lastModified ? formatAge(Math.max(0, Math.floor((Date.now() - new Date(item.lastModified).getTime()) / 86400000))) : '';
  // item.sizeBytes is file + version history combined whenever versionSizeBytes
  // is present (see ExplorerView's treemapItems) — split it back out for display.
  // versionSizeBytes undefined does NOT mean "confirmed zero" — SharePoint's
  // own storage-metrics field is populated by a lagging background job and
  // can simply be missing for a given file's row even on a library where
  // most files have it (a recently added/changed file is the common case).
  // Silently showing plain size made that indistinguishable from a file
  // that genuinely has no retained versions, so this says so explicitly,
  // matching the List view's "≥" treatment of the same gap.
  const sizeText = item.versionSizeBytes != null
    ? formatString(strings.Treemap_File_SizeWithVersion, {
      total: formatBytes(item.sizeBytes),
      current: formatBytes(item.sizeBytes - item.versionSizeBytes),
      version: formatBytes(item.versionSizeBytes),
    })
    : formatString(strings.Treemap_File_SizeUnmeasuredVersion, { bytes: formatBytes(item.sizeBytes) });
  const ageSuffix = ageText ? formatString(strings.Treemap_File_ModifiedSuffix, { age: ageText }) : '';
  return formatString(strings.Treemap_File_Tooltip, {
    label: item.label,
    sizeText,
    tier: tierLabel(item.tier ?? CandidateTier.Active),
    ageText: ageSuffix,
  });
}

export interface TreemapProps {
  items: TreemapItem[];
  height?: number;
  onFolderClick: (item: TreemapItem) => void;
}

// Hand-rolled WizTree-style squarified treemap: no charting dependency, just
// the pure layout from treemapLayout.ts rendered as absolutely-positioned
// divs (simpler text/tooltip handling than SVG for this use case).
export const Treemap: React.FC<TreemapProps> = ({ items, height = 420, onFolderClick }) => {
  const containerRef = React.useRef<HTMLDivElement>(null);
  const [width, setWidth] = React.useState(0);

  React.useEffect(() => {
    if (!containerRef.current) return;
    const el = containerRef.current;
    const observer = new ResizeObserver((entries) => {
      const w = entries[0]?.contentRect.width;
      if (w) setWidth(w);
    });
    observer.observe(el);
    setWidth(el.clientWidth);
    return () => observer.disconnect();
  }, []);

  const rects = React.useMemo(() => layoutTreemap(items, width, height), [items, width, height]);
  // The container's width is only known after its first ResizeObserver/
  // clientWidth measurement — until then layoutTreemap has nothing to size
  // against and returns no rects, which otherwise renders as a blank box
  // indistinguishable from "this folder is empty" for that one frame.
  const measuring = width <= 0 && items.length > 0;

  return (
    <div ref={containerRef} style={{ position: 'relative', width: '100%', height: `${height}px`, background: tokens.colorNeutralBackground2 }}>
      {measuring && (
        <div style={{ position: 'absolute', inset: 0, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
          <Spinner size="small" label={strings.Treemap_Rendering} />
        </div>
      )}
      {rects.map((r) => {
        const clickable = r.kind === 'folder';
        const showLabel = r.width > 44 && r.height > 20;
        const showDetail = showLabel && r.height > 36;
        const background = cellColor(r);
        // Black or white per cell: white text on the amber Stale color was
        // ~1.8:1. The striped unknown-size pattern is dark, so white.
        const textColor = r.kind === 'folder' && r.sizeUnknown ? '#ffffff' : readableTextOn(background);
        const textShadow = textColor === '#ffffff' ? '0 1px 2px rgba(0,0,0,0.6)' : 'none';
        const tierText = r.kind === 'file' ? tierLabel(r.tier ?? CandidateTier.Active) : '';
        const cell = (
          <div
            key={r.id}
            onClick={clickable ? () => onFolderClick(r) : undefined}
            // Every cell is focusable so a file's size, tier and age (otherwise
            // only in the hover tooltip) are reachable by keyboard; the Tooltip
            // supplies the accessible name.
            role={clickable ? 'button' : 'img'}
            tabIndex={0}
            onKeyDown={clickable ? (e) => {
              if (e.key === 'Enter' || e.key === ' ') {
                e.preventDefault();
                onFolderClick(r);
              }
            } : undefined}
            style={{
              position: 'absolute',
              left: `${r.x + 1}px`,
              top: `${r.y + 1}px`,
              width: `${Math.max(0, r.width - 2)}px`,
              height: `${Math.max(0, r.height - 2)}px`,
              background,
              border: `1px solid ${tokens.colorNeutralBackground2}`,
              boxSizing: 'border-box',
              overflow: 'hidden',
              cursor: clickable ? 'zoom-in' : 'default',
              display: 'flex',
              flexDirection: 'column',
              alignItems: 'flex-start',
              padding: '2px 4px',
            }}
          >
            {showLabel && (
              <Text
                style={{
                  color: textColor,
                  fontSize: tokens.fontSizeBase200,
                  textShadow,
                  overflow: 'hidden',
                  textOverflow: 'ellipsis',
                  whiteSpace: 'nowrap',
                  width: '100%',
                }}
              >
                {r.label}
              </Text>
            )}
            {showDetail && (
              <Text
                style={{
                  color: textColor,
                  fontSize: tokens.fontSizeBase100,
                  textShadow,
                  overflow: 'hidden',
                  textOverflow: 'ellipsis',
                  whiteSpace: 'nowrap',
                  width: '100%',
                  opacity: 0.9,
                }}
              >
                {r.kind === 'folder'
                  ? (r.sizeUnknown
                    ? strings.Treemap_Detail_SizeUnknown
                    : `${r.sizeApproximate ? '≥ ' : ''}${formatBytes(r.sizeBytes)}${r.itemCount != null ? formatString(r.itemCount === 1 ? strings.Treemap_Detail_FileCount_Singular : strings.Treemap_Detail_FileCount_Plural, { count: r.itemCount }) : ''}`)
                  : r.kind === 'other'
                    ? formatString(strings.Treemap_Other_SizeCombined, { total: formatBytes(r.sizeBytes) })
                    // The tier as text, not only as the cell's color. "≥"
                    // when version history isn't measured for this file (see
                    // cellTooltip) — not the same as a confirmed 0 B.
                    : r.versionSizeBytes != null
                      ? formatString(strings.Treemap_Detail_FileWithVersion, { tier: tierText, current: formatBytes(r.sizeBytes - r.versionSizeBytes), version: formatBytes(r.versionSizeBytes) })
                      : formatString(strings.Treemap_Detail_FileWithoutVersion, { tier: tierText, bytes: formatBytes(r.sizeBytes) })}
              </Text>
            )}
          </div>
        );
        return (
          <Tooltip key={r.id} content={cellTooltip(r)} relationship="label">
            {cell}
          </Tooltip>
        );
      })}
    </div>
  );
};

export interface TierLegendProps {
  staleDays: number;
  veryStaleDays: number;
  showFolder?: boolean;
}

// Shared key for the Active/Stale/Very stale color coding, used by the
// Treemap, the Explorer List view, and the Storage Report. Spells out the
// actual configured day thresholds (staleDays/veryStaleDays are user-editable
// in Settings) rather than just naming the tiers, since "Stale" alone
// doesn't say what makes a file stale. showFolder is off for the Storage
// Report, which lists files only — no folder rollups to explain.
export const TierLegend: React.FC<TierLegendProps> = ({ staleDays, veryStaleDays, showFolder = true }) => {
  const items: { color: string; label: string }[] = [
    ...(showFolder ? [{ color: FOLDER_COLOR.light, label: strings.Treemap_Legend_Folder }] : []),
    { color: TIER_COLORS[CandidateTier.Active].light, label: formatString(strings.Treemap_Legend_Active, { days: staleDays }) },
    { color: TIER_COLORS[CandidateTier.Stale].light, label: formatString(strings.Treemap_Legend_Stale, { from: staleDays, to: veryStaleDays - 1 }) },
    { color: TIER_COLORS[CandidateTier.VeryStale].light, label: formatString(strings.Treemap_Legend_VeryStale, { days: veryStaleDays }) },
  ];
  return (
    <div style={{ display: 'flex', gap: tokens.spacingHorizontalM, flexWrap: 'wrap', padding: `${tokens.spacingVerticalXS} 0` }}>
      {items.map((l) => (
        <div key={l.label} style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
          <span style={{ width: '10px', height: '10px', background: l.color, borderRadius: '2px', flexShrink: 0 }} />
          <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3 }}>{l.label}</Text>
        </div>
      ))}
    </div>
  );
};
