import * as React from 'react';
import { Spinner, Text, Tooltip, tokens } from '@fluentui/react-components';
import { CandidateTier, TreemapItem, TreemapRect } from '../../models/models';
import { layoutTreemap } from '../../utils/treemapLayout';
import { FOLDER_COLOR, OTHER_COLOR, TIER_COLORS, tierLabel } from './tierBadge';
import { formatBytes, formatAge } from './formatBytes';

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
  if (item.kind === 'other') return `${item.label} — ${formatBytes(item.sizeBytes)}`;
  if (item.kind === 'folder') {
    if (item.sizeUnknown) {
      return `${item.label} (folder) — size unknown: SharePoint throttling or an error kept this from being measured. Not a confirmed empty folder — try again in a moment.`;
    }
    const fileCount = item.itemCount != null ? `, ${item.itemCount} file${item.itemCount === 1 ? '' : 's'}` : '';
    return `${item.label} (folder) — ${formatBytes(item.sizeBytes)}${fileCount}${item.lastModified ? `, most recent activity ${new Date(item.lastModified).toLocaleDateString()}` : ''}`;
  }
  const ageText = item.lastModified ? formatAge(Math.max(0, Math.floor((Date.now() - new Date(item.lastModified).getTime()) / 86400000))) : '';
  // item.sizeBytes is file + version history combined whenever versionSizeBytes
  // is present (see ExplorerView's treemapItems) — split it back out for display.
  const sizeText = item.versionSizeBytes != null
    ? `${formatBytes(item.sizeBytes)} (${formatBytes(item.sizeBytes - item.versionSizeBytes)} file + ${formatBytes(item.versionSizeBytes)} version history)`
    : formatBytes(item.sizeBytes);
  return `${item.label} — ${sizeText} — ${tierLabel(item.tier ?? CandidateTier.Active)}${ageText ? ` — modified ${ageText} ago` : ''}`;
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
          <Spinner size="small" label="Rendering treemap…" />
        </div>
      )}
      {rects.map((r) => {
        const clickable = r.kind === 'folder';
        const showLabel = r.width > 44 && r.height > 20;
        const showDetail = showLabel && r.height > 36 && (r.kind === 'folder' || r.versionSizeBytes != null);
        const cell = (
          <div
            key={r.id}
            onClick={clickable ? () => onFolderClick(r) : undefined}
            role={clickable ? 'button' : undefined}
            tabIndex={clickable ? 0 : undefined}
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
              background: cellColor(r),
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
                  color: 'white',
                  fontSize: tokens.fontSizeBase200,
                  textShadow: '0 1px 2px rgba(0,0,0,0.6)',
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
                  color: 'white',
                  fontSize: tokens.fontSizeBase100,
                  textShadow: '0 1px 2px rgba(0,0,0,0.6)',
                  overflow: 'hidden',
                  textOverflow: 'ellipsis',
                  whiteSpace: 'nowrap',
                  width: '100%',
                  opacity: 0.85,
                }}
              >
                {r.kind === 'folder'
                  ? (r.sizeUnknown ? 'Size unknown' : `${formatBytes(r.sizeBytes)}${r.itemCount != null ? ` · ${r.itemCount} file${r.itemCount === 1 ? '' : 's'}` : ''}`)
                  : `${formatBytes(r.sizeBytes - (r.versionSizeBytes ?? 0))} + ${formatBytes(r.versionSizeBytes ?? 0)} history`}
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
    ...(showFolder ? [{ color: FOLDER_COLOR.light, label: 'Folder' }] : []),
    { color: TIER_COLORS[CandidateTier.Active].light, label: `Active (< ${staleDays}d)` },
    { color: TIER_COLORS[CandidateTier.Stale].light, label: `Stale (${staleDays}–${veryStaleDays - 1}d)` },
    { color: TIER_COLORS[CandidateTier.VeryStale].light, label: `Very stale (${veryStaleDays}d+)` },
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
