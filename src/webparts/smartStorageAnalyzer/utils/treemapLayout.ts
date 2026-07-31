import { TreemapItem, TreemapRect } from '../models/models';

interface Box { x: number; y: number; width: number; height: number; }
interface SizedItem { id: string; value: number; }

// Cap on rendered cells per level — beyond this, the smallest items are
// folded into a single "Other" cell rather than rendering unreadable
// slivers (same idea as WizTree's small-file grouping).
export const MAX_TREEMAP_CELLS = 40;

// Smallest a cell is allowed to shrink to under pure area-proportional
// sizing, in CSS pixels. Below this a box becomes hard to hit with a mouse
// (and folder cells are the ones users need to click to drill in), so
// undersized items get bumped up to roughly this square footprint at the
// expense of a little precision in the largest cells' proportions.
const MIN_CELL_PX = 22;

function worst(row: SizedItem[], length: number): number {
  if (row.length === 0) return Infinity;
  const sum = row.reduce((s, r) => s + r.value, 0);
  const max = Math.max(...row.map((r) => r.value));
  const min = Math.min(...row.map((r) => r.value));
  const sqLen = length * length;
  return Math.max((sqLen * max) / (sum * sum), (sum * sum) / (sqLen * min));
}

// Classic squarified-treemap algorithm (Bruls/Huizing/van Wijk): grows a
// "row" of items along the shorter side of the remaining box for as long as
// doing so improves (or does not worsen) the worst aspect ratio in the row,
// then lays that row out and recurses on what's left.
function squarify(items: SizedItem[], box: Box): Map<string, Box> {
  const result = new Map<string, Box>();
  const total = items.reduce((s, i) => s + i.value, 0);
  if (total <= 0 || box.width <= 0 || box.height <= 0) return result;

  // Scale values into area units so their sum equals the box's pixel area.
  const scale = (box.width * box.height) / total;
  const scaled = items.map((i) => ({ id: i.id, value: i.value * scale }));

  function place(row: SizedItem[], b: Box): Box {
    // Lay the row out along the box's SHORTER side (fixed), subdividing the
    // longer side among the row's items — the part of the algorithm that
    // actually makes cells square-ish. A wide box (width >= height) gets a
    // narrow vertical column using the full height; a tall box gets a
    // horizontal band using the full width. Getting this backwards (using
    // the longer side as the fixed dimension) is what produced long thin
    // strips instead of near-square cells.
    const sum = row.reduce((s, r) => s + r.value, 0);
    if (b.width >= b.height) {
      const colWidth = b.height > 0 ? sum / b.height : 0;
      let cy = b.y;
      for (const r of row) {
        const h = colWidth > 0 ? r.value / colWidth : 0;
        result.set(r.id, { x: b.x, y: cy, width: colWidth, height: h });
        cy += h;
      }
      return { x: b.x + colWidth, y: b.y, width: Math.max(0, b.width - colWidth), height: b.height };
    }
    const rowHeight = b.width > 0 ? sum / b.width : 0;
    let cx = b.x;
    for (const r of row) {
      const w = rowHeight > 0 ? r.value / rowHeight : 0;
      result.set(r.id, { x: cx, y: b.y, width: w, height: rowHeight });
      cx += w;
    }
    return { x: b.x, y: b.y + rowHeight, width: b.width, height: Math.max(0, b.height - rowHeight) };
  }

  function layout(remaining: SizedItem[], row: SizedItem[], b: Box): void {
    if (remaining.length === 0) {
      if (row.length) place(row, b);
      return;
    }
    const length = Math.min(b.width, b.height);
    const next = remaining[0];
    const rest = remaining.slice(1);
    if (row.length === 0 || worst([...row, next], length) <= worst(row, length)) {
      layout(rest, [...row, next], b);
    } else {
      const newBox = place(row, b);
      layout(remaining, [], newBox);
    }
  }

  layout(scaled, [], box);
  return result;
}

// Public entry point: takes raw treemap items (already-sized folders/files
// for one level) plus a pixel box, and returns positioned rectangles ready
// to render. Items beyond MAX_TREEMAP_CELLS (by size, smallest first) are
// folded into a single synthetic "Other" item.
export function layoutTreemap(items: TreemapItem[], width: number, height: number): TreemapRect[] {
  if (width <= 0 || height <= 0) return [];
  // Previously items with sizeBytes <= 0 were dropped entirely — a folder
  // whose rollup resolves to 0 (empty, or its recursive size lookup failed
  // and fell back to an empty estimate) would silently vanish from the
  // treemap instead of appearing as a box. Keep every item and give
  // non-positive ones a visual floor below instead.
  const sorted = [...items].sort((a, b) => b.sizeBytes - a.sizeBytes);
  if (sorted.length === 0) return [];

  let visible = sorted;
  if (sorted.length > MAX_TREEMAP_CELLS) {
    const kept = sorted.slice(0, MAX_TREEMAP_CELLS - 1);
    const folded = sorted.slice(MAX_TREEMAP_CELLS - 1);
    const other: TreemapItem = {
      id: '__other__',
      label: `Other (${folded.length} items)`,
      sizeBytes: folded.reduce((s, i) => s + i.sizeBytes, 0),
      kind: 'other',
      count: folded.length,
    };
    visible = [...kept, other];
  }

  const byId = new Map(visible.map((i) => [i.id, i]));

  // Zero/negative-size items (e.g. an empty folder, or a library whose
  // rollup hasn't populated) get NO nominal share here — only real values
  // count toward the total the click-target floor is computed from. They
  // still end up visible: the MIN_CELL_PX floor below applies uniformly to
  // every item, zero or not, and is what actually guarantees clickability.
  //
  // This used to give zero-sized items 1% of the total instead, which is
  // fine for a folder-level view but breaks badly at library-root scale:
  // 1% of a few-hundred-GB site total is several GB — enough for a
  // genuinely-empty (or not-yet-measured) library's square to render
  // BIGGER than a library that actually holds 2-3GB. The fix is to size
  // zero items off the MIN_CELL_PX pixel target alone, never off a
  // fraction of everyone else's total.
  const positiveTotal = visible.reduce((s, i) => s + Math.max(0, i.sizeBytes), 0);
  const baseValues = visible.map((i) => Math.max(0, i.sizeBytes));
  // Falls back to 1 (not 0) only so minValue's division below never sees an
  // all-zero total — every item is 0 in that case anyway, so the exact
  // fallback value doesn't matter, only that it's positive.
  const baseTotal = positiveTotal || 1;

  // Convert the MIN_CELL_PX × MIN_CELL_PX target into a "value" floor in the
  // same units as sizeBytes, using the box's current scale (value → area).
  // squarify's own total will grow once values are bumped up, so the actual
  // rendered area ends up a bit under this target — close enough for a
  // click-target guarantee, not meant to be exact.
  const boxArea = width * height;
  const minValue = boxArea > 0 ? ((MIN_CELL_PX * MIN_CELL_PX) / boxArea) * baseTotal : 0;

  const boxes = squarify(
    visible.map((i, idx) => ({ id: i.id, value: Math.max(baseValues[idx], minValue, 1) })),
    { x: 0, y: 0, width, height },
  );

  const rects: TreemapRect[] = [];
  boxes.forEach((box, id) => {
    const item = byId.get(id);
    if (!item) return;
    rects.push({ ...item, ...box });
  });
  return rects;
}
