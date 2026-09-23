// Picks black or white text for a solid background color by WCAG relative
// luminance, so text stays readable on any site brand color or tier color
// (white on the amber "Stale" color is ~1.8:1, well under the 4.5:1 minimum).

function channel(c: number): number {
  const s = c / 255;
  return s <= 0.03928 ? s / 12.92 : Math.pow((s + 0.055) / 1.055, 2.4);
}

function luminance(hex: string): number | null {
  const m = /^#?([0-9a-f]{3}|[0-9a-f]{6})$/i.exec(hex.trim());
  if (!m) return null;
  let h = m[1];
  if (h.length === 3) h = h.split('').map((x) => x + x).join('');
  const r = parseInt(h.slice(0, 2), 16);
  const g = parseInt(h.slice(2, 4), 16);
  const b = parseInt(h.slice(4, 6), 16);
  return 0.2126 * channel(r) + 0.7152 * channel(g) + 0.0722 * channel(b);
}

export function readableTextOn(background: string): '#ffffff' | '#000000' {
  const l = luminance(background);
  if (l === null) return '#ffffff';
  // Contrast against white vs. against black; choose the higher.
  const vsWhite = 1.05 / (l + 0.05);
  const vsBlack = (l + 0.05) / 0.05;
  return vsWhite >= vsBlack ? '#ffffff' : '#000000';
}
