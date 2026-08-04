// Shared clamps for user-tunable settings. Applied at every layer they pass
// through (localStorage read, Settings UI input, SpApiClient setter) because
// each layer can independently receive an out-of-range value: a hand-edited
// localStorage entry, a typed SpinButton value (Fluent's min/max props only
// constrain the stepper buttons, not keyboard input), or a stale combination
// of staleDays/veryStaleDays left over from an older version's defaults.
// An invalid concurrency in particular is not just wrong but fatal: TaskQueue
// (spCore.ts) never runs a task when `active < concurrency` is never true,
// so NaN/0/negative silently hangs every scan and folder load forever.

export const MIN_CONCURRENCY = 1;
export const MAX_CONCURRENCY = 15;
export const DEFAULT_CONCURRENCY = 6;

export const MIN_STALE_DAYS = 1;
export const MAX_STALE_DAYS = 3649;
export const DEFAULT_STALE_DAYS = 180;

export const MIN_VERY_STALE_DAYS = 2;
export const MAX_VERY_STALE_DAYS = 3650;
export const DEFAULT_VERY_STALE_DAYS = 365;

// How many files ONE library may measure individually in a Quick version-history
// scan, when no bulk mechanism for version size works on that list and the only
// remaining option is one request per file (see versionSizes.ts).
//
// Not currently user-tunable, deliberately: the meaningful choice is Quick vs
// Full (a scan option in the UI), and a third numeric knob in Settings would be
// one nobody could reason about. Exported here rather than buried in
// versionSizes.ts so the UI can name the actual number in its own hint text
// instead of hardcoding a second copy of it.
//
// 5,000 is chosen against a real 193,915-item library: that is ~2.5% of the
// files, and because the budget is spent LARGEST-FIRST it still captures the
// large majority of retained-version bytes (version size correlates strongly
// with current file size). A cap that measured an arbitrary 5,000 would be
// nearly worthless by comparison.
export const QUICK_VERSION_FILE_LIMIT = 5000;

function clampInt(v: unknown, min: number, max: number, fallback: number): number {
  const n = typeof v === 'number' ? v : parseInt(String(v), 10);
  if (!isFinite(n)) return fallback;
  return Math.min(max, Math.max(min, Math.round(n)));
}

export function clampConcurrency(v: unknown): number {
  return clampInt(v, MIN_CONCURRENCY, MAX_CONCURRENCY, DEFAULT_CONCURRENCY);
}

export function clampStaleDays(v: unknown): number {
  return clampInt(v, MIN_STALE_DAYS, MAX_STALE_DAYS, DEFAULT_STALE_DAYS);
}

// veryStaleDays must stay strictly greater than staleDays — the SpinButton's
// min prop only enforces this against click-stepping, not typed/localStorage
// values, so it's re-checked here.
export function clampVeryStaleDays(v: unknown, staleDays: number): number {
  const min = Math.max(MIN_VERY_STALE_DAYS, staleDays + 1);
  const fallback = Math.max(min, DEFAULT_VERY_STALE_DAYS);
  return clampInt(v, min, MAX_VERY_STALE_DAYS, fallback);
}
