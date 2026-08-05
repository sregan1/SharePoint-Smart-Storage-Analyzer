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

// There was a QUICK_VERSION_FILE_LIMIT here, capping how many files one library
// would measure individually when no bulk mechanism for version size worked.
// Removed after measurement rather than on taste: on a real 184,859-file
// library only 8,168 files are candidates at all (version labels prove the rest
// have no retained versions), the capped run and the complete run agreed to
// 0.06%, and the complete run was fast enough not to need a ceiling. See
// versionSizes.ts.

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
