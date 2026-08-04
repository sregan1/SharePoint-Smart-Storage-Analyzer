const UNITS = ['B', 'KB', 'MB', 'GB', 'TB'];

export function formatBytes(bytes: number): string {
  if (!isFinite(bytes) || bytes <= 0) return '0 B';
  const exp = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), UNITS.length - 1);
  const value = bytes / Math.pow(1024, exp);
  return `${value.toFixed(exp === 0 ? 0 : 1)} ${UNITS[exp]}`;
}

export function formatAge(ageDays: number): string {
  if (ageDays < 30) return `${ageDays}d`;
  if (ageDays < 365) return `${Math.round(ageDays / 30)}mo`;
  return `${(ageDays / 365).toFixed(1)}y`;
}

// Coarse "time remaining" style formatting — one or two units, no seconds
// once minutes are in play, since an ETA precise to the second is false
// precision when it's derived from a throughput estimate that can swing with
// the next throttle response.
export function formatDuration(totalSeconds: number): string {
  const s = Math.max(0, Math.round(totalSeconds));
  if (s < 60) return `${s}s`;
  const minutes = Math.floor(s / 60);
  if (minutes < 60) return `${minutes}m`;
  const hours = Math.floor(minutes / 60);
  return `${hours}h ${minutes % 60}m`;
}

// "Time so far" formatting — unlike formatDuration (an estimate, where
// seconds are false precision once minutes are in play), this is a live
// stopwatch reading a real elapsed value, so minutes AND seconds stay
// meaningful together on a long-running scan.
export function formatElapsed(totalSeconds: number): string {
  const s = Math.max(0, Math.round(totalSeconds));
  if (s < 60) return `${s}s`;
  const minutes = Math.floor(s / 60);
  const seconds = s % 60;
  if (minutes < 60) return `${minutes}m ${seconds}s`;
  const hours = Math.floor(minutes / 60);
  return `${hours}h ${minutes % 60}m`;
}
