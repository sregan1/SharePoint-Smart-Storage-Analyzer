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
