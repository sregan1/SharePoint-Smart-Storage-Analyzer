import { CandidateTier } from '../models/models';

export const DEFAULT_STALE_DAYS = 180;
export const DEFAULT_VERY_STALE_DAYS = 365;

// Age in whole days from a modified timestamp to now. Deliberately based on
// TimeLastModified only — SharePoint does not expose a reliable
// last-accessed signal at scale via REST/Graph without per-item calls and
// extra admin-consented permissions.
export function ageInDays(timeLastModified: string, now: Date = new Date()): number {
  const modified = new Date(timeLastModified).getTime();
  if (isNaN(modified)) return 0;
  return Math.max(0, Math.floor((now.getTime() - modified) / 86400000));
}

export function classify(ageDays: number, staleDays: number, veryStaleDays: number): CandidateTier {
  if (ageDays >= veryStaleDays) return CandidateTier.VeryStale;
  if (ageDays >= staleDays) return CandidateTier.Stale;
  return CandidateTier.Active;
}
