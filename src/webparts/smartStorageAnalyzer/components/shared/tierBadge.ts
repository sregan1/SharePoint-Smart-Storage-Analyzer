import { CandidateTier } from '../../models/models';

// Fixed status palette (never themed — same hex regardless of site brand
// color) shared between the file-list table badges, the Excel export fill
// colors, and the Treemap's file-cell colors, so the tier color language is
// identical everywhere it appears.
export const TIER_COLORS: Record<CandidateTier, { light: string; dark: string; label: string; description: string }> = {
  [CandidateTier.Active]: {
    light: '#0ca30c',
    dark: '#0ca30c',
    label: 'Active',
    description: 'Modified recently',
  },
  [CandidateTier.Stale]: {
    light: '#fab219',
    dark: '#fab219',
    label: 'Stale',
    description: 'Candidate for review',
  },
  [CandidateTier.VeryStale]: {
    light: '#d03b3b',
    dark: '#d03b3b',
    label: 'Very stale',
    description: 'Strong archival candidate',
  },
};

// Neutral categorical color for folders in the Treemap — containers, not a
// health signal, so they deliberately don't use the tier status colors.
export const FOLDER_COLOR = { light: '#2a78d6', dark: '#3987e5' };
export const OTHER_COLOR = { light: '#898781', dark: '#898781' };

export function tierColor(tier: CandidateTier): string {
  return TIER_COLORS[tier].light;
}

export function tierLabel(tier: CandidateTier): string {
  return TIER_COLORS[tier].label;
}
