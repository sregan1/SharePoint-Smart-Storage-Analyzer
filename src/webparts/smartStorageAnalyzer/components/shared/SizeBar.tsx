import * as React from 'react';
import { tokens } from '@fluentui/react-components';

export interface SizeBarProps {
  value: number;
  max: number;
  color?: string;
}

// Small horizontal fill bar showing a value's share of a max — used next to
// library/folder rows so relative storage weight is visible at a glance
// without needing a separate chart.
export const SizeBar: React.FC<SizeBarProps> = ({ value, max, color }) => {
  const pct = max > 0 ? Math.min(100, (value / max) * 100) : 0;
  return (
    <div
      style={{
        width: '100%',
        height: '6px',
        background: tokens.colorNeutralBackground3,
        borderRadius: '3px',
        overflow: 'hidden',
      }}
      role="img"
      aria-label={`${pct.toFixed(0)}% of largest`}
    >
      <div
        style={{
          width: `${pct}%`,
          height: '100%',
          background: color ?? tokens.colorBrandBackground,
          borderRadius: '3px',
        }}
      />
    </div>
  );
};
