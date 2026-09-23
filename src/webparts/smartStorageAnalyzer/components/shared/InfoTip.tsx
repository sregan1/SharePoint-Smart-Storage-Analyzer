import * as React from 'react';
import { Button, Tooltip, TooltipProps } from '@fluentui/react-components';
import { Info16Regular } from '@fluentui/react-icons';

export interface InfoTipProps {
  content: TooltipProps['content'];
  // Accessible name for the icon button, e.g. "About Current File Size".
  label: string;
}

// An info icon whose tooltip can be reached by keyboard and read by screen
// readers — a bare icon inside a Tooltip is neither focusable nor announced.
// Same pattern SettingsView already uses.
export const InfoTip: React.FC<InfoTipProps> = ({ content, label }) => (
  <Tooltip content={content} relationship="description" withArrow>
    <Button
      appearance="transparent"
      size="small"
      icon={<Info16Regular />}
      aria-label={label}
      style={{ minWidth: 'unset', padding: '0 2px', height: 'auto', verticalAlign: 'middle' }}
    />
  </Tooltip>
);
