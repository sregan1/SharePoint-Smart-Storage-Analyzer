import * as React from 'react';
import {
  Button,
  Checkbox,
  Text,
  Title3,
  Label,
  Divider,
  Tooltip,
  SpinButton,
  tokens,
  makeStyles,
} from '@fluentui/react-components';
import { ArrowLeft24Regular, Info16Regular } from '@fluentui/react-icons';

const useStyles = makeStyles({
  root: {
    padding: tokens.spacingVerticalL,
    maxWidth: '540px',
    margin: '0 auto',
  },
  header: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    marginBottom: tokens.spacingVerticalL,
  },
  row: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
  },
  section: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
  },
  hint: {
    display: 'block',
    color: tokens.colorNeutralForeground3,
    marginLeft: '24px',
    lineHeight: '1.5',
  },
  instructionList: {
    margin: '8px 0 0 0',
    paddingLeft: '20px',
    lineHeight: '1.8',
  },
});

export interface SettingsViewProps {
  includeHidden: boolean;
  onIncludeHiddenChange: (val: boolean) => void;
  includeSubsites: boolean;
  onIncludeSubsitesChange: (val: boolean) => void;
  scanConcurrency: number;
  onScanConcurrencyChange: (val: number) => void;
  staleDays: number;
  onStaleDaysChange: (val: number) => void;
  veryStaleDays: number;
  onVeryStaleDaysChange: (val: number) => void;
  onBack: () => void;
}

export const SettingsView: React.FC<SettingsViewProps> = ({
  includeHidden,
  onIncludeHiddenChange,
  includeSubsites,
  onIncludeSubsitesChange,
  scanConcurrency,
  onScanConcurrencyChange,
  staleDays,
  onStaleDaysChange,
  veryStaleDays,
  onVeryStaleDaysChange,
  onBack,
}) => {
  const styles = useStyles();

  return (
    <div className={styles.root}>
      <div className={styles.header}>
        <Button appearance="subtle" icon={<ArrowLeft24Regular />} onClick={onBack}>
          Back
        </Button>
        <Title3>Settings</Title3>
      </div>

      <div style={{ display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalXL }}>

        {/* ── Scope ── */}
        <div className={styles.section}>
          <Text weight="semibold" style={{ display: 'block' }}>Scope</Text>
          <div className={styles.row}>
            <Checkbox
              label="Include subsites in Storage Report scans"
              checked={includeSubsites}
              onChange={(_, d) => onIncludeSubsitesChange(!!d.checked)}
            />
            <Tooltip
              content="When checked, the Storage Report also walks every subsite beneath this site. Does not affect Library Overview or Folder Explorer, which are always scoped to the current site."
              relationship="description"
              withArrow
            >
              <Button appearance="transparent" icon={<Info16Regular />} size="small" style={{ minWidth: 'unset', padding: '2px' }} aria-label="More info about subsites" />
            </Tooltip>
          </div>
          <div className={styles.row}>
            <Checkbox
              label="Include system and hidden libraries"
              checked={includeHidden}
              onChange={(_, d) => onIncludeHiddenChange(!!d.checked)}
            />
            <Tooltip
              content="When checked, includes Style Library, Form Templates, and other libraries normally hidden from default views. Applies to all tools."
              relationship="description"
              withArrow
            >
              <Button appearance="transparent" icon={<Info16Regular />} size="small" style={{ minWidth: 'unset', padding: '2px' }} aria-label="More info about hidden libraries" />
            </Tooltip>
          </div>
        </div>

        <Divider />

        {/* ── Archival thresholds ── */}
        <div className={styles.section}>
          <Text weight="semibold" style={{ display: 'block' }}>Archival thresholds</Text>
          <Text size={200} className={styles.hint} style={{ marginLeft: 0 }}>
            Based on last modified date only — SharePoint does not expose a reliable last-accessed
            signal at scale.
          </Text>

          <div className={styles.row}>
            <Label>Stale after (days):</Label>
            <SpinButton
              value={staleDays}
              min={1}
              max={veryStaleDays - 1}
              step={30}
              onChange={(_, d) => onStaleDaysChange(d.value !== undefined ? d.value : parseInt(d.displayValue ?? '180', 10))}
              style={{ width: '90px' }}
            />
          </div>
          <Text size={200} className={styles.hint}>
            Files not modified in this many days are flagged "Stale" — a candidate for review. Default: 180.
          </Text>

          <div className={styles.row} style={{ marginTop: tokens.spacingVerticalS }}>
            <Label>Very stale after (days):</Label>
            <SpinButton
              value={veryStaleDays}
              min={staleDays + 1}
              max={3650}
              step={30}
              onChange={(_, d) => onVeryStaleDaysChange(d.value !== undefined ? d.value : parseInt(d.displayValue ?? '365', 10))}
              style={{ width: '90px' }}
            />
          </div>
          <Text size={200} className={styles.hint}>
            Files not modified in this many days are flagged "Very stale" — a strong archival candidate. Default: 365.
          </Text>
        </div>

        <Divider />

        {/* ── Performance ── */}
        <div className={styles.section}>
          <Text weight="semibold" style={{ display: 'block' }}>Performance</Text>

          <div className={styles.row}>
            <Label>Concurrent API requests:</Label>
            <SpinButton
              value={scanConcurrency}
              min={1}
              max={15}
              onChange={(_, d) => onScanConcurrencyChange(d.value !== undefined ? d.value : parseInt(d.displayValue ?? '6', 10))}
              style={{ width: '80px' }}
            />
            <Tooltip
              content="How many SharePoint API requests run in parallel during scans and folder loads. SharePoint doesn't publish a fixed throttling limit — it's dynamic per tenant — so there's no universal 'right' number. If you push this up and start seeing things slow down instead of speed up, that's throttling kicking in; the app retries automatically, but turning this back down will help more than waiting it out."
              relationship="description"
              withArrow
            >
              <Button appearance="transparent" icon={<Info16Regular />} size="small" style={{ minWidth: 'unset', padding: '2px' }} aria-label="More info about concurrency" />
            </Tooltip>
          </div>
          <Text size={200} className={styles.hint}>
            Higher values scan faster but may trigger SharePoint throttling — the app retries automatically, but very high values can still net out slower. Default: 6.
          </Text>
        </div>

        <Divider />

        {/* ── Default view instructions ── */}
        <div className={styles.section}>
          <Text weight="semibold" style={{ display: 'block' }}>Default view on load</Text>
          <Text size={300} style={{ display: 'block', color: tokens.colorNeutralForeground2 }}>
            To change which screen opens when the web part first loads, edit the web part properties:
          </Text>
          <ol className={styles.instructionList}>
            {[
              <>Put the SharePoint page into <strong>Edit</strong> mode.</>,
              <>Click the <strong>pencil (edit)</strong> icon on the Smart Storage Analyzer web part.</>,
              <>In the property panel, choose a view from the <strong>Default view on open</strong> dropdown.</>,
              <><strong>Republish</strong> the page to save the change.</>,
            ].map((step, i) => (
              <li key={i}>
                <Text size={300} style={{ color: tokens.colorNeutralForeground2 }}>{step}</Text>
              </li>
            ))}
          </ol>
        </div>

      </div>
    </div>
  );
};
