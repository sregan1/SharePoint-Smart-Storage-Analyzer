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
import { clampConcurrency, clampStaleDays, clampVeryStaleDays } from '../utils/settingsBounds';
import * as strings from 'SmartStorageAnalyzerWebPartStrings';

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
          {strings.Settings_Back}
        </Button>
        <Title3>{strings.Settings_Title}</Title3>
      </div>

      <div style={{ display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalXL }}>

        {/* ── Scope ── */}
        <div className={styles.section}>
          <Text weight="semibold" style={{ display: 'block' }}>{strings.Settings_Section_Scope}</Text>
          <div className={styles.row}>
            <Checkbox
              label={strings.Settings_IncludeSubsites_Label}
              checked={includeSubsites}
              onChange={(_, d) => onIncludeSubsitesChange(!!d.checked)}
            />
            <Tooltip
              content={strings.Settings_IncludeSubsites_Tooltip}
              relationship="description"
              withArrow
            >
              <Button appearance="transparent" icon={<Info16Regular />} size="small" style={{ minWidth: 'unset', padding: '2px' }} aria-label={strings.Settings_IncludeSubsites_InfoAria} />
            </Tooltip>
          </div>
          <div className={styles.row}>
            <Checkbox
              label={strings.Settings_IncludeHidden_Label}
              checked={includeHidden}
              onChange={(_, d) => onIncludeHiddenChange(!!d.checked)}
            />
            <Tooltip
              content={strings.Settings_IncludeHidden_Tooltip}
              relationship="description"
              withArrow
            >
              <Button appearance="transparent" icon={<Info16Regular />} size="small" style={{ minWidth: 'unset', padding: '2px' }} aria-label={strings.Settings_IncludeHidden_InfoAria} />
            </Tooltip>
          </div>
        </div>

        <Divider />

        {/* ── Archival thresholds ── */}
        <div className={styles.section}>
          <Text weight="semibold" style={{ display: 'block' }}>{strings.Settings_Section_Thresholds}</Text>
          <Text size={200} className={styles.hint} style={{ marginLeft: 0 }}>
            {strings.Settings_Thresholds_Hint}
          </Text>

          <div className={styles.row}>
            <Label>{strings.Settings_StaleAfter_Label}</Label>
            <SpinButton
              value={staleDays}
              min={1}
              max={veryStaleDays - 1}
              step={30}
              onChange={(_, d) => onStaleDaysChange(clampStaleDays(d.value ?? d.displayValue))}
              style={{ width: '90px' }}
            />
          </div>
          <Text size={200} className={styles.hint}>
            {strings.Settings_StaleAfter_Hint}
          </Text>

          <div className={styles.row} style={{ marginTop: tokens.spacingVerticalS }}>
            <Label>{strings.Settings_VeryStaleAfter_Label}</Label>
            <SpinButton
              value={veryStaleDays}
              min={staleDays + 1}
              max={3650}
              step={30}
              onChange={(_, d) => onVeryStaleDaysChange(clampVeryStaleDays(d.value ?? d.displayValue, staleDays))}
              style={{ width: '90px' }}
            />
          </div>
          <Text size={200} className={styles.hint}>
            {strings.Settings_VeryStaleAfter_Hint}
          </Text>
        </div>

        <Divider />

        {/* ── Performance ── */}
        <div className={styles.section}>
          <Text weight="semibold" style={{ display: 'block' }}>{strings.Settings_Section_Performance}</Text>

          <div className={styles.row}>
            <Label>{strings.Settings_Concurrency_Label}</Label>
            <SpinButton
              value={scanConcurrency}
              min={1}
              max={15}
              onChange={(_, d) => onScanConcurrencyChange(clampConcurrency(d.value ?? d.displayValue))}
              style={{ width: '80px' }}
            />
            <Tooltip
              content={strings.Settings_Concurrency_Tooltip}
              relationship="description"
              withArrow
            >
              <Button appearance="transparent" icon={<Info16Regular />} size="small" style={{ minWidth: 'unset', padding: '2px' }} aria-label={strings.Settings_Concurrency_InfoAria} />
            </Tooltip>
          </div>
          <Text size={200} className={styles.hint}>
            {strings.Settings_Concurrency_Hint}
          </Text>
        </div>

        <Divider />

        {/* ── Default view instructions ── */}
        <div className={styles.section}>
          <Text weight="semibold" style={{ display: 'block' }}>{strings.Settings_Section_DefaultView}</Text>
          <Text size={300} style={{ display: 'block', color: tokens.colorNeutralForeground2 }}>
            {strings.Settings_DefaultView_Intro}
          </Text>
          <ol className={styles.instructionList}>
            {[
              <>{strings.Settings_Step1_Pre}<strong>{strings.Settings_Step1_Bold}</strong>{strings.Settings_Step1_Post}</>,
              <>{strings.Settings_Step2_Pre}<strong>{strings.Settings_Step2_Bold}</strong>{strings.Settings_Step2_Post}</>,
              <>{strings.Settings_Step3_Pre}<strong>{strings.Settings_Step3_Bold}</strong>{strings.Settings_Step3_Post}</>,
              <><strong>{strings.Settings_Step4_Bold}</strong>{strings.Settings_Step4_Post}</>,
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
