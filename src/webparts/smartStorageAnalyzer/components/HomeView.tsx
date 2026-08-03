import * as React from 'react';
import {
  Button,
  Card,
  Text,
  Body1,
  Tooltip,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import {
  HardDrive24Regular,
  DataTreemap24Regular,
  AppsList24Regular,
  DataBarVertical24Regular,
  LockClosed16Regular,
} from '@fluentui/react-icons';
import { AppView } from './App';
import screenshotTree from '../assets/screenshot_tree.png';
import screenshotList from '../assets/screenshot_list.png';
import screenshotReport from '../assets/screenshot_report.png';

const useStyles = makeStyles({
  banner: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    paddingTop: tokens.spacingVerticalS,
    paddingBottom: tokens.spacingVerticalS,
    paddingLeft: tokens.spacingHorizontalM,
    paddingRight: tokens.spacingHorizontalM,
  },
  root: {
    padding: tokens.spacingVerticalXL,
    maxWidth: '1100px',
    margin: '0 auto',
  },
  subtitle: {
    marginBottom: tokens.spacingVerticalXL,
  },
  grid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(3, 1fr)',
    gap: tokens.spacingHorizontalL,
    '@media (max-width: 800px)': {
      gridTemplateColumns: 'repeat(2, 1fr)',
    },
    '@media (max-width: 500px)': {
      gridTemplateColumns: '1fr',
    },
  },
  card: {
    cursor: 'pointer',
    display: 'flex',
    flexDirection: 'column',
    overflow: 'hidden',
    padding: '0',
    ':hover': {
      boxShadow: tokens.shadow16,
    },
  },
  cardImage: {
    width: '100%',
    // Matches the screenshots' own 900x479 aspect ratio (~1.88:1) rather than
    // a fixed pixel height, so the full mockup renders with no cropping
    // regardless of the card's actual rendered width at each grid breakpoint.
    aspectRatio: '900 / 479',
    flexShrink: 0,
    objectFit: 'cover',
    objectPosition: 'center top',
    display: 'block',
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  cardBody: {
    display: 'flex',
    flexDirection: 'column',
    flexGrow: 1,
    padding: tokens.spacingVerticalM,
    gap: tokens.spacingVerticalS,
  },
  cardTitleRow: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
  },
  cardTitle: {
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightSemibold,
  },
  cardDesc: {
    flexGrow: 1,
    color: tokens.colorNeutralForeground2,
    fontSize: tokens.fontSizeBase200,
  },
  cardFooter: {
    padding: tokens.spacingVerticalM,
  },
  navButton: {
    width: '100%',
    minHeight: '36px',
  },
  cardDisabled: {
    opacity: '0.15',
    pointerEvents: 'none',
  },
});

export interface HomeViewProps {
  onNavigate: (view: AppView) => void;
  primaryColor: string;
  canManageWeb: boolean | null;
}

export const HomeView: React.FC<HomeViewProps> = ({ onNavigate, primaryColor, canManageWeb }) => {
  const styles = useStyles();

  const cards = [
    {
      view: 'tree' as AppView,
      icon: <DataTreemap24Regular style={{ flexShrink: 0 }} />,
      title: 'Tree View',
      screenshot: screenshotTree,
      alt: 'Tree View treemap of libraries and folders sized by storage',
      desc: 'See every library and folder as a proportional treemap, sized by storage used, and drill straight into the biggest ones.',
      buttonLabel: 'Open Tree View',
    },
    {
      view: 'list' as AppView,
      icon: <AppsList24Regular style={{ flexShrink: 0 }} />,
      title: 'List View',
      screenshot: screenshotList,
      alt: 'List View sortable table of folders and files',
      desc: 'Browse folders and files in a sortable table — size, file count, last modified — for a more precise, spreadsheet-like look.',
      buttonLabel: 'Open List View',
    },
    {
      view: 'report' as AppView,
      icon: <DataBarVertical24Regular style={{ flexShrink: 0 }} />,
      title: 'Storage Report',
      screenshot: screenshotReport,
      alt: 'Storage Report configuration and export screen',
      desc: 'Run a full scan of the site (optionally including subsites) and export a detailed Excel report of stale and very stale files.',
      buttonLabel: 'Run Storage Report',
    },
  ];

  return (
    <div>
      <div className={styles.banner} style={{ background: primaryColor }}>
        <HardDrive24Regular style={{ color: 'white', fontSize: '20px', flexShrink: 0 }} />
        <Text style={{ color: 'white', fontWeight: tokens.fontWeightSemibold, whiteSpace: 'nowrap' }}>
          SharePoint Smart Storage Analyzer
        </Text>
      </div>

      <div className={styles.root}>
        {canManageWeb === false && (
          <div style={{
            display: 'flex',
            alignItems: 'flex-start',
            gap: tokens.spacingHorizontalS,
            padding: tokens.spacingVerticalM,
            marginBottom: tokens.spacingVerticalL,
            background: tokens.colorPaletteYellowBackground1,
            border: `1px solid ${tokens.colorPaletteYellowBorder1}`,
            borderRadius: tokens.borderRadiusMedium,
            color: tokens.colorNeutralForeground1,
            fontSize: tokens.fontSizeBase300,
            lineHeight: tokens.lineHeightBase300,
          }}>
            <span style={{ flexShrink: 0, fontSize: '16px' }}>⚠️</span>
            <span>
              <strong>Site Owner access required — </strong>
              Storage totals rely on the same right as classic Site Settings → Storage Metrics. Contact a site owner if you have questions.
            </span>
          </div>
        )}

        <div className={styles.subtitle}>
          <Body1 style={{ color: tokens.colorNeutralForeground3 }}>
            Understand where your SharePoint storage is going — no PowerShell required.
          </Body1>
        </div>

        <div className={styles.grid}>
          {cards.map(({ view, icon, title, screenshot, alt, desc, buttonLabel }) => {
            const disabled = canManageWeb === false;
            const card = (
              <Card
                key={view}
                className={`${styles.card}${disabled ? ` ${styles.cardDisabled}` : ''}`}
                style={disabled ? { filter: 'grayscale(1)' } : undefined}
                onClick={disabled ? undefined : () => onNavigate(view)}
                role="button"
                tabIndex={disabled ? -1 : 0}
                aria-disabled={disabled}
                aria-label={buttonLabel}
                onKeyDown={disabled ? undefined : (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); onNavigate(view); } }}
              >
                <div className={styles.cardBody}>
                  <div className={styles.cardTitleRow}>
                    {icon}
                    <Text className={styles.cardTitle}>{title}</Text>
                  </div>
                  <Body1 className={styles.cardDesc}>{desc}</Body1>
                </div>
                <img src={screenshot} alt={alt} className={styles.cardImage} />
                <div className={styles.cardFooter}>
                  <Button
                    appearance="primary"
                    className={styles.navButton}
                    tabIndex={-1}
                    disabled={disabled}
                    icon={disabled ? <LockClosed16Regular /> : undefined}
                    iconPosition="after"
                  >
                    {buttonLabel}
                  </Button>
                </div>
              </Card>
            );
            return disabled ? (
              <Tooltip key={view} content="Requires Site Owner access" relationship="description">
                <div style={{ cursor: 'not-allowed' }}>{card}</div>
              </Tooltip>
            ) : card;
          })}
        </div>

        <div
          style={{
            marginTop: tokens.spacingVerticalXXL,
            padding: tokens.spacingVerticalM,
            background: tokens.colorNeutralBackground3,
            borderRadius: tokens.borderRadiusMedium,
          }}
        >
          <Body1 style={{ color: tokens.colorNeutralForeground3 }}>
            <strong>Note:</strong> This web part runs as the currently signed-in
            user. It can only see sites and libraries that user has permission
            to view.
          </Body1>
        </div>
      </div>
    </div>
  );
};
