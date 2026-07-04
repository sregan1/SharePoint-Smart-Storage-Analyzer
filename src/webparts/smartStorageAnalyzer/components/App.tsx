import * as React from 'react';
import {
  FluentProvider,
  webLightTheme,
  createDOMRenderer,
  RendererProvider,
  Button,
  Text,
  tokens,
  Theme,
} from '@fluentui/react-components';
import { Settings24Regular, HardDrive24Regular } from '@fluentui/react-icons';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import { StorageAnalyzerService } from '../services/StorageAnalyzerService';
import { ExcelExportService } from '../services/ExcelExportService';
import { ExplorerView } from './ExplorerView';
import { StorageReportView } from './StorageReportView';
import { SettingsView } from './SettingsView';
import { DEFAULT_STALE_DAYS, DEFAULT_VERY_STALE_DAYS } from '../utils/archivalClassification';

export type AppView = 'explorer' | 'report' | 'settings';

const LS_CONCURRENCY = 'sp-smart-storage-analyzer-concurrency';
const LS_HIDDEN = 'sp-smart-storage-analyzer-includeHidden';
const LS_SUBSITES = 'sp-smart-storage-analyzer-includeSubsites';
const LS_STALE_DAYS = 'sp-smart-storage-analyzer-staleDays';
const LS_VERY_STALE_DAYS = 'sp-smart-storage-analyzer-veryStaleDays';

export interface IBrandColors {
  primary: string;
  darkAlt: string;
  dark: string;
  darker: string;
  light: string;
  lighter: string;
}

function buildTheme(b: IBrandColors): Theme {
  return {
    ...webLightTheme,
    colorBrandBackground: b.primary,
    colorBrandBackgroundHover: b.darkAlt,
    colorBrandBackgroundPressed: b.dark,
    colorBrandBackgroundSelected: b.darkAlt,
    colorBrandBackgroundStatic: b.primary,
    colorBrandBackground2: b.lighter,
    colorBrandBackground2Hover: b.light,
    colorBrandBackground2Pressed: b.light,
    colorBrandBackground3Static: b.dark,
    colorBrandBackground4Static: b.darker,
    colorCompoundBrandBackground: b.primary,
    colorCompoundBrandBackgroundHover: b.darkAlt,
    colorCompoundBrandBackgroundPressed: b.dark,
    colorBrandForeground1: b.primary,
    colorBrandForeground2: b.darkAlt,
    colorBrandForeground2Hover: b.dark,
    colorBrandForeground2Pressed: b.darker,
    colorCompoundBrandForeground1: b.primary,
    colorCompoundBrandForeground1Hover: b.darkAlt,
    colorCompoundBrandForeground1Pressed: b.dark,
    colorBrandForegroundLink: b.primary,
    colorBrandForegroundLinkHover: b.darkAlt,
    colorBrandForegroundLinkPressed: b.dark,
    colorBrandForegroundLinkSelected: b.primary,
    colorBrandStroke1: b.primary,
    colorBrandStroke2: b.light,
    colorBrandStroke2Hover: b.primary,
    colorBrandStroke2Pressed: b.darkAlt,
    colorCompoundBrandStroke: b.primary,
    colorCompoundBrandStrokeHover: b.darkAlt,
    colorCompoundBrandStrokePressed: b.dark,
  };
}

export interface AppProps {
  context: WebPartContext;
  sp: StorageAnalyzerService;
  excel: ExcelExportService;
  defaultView?: AppView;
  brandColors: IBrandColors;
}

class ErrorBoundary extends React.Component<
  { children: React.ReactNode },
  { error: Error | null }
> {
  constructor(props: { children: React.ReactNode }) {
    super(props);
    this.state = { error: null };
  }

  static getDerivedStateFromError(error: Error): { error: Error } {
    return { error };
  }

  componentDidCatch(error: Error, info: React.ErrorInfo): void {
    console.error('[SmartStorageAnalyzer] Render error:', error, info.componentStack);
  }

  render(): React.ReactNode {
    const { error } = this.state;
    if (error) {
      return (
        <div style={{
          padding: '16px', fontFamily: 'Consolas, monospace', fontSize: '13px',
          background: '#fff3f3', border: '1px solid #c00', borderRadius: '4px', margin: '8px',
        }}>
          <strong style={{ color: '#c00', fontSize: '14px' }}>Smart Storage Analyzer — Render Error</strong>
          <br /><br />
          <strong>Message:</strong> {error.message || String(error)}
          <br /><br />
          <strong>Stack:</strong>
          <pre style={{
            fontSize: '11px', whiteSpace: 'pre-wrap', overflowWrap: 'anywhere',
            background: '#f5f5f5', padding: '8px', margin: '4px 0', borderRadius: '2px',
          }}>
            {error.stack ?? '(no stack available)'}
          </pre>
        </div>
      );
    }
    return this.props.children;
  }
}

let renderer: ReturnType<typeof createDOMRenderer>;
try {
  renderer = createDOMRenderer(document);
} catch (e: any) {
  console.error('[SmartStorageAnalyzer] createDOMRenderer failed:', e);
  throw e;
}

export const App: React.FC<AppProps> = ({ context, sp, excel, defaultView, brandColors }) => {
  const theme = React.useMemo(() => buildTheme(brandColors), [brandColors.primary]);

  const [view, setView] = React.useState<AppView>(defaultView ?? 'explorer');
  const [prevView, setPrevView] = React.useState<AppView>('explorer');

  // The web part always analyzes the site it's placed on — no cross-site
  // switching, so this is a plain constant rather than state.
  const siteUrl = context.pageContext.web.absoluteUrl;

  const [includeHidden, setIncludeHidden] = React.useState(
    () => localStorage.getItem(LS_HIDDEN) === 'true',
  );
  const [includeSubsites, setIncludeSubsites] = React.useState(
    () => localStorage.getItem(LS_SUBSITES) === 'true',
  );
  const [scanConcurrency, setScanConcurrency] = React.useState(
    () => parseInt(localStorage.getItem(LS_CONCURRENCY) ?? '6', 10),
  );
  const [staleDays, setStaleDays] = React.useState(
    () => parseInt(localStorage.getItem(LS_STALE_DAYS) ?? String(DEFAULT_STALE_DAYS), 10),
  );
  const [veryStaleDays, setVeryStaleDays] = React.useState(
    () => parseInt(localStorage.getItem(LS_VERY_STALE_DAYS) ?? String(DEFAULT_VERY_STALE_DAYS), 10),
  );

  React.useEffect(() => { localStorage.setItem(LS_HIDDEN, String(includeHidden)); }, [includeHidden]);
  React.useEffect(() => { localStorage.setItem(LS_SUBSITES, String(includeSubsites)); }, [includeSubsites]);
  React.useEffect(() => {
    localStorage.setItem(LS_CONCURRENCY, String(scanConcurrency));
    sp.scanConcurrency = scanConcurrency;
  }, [scanConcurrency]);
  React.useEffect(() => { localStorage.setItem(LS_STALE_DAYS, String(staleDays)); }, [staleDays]);
  React.useEffect(() => { localStorage.setItem(LS_VERY_STALE_DAYS, String(veryStaleDays)); }, [veryStaleDays]);

  const [canManageWeb, setCanManageWeb] = React.useState<boolean | null>(null);

  React.useEffect(() => {
    setCanManageWeb(null);
    let cancelled = false;
    sp.checkCanManageWeb(siteUrl).then((can: boolean) => {
      if (!cancelled) setCanManageWeb(can);
    }).catch(() => { if (!cancelled) setCanManageWeb(true); });
    return () => { cancelled = true; };
  }, [siteUrl]);

  const handleOpenSettings = (): void => {
    setPrevView(view === 'settings' ? prevView : view);
    setView('settings');
  };

  return (
    <ErrorBoundary>
    <RendererProvider renderer={renderer} targetDocument={document}>
    <FluentProvider theme={theme} style={{ minHeight: '400px', position: 'relative' }}>

      {view !== 'settings' && (
        <div
          style={{
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between',
            paddingTop: tokens.spacingVerticalS,
            paddingBottom: tokens.spacingVerticalS,
            paddingLeft: tokens.spacingHorizontalM,
            paddingRight: tokens.spacingHorizontalS,
            background: brandColors.primary,
            gap: tokens.spacingHorizontalM,
          }}
        >
          <div style={{ display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalS, flexShrink: 0 }}>
            <HardDrive24Regular style={{ color: 'white', fontSize: '20px' }} />
            <Text style={{ color: 'white', fontWeight: tokens.fontWeightSemibold, whiteSpace: 'nowrap' }}>
              SharePoint Smart Storage Analyzer
            </Text>
          </div>

          <div style={{ display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalS, flexShrink: 0 }}>
            <Button
              appearance={view === 'explorer' ? 'primary' : 'transparent'}
              size="small"
              style={view === 'explorer' ? undefined : { color: 'white' }}
              onClick={() => setView('explorer')}
            >
              Explorer
            </Button>
            <Button
              appearance={view === 'report' ? 'primary' : 'transparent'}
              size="small"
              style={view === 'report' ? undefined : { color: 'white' }}
              onClick={() => setView('report')}
            >
              Storage Report
            </Button>
            <Button
              appearance="transparent"
              icon={<Settings24Regular style={{ color: 'white' }} />}
              aria-label="Settings"
              title="Settings"
              onClick={handleOpenSettings}
            />
          </div>
        </div>
      )}

      {view !== 'settings' && canManageWeb === false && (
        <div style={{
          display: 'flex',
          alignItems: 'flex-start',
          gap: tokens.spacingHorizontalS,
          padding: tokens.spacingVerticalM,
          margin: `${tokens.spacingVerticalM} ${tokens.spacingHorizontalM}`,
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

      {view === 'explorer' && (
        <ExplorerView
          key={siteUrl}
          sp={sp}
          excel={excel}
          siteUrl={siteUrl}
          staleDays={staleDays}
          veryStaleDays={veryStaleDays}
        />
      )}
      {view === 'report' && (
        <StorageReportView
          key={siteUrl}
          sp={sp}
          excel={excel}
          siteUrl={siteUrl}
          includeSubsites={includeSubsites}
          includeHidden={includeHidden}
          staleDays={staleDays}
          veryStaleDays={veryStaleDays}
          onBack={() => setView('explorer')}
        />
      )}
      {view === 'settings' && (
        <SettingsView
          includeHidden={includeHidden}
          onIncludeHiddenChange={setIncludeHidden}
          includeSubsites={includeSubsites}
          onIncludeSubsitesChange={setIncludeSubsites}
          scanConcurrency={scanConcurrency}
          onScanConcurrencyChange={setScanConcurrency}
          staleDays={staleDays}
          onStaleDaysChange={setStaleDays}
          veryStaleDays={veryStaleDays}
          onVeryStaleDaysChange={setVeryStaleDays}
          onBack={() => setView(prevView)}
        />
      )}

    </FluentProvider>
    </RendererProvider>
    </ErrorBoundary>
  );
};
