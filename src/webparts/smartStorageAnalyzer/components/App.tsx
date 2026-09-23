import * as React from 'react';
import {
  FluentProvider,
  webLightTheme,
  webDarkTheme,
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
import { HomeView } from './HomeView';
import { ExplorerView } from './ExplorerView';
import { StorageReportView } from './StorageReportView';
import { SettingsView } from './SettingsView';
import { clampConcurrency, clampStaleDays, clampVeryStaleDays } from '../utils/settingsBounds';
import { safeGet, safeSet } from '../utils/safeStorage';
import { readableTextOn } from './shared/readableText';

export type AppView = 'home' | 'tree' | 'list' | 'report' | 'settings';

// Every value this switches on below; anything else (a stale saved value
// from a AppView union that no longer includes it, e.g. the pre-Home
// 'explorer') falls back to 'home' in resolveInitialView rather than
// matching no render branch and leaving the content area blank.
const KNOWN_VIEWS: AppView[] = ['home', 'tree', 'list', 'report', 'settings'];

function resolveInitialView(defaultView: AppView | undefined): AppView {
  return defaultView && KNOWN_VIEWS.indexOf(defaultView) !== -1 ? defaultView : 'home';
}

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

function buildTheme(b: IBrandColors, isDark: boolean): Theme {
  // On a dark background the brand's light shades are what read as "brand
  // text"; its pale tints (used for light-mode accent fills) would glare, so
  // the dark shades take their place.
  const fg = isDark ? b.light : b.primary;
  const fgHover = isDark ? b.lighter : b.darkAlt;
  const tint = isDark ? b.darker : b.lighter;
  const tintHover = isDark ? b.dark : b.light;
  return {
    ...(isDark ? webDarkTheme : webLightTheme),
    colorBrandBackground: b.primary,
    colorBrandBackgroundHover: b.darkAlt,
    colorBrandBackgroundPressed: b.dark,
    colorBrandBackgroundSelected: b.darkAlt,
    colorBrandBackgroundStatic: b.primary,
    colorBrandBackground2: tint,
    colorBrandBackground2Hover: tintHover,
    colorBrandBackground2Pressed: tintHover,
    colorBrandBackground3Static: b.dark,
    colorBrandBackground4Static: b.darker,
    colorCompoundBrandBackground: b.primary,
    colorCompoundBrandBackgroundHover: b.darkAlt,
    colorCompoundBrandBackgroundPressed: b.dark,
    colorBrandForeground1: fg,
    colorBrandForeground2: fgHover,
    colorBrandForeground2Hover: fg,
    colorBrandForeground2Pressed: fg,
    colorCompoundBrandForeground1: fg,
    colorCompoundBrandForeground1Hover: fgHover,
    colorCompoundBrandForeground1Pressed: fg,
    colorBrandForegroundLink: fg,
    colorBrandForegroundLinkHover: fgHover,
    colorBrandForegroundLinkPressed: fg,
    colorBrandForegroundLinkSelected: fg,
    colorBrandStroke1: b.primary,
    colorBrandStroke2: isDark ? b.dark : b.light,
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
  isDark?: boolean;
}

// tokens.* are "var(--name)" strings; this turns one into "var(--name, fallback)".
function withFallback(token: string, fallback: string): string {
  return token.replace(/\)$/, `, ${fallback})`);
}

interface ErrorBoundaryProps {
  children: React.ReactNode;
  // When set, the fallback offers a way out instead of leaving a dead screen.
  onReset?: () => void;
}

// Used twice: once around the whole app (last resort, outside the theme), and
// once per screen inside it, so a render error in one screen can be escaped by
// going back to Home rather than replacing the entire web part.
class ErrorBoundary extends React.Component<ErrorBoundaryProps, { error: Error | null }> {
  constructor(props: ErrorBoundaryProps) {
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
        <div
          role="alert"
          // Theme tokens with literal fallbacks: the outermost boundary sits
          // outside FluentProvider, where the token CSS variables don't exist.
          style={{
            padding: '16px', fontSize: '13px',
            background: withFallback(tokens.colorPaletteRedBackground1, '#fff3f3'),
            color: withFallback(tokens.colorNeutralForeground1, '#242424'),
            border: `1px solid ${withFallback(tokens.colorPaletteRedBorder2, '#c00')}`,
            borderRadius: '4px', margin: '8px',
          }}
        >
          <strong style={{ fontSize: '14px' }}>Smart Storage Analyzer — something went wrong</strong>
          <div style={{ marginTop: '8px' }}>{error.message || String(error)}</div>
          {this.props.onReset && (
            <Button
              appearance="primary"
              style={{ marginTop: '12px' }}
              onClick={() => { this.setState({ error: null }); this.props.onReset!(); }}
            >
              Back to Home
            </Button>
          )}
          <details style={{ marginTop: '12px' }}>
            <summary style={{ cursor: 'pointer' }}>Technical details</summary>
            <pre style={{
              fontFamily: 'Consolas, monospace', fontSize: '11px', whiteSpace: 'pre-wrap',
              overflowWrap: 'anywhere', padding: '8px', margin: '4px 0',
            }}>
              {error.stack ?? '(no stack available)'}
            </pre>
          </details>
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

export const App: React.FC<AppProps> = ({ context, sp, excel, defaultView, brandColors, isDark = false }) => {
  // Depend on every palette slot, not just primary — a theme variant switch
  // that keeps the same primary but changes the others (e.g. dark/light alt
  // shades) would otherwise render with a stale theme object.
  const theme = React.useMemo(
    () => buildTheme(brandColors, isDark),
    [brandColors.primary, brandColors.darkAlt, brandColors.dark, brandColors.darker, brandColors.light, brandColors.lighter, isDark],
  );
  const headerText = readableTextOn(brandColors.primary);

  const [view, setView] = React.useState<AppView>(resolveInitialView(defaultView));
  const [prevView, setPrevView] = React.useState<AppView>('home');

  // The web part always analyzes the site it's placed on — no cross-site
  // switching, so this is a plain constant rather than state.
  const siteUrl = context.pageContext.web.absoluteUrl;

  const [includeHidden, setIncludeHidden] = React.useState(
    () => safeGet(LS_HIDDEN) === 'true',
  );
  const [includeSubsites, setIncludeSubsites] = React.useState(
    () => safeGet(LS_SUBSITES) === 'true',
  );
  const [scanConcurrency, setScanConcurrency] = React.useState(
    () => clampConcurrency(safeGet(LS_CONCURRENCY)),
  );
  const [staleDays, setStaleDays] = React.useState(
    () => clampStaleDays(safeGet(LS_STALE_DAYS)),
  );
  const [veryStaleDays, setVeryStaleDays] = React.useState(
    () => clampVeryStaleDays(safeGet(LS_VERY_STALE_DAYS), clampStaleDays(safeGet(LS_STALE_DAYS))),
  );

  React.useEffect(() => { safeSet(LS_HIDDEN, String(includeHidden)); }, [includeHidden]);
  React.useEffect(() => { safeSet(LS_SUBSITES, String(includeSubsites)); }, [includeSubsites]);
  React.useEffect(() => {
    safeSet(LS_CONCURRENCY, String(scanConcurrency));
    sp.scanConcurrency = scanConcurrency;
  }, [scanConcurrency]);
  React.useEffect(() => { safeSet(LS_STALE_DAYS, String(staleDays)); }, [staleDays]);
  React.useEffect(() => { safeSet(LS_VERY_STALE_DAYS, String(veryStaleDays)); }, [veryStaleDays]);

  // Re-checked whenever staleDays changes (not just at load) because
  // veryStaleDays is a separately-stored value that may not move in the same
  // interaction — without this, raising "Stale after" past the current
  // "Very stale after" would leave an invalid combination in state even
  // though each value is individually within its own bounds.
  React.useEffect(() => {
    setVeryStaleDays((v) => clampVeryStaleDays(v, staleDays));
  }, [staleDays]);

  const [canManageWeb, setCanManageWeb] = React.useState<boolean | null>(null);

  React.useEffect(() => {
    setCanManageWeb(null);
    let canceled = false;
    sp.checkCanManageWeb(siteUrl).then((can: boolean) => {
      if (!canceled) setCanManageWeb(can);
    }).catch(() => { if (!canceled) setCanManageWeb(true); });
    return () => { canceled = true; };
  }, [siteUrl]);

  const handleOpenSettings = (): void => {
    setPrevView(view === 'settings' ? prevView : view);
    setView('settings');
  };

  return (
    <ErrorBoundary>
    <RendererProvider renderer={renderer} targetDocument={document}>
    <FluentProvider theme={theme} style={{ minHeight: '400px', position: 'relative' }}>

      {view !== 'home' && view !== 'settings' && (
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
            <HardDrive24Regular style={{ color: headerText, fontSize: '20px' }} />
            <Text style={{ color: headerText, fontWeight: tokens.fontWeightSemibold, whiteSpace: 'nowrap' }}>
              SharePoint Smart Storage Analyzer
            </Text>
          </div>

          <Button
            appearance="transparent"
            icon={<Settings24Regular style={{ color: headerText }} />}
            aria-label="Settings"
            title="Settings"
            onClick={handleOpenSettings}
          />
        </div>
      )}

      {view === 'home' && (
        <div style={{ position: 'absolute', top: '4px', right: '8px', zIndex: 10 }}>
          <Button
            appearance="transparent"
            icon={<Settings24Regular />}
            aria-label="Settings"
            title="Settings"
            onClick={handleOpenSettings}
          />
        </div>
      )}

      {view !== 'home' && view !== 'settings' && canManageWeb === false && (
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

      <ErrorBoundary key={view} onReset={() => setView('home')}>
      {view === 'home' && (
        <HomeView
          onNavigate={setView}
          primaryColor={brandColors.primary}
          canManageWeb={canManageWeb}
        />
      )}
      {(view === 'tree' || view === 'list') && (
        <ExplorerView
          key={`${siteUrl}-${view}`}
          sp={sp}
          excel={excel}
          siteUrl={siteUrl}
          staleDays={staleDays}
          veryStaleDays={veryStaleDays}
          initialViewMode={view === 'list' ? 'list' : 'treemap'}
          onBack={() => setView('home')}
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
          onBack={() => setView('home')}
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
      </ErrorBoundary>

    </FluentProvider>
    </RendererProvider>
    </ErrorBoundary>
  );
};
