import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneDropdown } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import * as React from 'react';
import * as ReactDom from 'react-dom';

import { App, AppView, IBrandColors } from './components/App';
import { StorageAnalyzerService } from './services/StorageAnalyzerService';
import { ExcelExportService } from './services/ExcelExportService';

export interface ISmartStorageAnalyzerWebPartProps {
  defaultView: AppView;
}

// Views a user can be dropped onto directly — excludes 'settings', which is
// reachable only via the in-app gear icon, never a configured landing page.
const VALID_DEFAULT_VIEWS: AppView[] = ['home', 'tree', 'list', 'report'];

// this.properties.defaultView isn't just "unset or valid" — a web part placed
// before Home existed has 'explorer' persisted from the old AppView union, and
// any other unrecognized value could reach here the same way (a hand-edited
// page, a future removed view). App.tsx's view switch has no branch for those,
// so the content area silently renders nothing rather than erroring — the `??
// 'home'` fallback below only catches null/undefined, not a stale string. This
// normalizes both, so an invalid saved value is treated exactly like an unset
// one instead of producing a blank web part.
function resolveDefaultView(saved: AppView | undefined): AppView {
  return saved && VALID_DEFAULT_VIEWS.indexOf(saved) !== -1 ? saved : 'home';
}

export default class SmartStorageAnalyzerWebPart extends BaseClientSideWebPart<ISmartStorageAnalyzerWebPartProps> {
  private _sp: StorageAnalyzerService;
  private _excel: ExcelExportService;
  private _brandColors: IBrandColors = {
    primary: '#0078d4',
    darkAlt: '#106ebe',
    dark: '#005a9e',
    darker: '#004578',
    light: '#c7e0f4',
    lighter: '#deecf9',
  };
  // True for an inverted (dark) section or theme variant.
  private _isDark = false;
  private _ready = false;

  protected async onInit(): Promise<void> {
    await super.onInit();
    try {
      this._sp = new StorageAnalyzerService(this.context);
      this._excel = new ExcelExportService();
    } catch (err: any) {
      throw new Error(`[SmartStorageAnalyzer] Service init failed: ${err?.message ?? String(err)}\n${err?.stack ?? ''}`);
    }
    this._ready = true;
  }

  // SPFx calls this before onInit (supportsThemeVariants is on), and again on
  // every section/variant change. Store the theme and only repaint once
  // initialized — rendering earlier reads this.properties before it exists and
  // aborts the load with an opaque "[object Object]".
  protected onThemeChanged(theme: IReadonlyTheme | undefined): void {
    const p = theme?.palette;
    if (p?.themePrimary) {
      this._brandColors = {
        primary: p.themePrimary,
        darkAlt: p.themeDarkAlt ?? p.themePrimary,
        dark: p.themeDark ?? p.themePrimary,
        darker: p.themeDarker ?? p.themeDark ?? p.themePrimary,
        light: p.themeLight ?? '#c7e0f4',
        lighter: p.themeLighter ?? '#deecf9',
      };
    }
    this._isDark = !!theme?.isInverted;
    if (!this._ready) return;
    this.render();
  }

  public render(): void {
    if (!this._sp) return;
    try {
      const element = React.createElement(App, {
        context: this.context,
        sp: this._sp,
        excel: this._excel,
        defaultView: resolveDefaultView(this.properties.defaultView),
        brandColors: this._brandColors,
        isDark: this._isDark,
      });
      ReactDom.render(element, this.domElement);
    } catch (err: any) {
      const msg = err?.message ?? String(err);
      const stack = err?.stack ?? '(no stack)';
      this.domElement.innerHTML =
        `<div style="padding:16px;font-family:Consolas,monospace;font-size:13px;` +
        `background:#fff3f3;border:1px solid #c00;border-radius:4px;margin:8px">` +
        `<strong style="color:#c00;font-size:14px">Smart Storage Analyzer — Startup Error</strong><br><br>` +
        `<strong>Message:</strong> ${this._escHtml(msg)}<br><br>` +
        `<strong>Stack:</strong><pre style="font-size:11px;white-space:pre-wrap;` +
        `background:#f5f5f5;padding:8px;margin:4px 0;border-radius:2px">` +
        `${this._escHtml(stack)}</pre></div>`;
    }
  }

  private _escHtml(s: string): string {
    return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [{
        header: { description: 'Smart Storage Analyzer configuration' },
        groups: [{
          groupName: 'General',
          groupFields: [
            PropertyPaneDropdown('defaultView', {
              label: 'Default view on open',
              options: [
                { key: 'home', text: 'Home' },
                { key: 'tree', text: 'Tree View' },
                { key: 'list', text: 'List View' },
                { key: 'report', text: 'Storage Report' },
              ],
              selectedKey: resolveDefaultView(this.properties.defaultView),
            }),
          ],
        }],
      }],
    };
  }
}
