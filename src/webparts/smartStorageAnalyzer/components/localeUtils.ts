import type { WebPartContext } from '@microsoft/sp-webpart-base';

export function formatString(template: string, params: { [key: string]: string | number }): string {
  return template.replace(/\{(\w+)\}/g, (match, key) => {
    const value = params[key];
    return value === undefined ? match : String(value);
  });
}

export function getEffectiveLocale(context: WebPartContext | undefined): string {
  return context?.pageContext?.cultureInfo?.currentUICultureName || 'en-US';
}
