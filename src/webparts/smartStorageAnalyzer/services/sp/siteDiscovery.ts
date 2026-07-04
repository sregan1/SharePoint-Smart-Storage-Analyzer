import { SpApiClient, isSystemLibrary, valueArray } from './spCore';
import { LibraryInfo } from '../../models/models';

// GETs _api/web?$select=EffectiveBasePermissions and checks whether the
// current user has ManageWeb (0x40000000) or ManagePermissions (0x02000000)
// on the Low bitmask value — the same right required to see classic Site
// Settings → Storage Metrics. Fails OPEN (returns true) on API error so a
// transient failure never wrongly hides the tool from a real owner.
export async function checkCanManageWeb(client: SpApiClient, siteUrl: string): Promise<boolean> {
  try {
    const data = await client.getJson(`${siteUrl}/_api/web?$select=EffectiveBasePermissions`);
    const low = parseInt(data?.EffectiveBasePermissions?.Low ?? data?.d?.EffectiveBasePermissions?.Low ?? '0', 10) >>> 0;
    return (low & 0x40000000) !== 0 || (low & 0x02000000) !== 0;
  } catch {
    return true;
  }
}

export interface SiteOwnerInfo {
  displayName: string;
  email?: string;
}

export async function getSiteOwners(client: SpApiClient, siteUrl: string): Promise<SiteOwnerInfo[]> {
  try {
    const data = await client.getJson(
      `${siteUrl}/_api/web/AssociatedOwnerGroup/users?$select=Title,Email`,
    );
    return valueArray(data).map((u: any) => ({ displayName: u.Title, email: u.Email || undefined }));
  } catch {
    return [];
  }
}

// Document/picture/site-pages libraries on a single web (not recursive into
// subwebs — callers that support includeSubsites loop this per subweb URL).
export async function getLibraries(
  client: SpApiClient,
  siteUrl: string,
  includeHidden: boolean,
): Promise<LibraryInfo[]> {
  const data = await client.getJson(
    `${siteUrl}/_api/web/lists?$select=Title,BaseTemplate,ItemCount,NoCrawl,RootFolder/ServerRelativeUrl,RootFolder/TimeLastModified,IsSiteAssetsLibrary,Hidden` +
      `&$expand=RootFolder&$filter=BaseTemplate eq 101`,
  );
  const lists = valueArray(data);
  return lists
    .filter((lib: any) => !lib.Hidden)
    .filter((lib: any) => includeHidden || !isSystemLibrary(lib))
    .map((lib: any) => ({
      title: lib.Title as string,
      serverRelativeUrl: lib.RootFolder?.ServerRelativeUrl as string,
      itemCount: lib.ItemCount as number,
      noCrawl: !!lib.NoCrawl,
      baseTemplate: lib.BaseTemplate as number,
      lastModified: lib.RootFolder?.TimeLastModified as string | undefined,
    }));
}

export interface SubwebInfo {
  title: string;
  url: string;
}

// Immediate + recursive subwebs below a site, via the classic webs REST
// endpoint. Silently skips a branch on 403 (no access to that subweb) rather
// than failing the whole scan — matches how SharePoint itself hides subsites
// the current user cannot see.
export async function getSubwebsRecursive(client: SpApiClient, siteUrl: string): Promise<SubwebInfo[]> {
  const result: SubwebInfo[] = [];
  async function walk(webUrl: string): Promise<void> {
    let children: any[];
    try {
      const data = await client.getJson(`${webUrl}/_api/web/webs?$select=Title,Url&$top=500`);
      children = valueArray(data);
    } catch {
      return; // no access — skip silently
    }
    for (const child of children) {
      result.push({ title: child.Title, url: child.Url });
      await walk(child.Url);
    }
  }
  await walk(siteUrl);
  return result;
}
