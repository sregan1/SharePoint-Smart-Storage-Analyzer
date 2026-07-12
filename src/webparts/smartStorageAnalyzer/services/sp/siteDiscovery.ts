import { SpApiClient, isSystemLibrary, LIBRARY_TEMPLATES, valueArray } from './spCore';
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

// Document/picture/site-pages libraries on a single web (not recursive into
// subwebs — callers that support includeSubsites loop this per subweb URL).
export async function getLibraries(
  client: SpApiClient,
  siteUrl: string,
  includeHidden: boolean,
): Promise<LibraryInfo[]> {
  // Built from LIBRARY_TEMPLATES (spCore.ts) rather than hardcoding
  // "eq 101" so the walkable-template list can't drift from the constant
  // that everything else keys off — this used to only ever request 101
  // (Document Library), silently excluding Picture Libraries (109) and Site
  // Pages (119) from every scan/rollup despite both being able to hold
  // real storage.
  const templateFilter = LIBRARY_TEMPLATES.map((t) => `BaseTemplate eq ${t}`).join(' or ');
  const data = await client.getJson(
    `${siteUrl}/_api/web/lists?$select=Title,BaseTemplate,ItemCount,NoCrawl,RootFolder/ServerRelativeUrl,RootFolder/TimeLastModified,IsSiteAssetsLibrary,Hidden` +
      `&$expand=RootFolder&$filter=(${templateFilter})`,
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
