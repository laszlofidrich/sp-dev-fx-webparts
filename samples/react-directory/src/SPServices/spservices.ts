import { WebPartContext } from "@microsoft/sp-webpart-base";
import { sp } from '@pnp/sp';
import { SearchResults, ISearchQuery, SortDirection } from '@pnp/sp/search';
import { ISPServices } from "./ISPServices";
import { MSGraphClientV3 } from '@microsoft/sp-http';

/* =========================
   Szűrési segédek (People Search)
   ========================= */

// Névben CONTAINS tiltó tokenek
const EXCLUDED_TOKENS = ['(X)', '(SZ)', 'ADM', 'guest', 'UPS', 'Teszt', 'Test'];

// Csak ezek a domainek legyenek láthatók
const ALLOWED_EMAIL_DOMAINS = ['value4real.com','ofysol.com']; // bővíthető: ['value4real.com','contoso.com']

const shouldHideUserFromSearch = (u: any) => {
  const name = ((u?.PreferredName ?? u?.Title ?? '') as string).trim().toLowerCase();
  if (!name) return false;
  return EXCLUDED_TOKENS.some(tok => name.includes(tok.toLowerCase()));
};


const getBestEmailLike = (u: any): string => {
  const email = (u?.WorkEmail ?? '').trim();
  const upnLike = (u?.UserName ?? u?.AccountName ?? '').trim();
  return email || upnLike;
};

const extractDomain = (val: string): string => {
  if (!val) return '';
  let s = val;
  const pipeIdx = s.lastIndexOf('|');
  if (pipeIdx >= 0) s = s.substring(pipeIdx + 1);
  const atIdx = s.lastIndexOf('@');
  return atIdx >= 0 ? s.substring(atIdx + 1).toLowerCase() : '';
};

const isAllowedDomain = (u: any): boolean => {
  const d = extractDomain(getBestEmailLike(u));
  return d.length > 0 && ALLOWED_EMAIL_DOMAINS.some(ad => ad.toLowerCase() === d);
};

export class spservices implements ISPServices {
  constructor(private context: WebPartContext) {
    sp.setup({
      spfxContext: {
        pageContext: { web: { absoluteUrl: this.context.pageContext.web.absoluteUrl } },
      },
    });
  }

  public async searchUsersNew(
    searchString: string,
    srchQry: string,
    isInitialSearch: boolean
  ): Promise<SearchResults> {

    // Query text
    let qrytext = '';
    if (isInitialSearch) {
      qrytext = `FirstName:${searchString}* OR LastName:${searchString}*`;
    } else {
      if (srchQry) qrytext = srchQry;
      else if (searchString) qrytext = searchString;
      if (qrytext.length <= 0) qrytext = `*`;
    }

    const searchProperties: string[] = [
      'FirstName','LastName','PreferredName','WorkEmail',
      'OfficeNumber','PictureURL',
      'WorkPhone',
      // mobil minden lehetséges neve People Search-ben:
      'MobilePhone','CellPhone','SPS-CellPhone','SPS-MOBILEPHONE',
      'JobTitle','Department','Skills','PastProjects',
      'BaseOfficeLocation','SPS-UserType','GroupId',
      'UserName','AccountName',
    ];

    try {
      const users = await sp.search(<ISearchQuery>{
        Querytext: qrytext,
        RowLimit: 500,
        EnableInterleaving: true,
        SelectProperties: searchProperties,
        SourceId: 'b09a7990-05ea-4af9-81ef-edfab16c4e31',
        SortList: [{ Property: 'LastName', Direction: SortDirection.Ascending }],
      });

      const primary = (users?.PrimarySearchResults ?? []) as any[];

      // 1) prefix + domain szűrés
      let items = primary
        .filter(u => !shouldHideUserFromSearch(u))
        .filter(u => isAllowedDomain(u))
        .map(u => {
          // 2) fotó normalizálás
          const email = (u?.WorkEmail || '').trim();
          const picture = email
            ? `/_layouts/15/userphoto.aspx?size=L&accountname=${encodeURIComponent(email)}`
            : u?.PictureURL;

          // 3) mobil egységesítése a People Search mezőiből
          const unifiedMobile =
            (u?.MobilePhone ||
             u?.CellPhone ||
             u?.['SPS-CellPhone'] ||
             u?.['SPS-MOBILEPHONE'] ||
             ''
            ).toString().trim();

          return { ...u, PictureURL: picture, MobilePhone: unifiedMobile };
        });

      // 4) Fallback Graphról: ahol még üres a MobilePhone, próbáljuk meg Entra-ból
      const needGraph = items.filter(u => !u.MobilePhone && u.WorkEmail);
      if (needGraph.length > 0) {
        try {
          const graphClient = await (this.context as any).msGraphClientFactory.getClient('3') as MSGraphClientV3;

          // Egyszerű: per-user hívások (kisebb listáknál bőven elég)
          await Promise.all(needGraph.map(async u => {
            try {
              const res = await graphClient
                .api(`/users/${encodeURIComponent(u.WorkEmail)}`)
                .select('mobilePhone')
                .get();
              const m = (res?.mobilePhone || '').toString().trim();
              if (m) u.MobilePhone = m;
            } catch { /* swallow */ }
          }));
        } catch { /* ha nincs consent, csak kihagyjuk a fallbacket */ }
      }

      const out = { ...users, PrimarySearchResults: items } as unknown as SearchResults;
      return out;

    } catch (error: any) {
      throw (error instanceof Error) ? error : new Error(String(error));
    }
  }
}
