import { WebPartContext } from "@microsoft/sp-webpart-base";
import { sp } from '@pnp/sp';
import { SearchResults, ISearchQuery, SortDirection } from '@pnp/sp/search';
import { ISPServices } from "./ISPServices";

/* =========================
   Filtering helpers (People Search)
   ========================= */

// NÉV ELEJI PREFIXEK kizárása (kis/nagybetű nem számít)
const EXCLUDED_PREFIXES = ['(X)', '(SZ)', 'ADM', 'guest', 'UPS', 'Test', 'Teszt'];

// CSAK EZEK A DOMAINEK legyenek láthatók
const ALLOWED_EMAIL_DOMAINS = ['value4real.com']; // <-- bővíthető: ['xy.com','contoso.com']

const shouldHideUserFromSearch = (u: any) => {
  const name = ((u?.PreferredName ?? u?.Title ?? '') as string).trim().toLowerCase();
  if (!name) return false;
  return EXCLUDED_PREFIXES.some(p => name.startsWith(p.toLowerCase()));
};

// People Search-ből “email-szerű” érték kiválasztása
const getBestEmailLike = (u: any): string => {
  const email = (u?.WorkEmail ?? '').trim();
  const upnLike = (u?.UserName ?? u?.AccountName ?? '').trim(); // claims / UPN jellegű érték
  return email || upnLike;
};

// claims formátum kezelése és domain kinyerése
const extractDomain = (val: string): string => {
  if (!val) return '';
  let s = val;
  const pipeIdx = s.lastIndexOf('|'); // pl. i:0#.f|membership|user@contoso.com
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
        pageContext: {
          web: {
            absoluteUrl: this.context.pageContext.web.absoluteUrl,
          },
        },
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
      if (srchQry) {
        qrytext = srchQry;
      } else if (searchString) {
        qrytext = searchString;
      }
      if (qrytext.length <= 0) qrytext = `*`;
    }

    const searchProperties: string[] = [
      'FirstName',
      'LastName',
      'PreferredName',
      'WorkEmail',
      'OfficeNumber',
      'PictureURL',
      'WorkPhone',
      // mobil minden lehetséges neve:
      'MobilePhone',
      'CellPhone',
      'SPS-CellPhone',
      'SPS-MOBILEPHONE',
      'JobTitle',
      'Department',
      'Skills',
      'PastProjects',
      'BaseOfficeLocation',
      'SPS-UserType',
      'GroupId',
      'UserName',
      'AccountName',
    ];

    try {
      const users = await sp.search(<ISearchQuery>{
        Querytext: qrytext,
        RowLimit: 500,
        EnableInterleaving: true,
        SelectProperties: searchProperties,
        SourceId: 'b09a7990-05ea-4af9-81ef-edfab16c4e31', // Local People Results
        SortList: [{ Property: 'LastName', Direction: SortDirection.Ascending }],
      });

      const primary = (users?.PrimarySearchResults ?? []) as any[];

      const filteredAndNormalized = primary
        .filter(u => !shouldHideUserFromSearch(u)) // prefix
        .filter(u => isAllowedDomain(u))           // domain
        .map(u => {
          // fotó URL normalizálás
          const email = (u?.WorkEmail || '').trim();
          const picture = email
            ? `/_layouts/15/userphoto.aspx?size=L&accountname=${encodeURIComponent(email)}`
            : u?.PictureURL;

          // mobil egységesítése
          const unifiedMobile =
            (u?.MobilePhone ||
             u?.CellPhone ||
             u?.['SPS-CellPhone'] ||
             u?.['SPS-MOBILEPHONE'] ||
             ''
            ).toString().trim();

          return {
            ...u,
            PictureURL: picture,
            MobilePhone: unifiedMobile, // innen kezdve mindig legyen kitöltve, ha van bármelyik forrásban
          };
        });

      // új példány visszaadása, nem írjuk felül a readonly mezőt helyben
      const out = { ...users, PrimarySearchResults: filteredAndNormalized } as unknown as SearchResults;
      return out;

    } catch (error: any) {
      throw (error instanceof Error) ? error : new Error(String(error));
    }
  }
}
