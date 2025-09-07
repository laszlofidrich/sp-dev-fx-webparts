import { WebPartContext } from "@microsoft/sp-webpart-base";
import { sp } from '@pnp/sp';
import { SearchResults, ISearchQuery, SortDirection } from '@pnp/sp/search';
import { ISPServices } from "./ISPServices";

/* =========================
   Filtering helpers (People Search)
   ========================= */
const EXCLUDED_PREFIXES = ['(X)', '(SZ)']; // edit as needed

// People Search typically returns PreferredName/Title fields
const shouldHideUserFromSearch = (u: any) => {
  const name = ((u?.PreferredName ?? u?.Title ?? '') as string).trim().toLowerCase();
  if (!name) return false;
  return EXCLUDED_PREFIXES.some(p => name.startsWith(p.toLowerCase()));
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

    // Build the query text exactly like your original
    let qrytext = '';
    if (isInitialSearch) {
      qrytext = `FirstName:${searchString}* OR LastName:${searchString}*`;
    } else {
      if (srchQry) {
        qrytext = srchQry;
      } else {
        if (searchString) qrytext = searchString;
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
      'MobilePhone',
      'JobTitle',
      'Department',
      'Skills',
      'PastProjects',
      'BaseOfficeLocation',
      'SPS-UserType',
      'GroupId',
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

      // IMPORTANT:
      // - Do NOT assign to users.PrimarySearchResults (it's readonly in the type).
      // - Instead, create a new object with a filtered copy.
      const primary = (users?.PrimarySearchResults ?? []) as any[];
      const filtered = primary.filter(u => !shouldHideUserFromSearch(u));

      // Return a new SearchResults-like object with the filtered list
      const out = { ...users, PrimarySearchResults: filtered } as unknown as SearchResults;
      return out;

    } catch (error: any) {
      throw (error instanceof Error) ? error : new Error(String(error));
    }
  }
}
