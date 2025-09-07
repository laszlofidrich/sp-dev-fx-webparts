import { WebPartContext } from "@microsoft/sp-webpart-base";

import { sp } from '@pnp/sp';
import { SearchResults, ISearchQuery, SortDirection } from '@pnp/sp/search';
import { ISPServices } from "./ISPServices";

// === Filtering helpers ===

// Add more prefixes later, e.g. '(sz)' — case-insensitive
const EXCLUDED_PREFIXES = ['(X)', '(SZ)'];

/** filter for SharePoint People Search results */
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

      // Filter out entries like "(X) John Doe" or "(SZ) Jane Doe"
      if (users?.PrimarySearchResults?.length) {
        users.PrimarySearchResults = users.PrimarySearchResults
          .filter(u => !shouldHideUserFromSearch(u))
          .map(u => {
            // Normalize photo URL to the large user photo if WorkEmail is present
            if (u?.PictureURL && u?.WorkEmail) {
              return {
                ...u,
                PictureURL: `/_layouts/15/userphoto.aspx?size=L&accountname=${u.WorkEmail}`,
              };
            }
            return u;
          });
      }

      return users;
    } catch (error) {
      throw new Error(error as any);
    }
  }
}
