import { WebPartContext } from "@microsoft/sp-webpart-base";

import { sp } from '@pnp/sp';
import { SearchResults, ISearchQuery, SortDirection } from '@pnp/sp/search';
import { ISPServices } from "./ISPServices";

// === Filtering helpers (add near the top of spservices.ts) ===

// Add more prefixes later, e.g. '(sz)'
const EXCLUDED_PREFIXES = ['(X)', '(SZ)'];

/** client-side safety check */
const shouldHideUser = (u: any) => {
  const name = (u?.displayName ?? '').trim().toLowerCase();
  const disabled = u?.accountEnabled === false;
  const badPrefix = EXCLUDED_PREFIXES.some(p => name.startsWith(p.toLowerCase()));
  return disabled || badPrefix;
};

/** builds a Graph $filter that excludes disabled users and names starting with our prefixes */
function buildUsersFilter(): string {
  const notStarts = EXCLUDED_PREFIXES
    .map(p => `startswith(displayName,'${p.replace(/'/g, "''")}') eq false`)
    .join(' and ');
  // Also exclude Guests if you want: add "and userType eq 'Member'"
  return `(accountEnabled eq true) and ${notStarts}`;
}

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
    if (isInitialSearch)
      qrytext = `FirstName:${searchString}* OR LastName:${searchString}*`;
    else {
      if (srchQry) qrytext = srchQry;
      else {
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
        SourceId: 'b09a7990-05ea-4af9-81ef-edfab16c4e31',
        SortList: [
          { Property: 'LastName', Direction: SortDirection.Ascending },
        ],
      });
      if (users && users.PrimarySearchResults.length > 0) {
        for (
          let index = 0;
          index < users.PrimarySearchResults.length;
          index++
        ) {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          let user: any = users.PrimarySearchResults[index];
          if (user.PictureURL) {
            user = {
              ...user,
              PictureURL: `/_layouts/15/userphoto.aspx?size=L&accountname=${user.WorkEmail}`,
            };
            users.PrimarySearchResults[index] = user;
          }
        }
      }
      return users;
    } catch (error) {
      throw new Error(error);
    }
  }
}
