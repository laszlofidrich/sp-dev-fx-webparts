import * as React from 'react';
import { useEffect, useState } from 'react';
import { Stack, IStackTokens } from 'office-ui-fabric-react/lib/Stack';
import { WebPartTitle } from '@pnp/spfx-controls-react';
import {
  Dropdown,
  Option,
  OptionOnSelectData,
  Overflow,
  OverflowItem,
  SearchBox,
  Field,
  Title2,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  Tab,
  TabList,
  tokens,
  makeStyles,
  mergeClasses,
} from '@fluentui/react-components';
import { People48Filled } from '@fluentui/react-icons';

import { ISPServices } from '../../../SPServices/ISPServices';
import { IDirectoryProps } from './IDirectoryProps';
import Paging from './Pagination/Paging';
import { OverflowAlphabetsMenu } from './OverflowAlphabetsMenu/OverflowAlphabetsMenu';
import styles from './Directory.module.scss';
import { PersonaCard } from './PersonaCard/PersonaCard';
import { spservices } from '../../../SPServices/spservices';
import { IDirectoryState } from './IDirectoryState';
import * as strings from 'DirectoryWebPartStrings';
import { Shimmer } from './Shimmer/Shimmer';

const wrapStackTokens: IStackTokens = { childrenGap: 30 };

const useFluentStyles = makeStyles({
  alphabets: {
    backgroundColor: tokens.colorNeutralBackground2,
    overflow: 'hidden',
    padding: '5px',
    zIndex: 0,
    borderRadius: '5px',
  },
  horizontal: {
    height: 'fit-content',
    minWidth: '100%',
  },
  tabList: {
    justifyContent: 'center',
  },
});

/* =========================
   Safety-net filters (CONTAINS)
   ========================= */

// legalább 3 karakter után keressünk
const MIN_SEARCH_LEN = 3;

// csak ezek a domainek látszódjanak
const ALLOWED_EMAIL_DOMAINS = ['value4real.com','ofysol.com'];

// névben CONTAINS tiltó tokenek
const EXCLUDED_TOKENS = ['(X)', '(SZ)', 'ADM', 'guest', 'UPS', 'Teszt', 'Test'];

const shouldHideUserFromSearch = (u: any) => {
  const name = ((u?.PreferredName ?? u?.Title ?? '') as string).trim().toLowerCase();
  if (!name) return false;
  return EXCLUDED_TOKENS.some(tok => name.includes(tok.toLowerCase()));
};

// E-mail/UPN-szerű érték kiválasztása
const getBestEmailLike = (u: any): string => {
  const email = (u?.WorkEmail ?? '').trim();
  const upnLike = (u?.UserName ?? u?.AccountName ?? '').trim();
  return email || upnLike;
};

// Claims formátum kezelése, domain kinyerése
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
  return d.length > 0 && ALLOWED_EMAIL_DOMAINS.some((ad) => ad.toLowerCase() === d);
};

const DirectoryHook: React.FC<IDirectoryProps> = (props) => {
  const _services: ISPServices = new spservices(props.context);
  const [az, setaz] = useState<string[]>([]);
  const [alphaKey, setalphaKey] = useState<string>('A');
  const [state, setstate] = useState<IDirectoryState>({
    users: [],
    isLoading: true,
    errorMessage: '',
    hasError: false,
    indexSelectedKey: 'A',
    searchString: 'LastName',
    searchText: '',
  });

  const orderOptions = [
    { value: 'FirstName', text: 'First Name' },
    { value: 'LastName', text: 'Last Name' },
    { value: 'Department', text: 'Department' },
    { value: 'Location', text: 'Location' },
    { value: 'JobTitle', text: 'Job Title' },
  ];

  const color = props.context.microsoftTeams ? 'white' : '';

  // Paging
  const [pagedItems, setPagedItems] = useState<any[]>([]);
  const [pageSize, setPageSize] = useState<number>(props.pageSize ? props.pageSize : 10);
  const [currentPage, setCurrentPage] = useState<number>(1);

  const _onPageUpdate = async (pageno?: number): Promise<void> => {
    const currentPge = pageno ? pageno : currentPage;
    const startItem = (currentPge - 1) * pageSize;
    const endItem = currentPge * pageSize;
    const filItems = state.users.slice(startItem, endItem);
    setCurrentPage(currentPge);
    setPagedItems(filItems);
  };

  const diretoryGrid =
    pagedItems && pagedItems.length > 0
      ? pagedItems.map((user: any, i) => {
          return (
            <PersonaCard
              context={props.context}
              key={'PersonaCard' + i}
              profileProperties={{
                DisplayName: user.PreferredName,
                Title: user.JobTitle,
                PictureUrl: user.PictureURL,
                Email: user.WorkEmail,
                Department: user.Department,
                // prefer MobilePhone if present
                WorkPhone: user.MobilePhone || user.WorkPhone,
                Location: user.OfficeNumber ? user.OfficeNumber : user.BaseOfficeLocation,
              }}
            />
          );
        })
      : [];

  const _loadAlphabets = (): void => {
    const alphabets: string[] = [];
    for (let i = 65; i < 91; i++) {
      alphabets.push(String.fromCharCode(i));
    }
    setaz(alphabets);
  };

  const _alphabetChange = async (item?: any): Promise<void> => {
    setstate({
      ...state,
      searchText: '',
      indexSelectedKey: item.target.innerText,
      isLoading: true,
    });
    setalphaKey(item.target.innerText);
    setCurrentPage(1);
  };

  const _searchByAlphabets = async (initialSearch: boolean): Promise<void> => {
    setstate({ ...state, isLoading: true, searchText: '' });
    let users = null;
    if (initialSearch) {
      if (props.searchFirstName) {
        users = await _services.searchUsersNew('', `FirstName:a*`, false);
      } else {
        users = await _services.searchUsersNew('a', '', true);
      }
    } else {
      if (props.searchFirstName) {
        users = await _services.searchUsersNew('', `FirstName:${alphaKey}*`, false);
      } else {
        users = await _services.searchUsersNew(`${alphaKey}`, '', true);
      }
    }

    const cleaned =
      users && users.PrimarySearchResults
        ? users.PrimarySearchResults
            .filter((u: any) => !shouldHideUserFromSearch(u))
            .filter((u: any) => isAllowedDomain(u))
        : [];

    setstate({
      ...state,
      searchText: '',
      indexSelectedKey: initialSearch ? 'A' : state.indexSelectedKey,
      users: cleaned,
      isLoading: false,
      errorMessage: '',
      hasError: false,
    });
  };

  // Hard guard: do nothing if less than 3 chars
  const _searchUsers = async (searchText: string): Promise<void> => {
    try {
      const trimmed = (searchText || '').trim();
      if (trimmed.length > 0 && trimmed.length < MIN_SEARCH_LEN) {
        // stop spinner if we showed it earlier
        setstate((prev) => ({ ...prev, isLoading: false }));
        return;
      }

      setstate((prev) => ({ ...prev, searchText, isLoading: true }));

      if (trimmed.length >= MIN_SEARCH_LEN) {
        const searchProps: string[] =
          props.searchProps && props.searchProps.length > 0
            ? props.searchProps.split(',')
            : ['FirstName', 'LastName', 'WorkEmail', 'Department'];

        let qryText = '';
        const finalSearchText: string = trimmed.replace(/ /g, '+');

        if (props.clearTextSearchProps) {
          const tmpCTProps: string[] =
            props.clearTextSearchProps.indexOf(',') >= 0
              ? props.clearTextSearchProps.split(',')
              : [props.clearTextSearchProps];

          if (tmpCTProps.length > 0) {
            searchProps.forEach((srchprop, index) => {
              const ctPresent: any[] = tmpCTProps.filter((o) => o.toLowerCase() === srchprop.toLowerCase());
              qryText += `${srchprop}:${ctPresent.length > 0 ? trimmed : finalSearchText}*`;
              if (index !== searchProps.length - 1) qryText += ' OR ';
            });
          } else {
            searchProps.forEach((srchprop, index) => {
              qryText += `${srchprop}:${finalSearchText}*`;
              if (index !== searchProps.length - 1) qryText += ' OR ';
            });
          }
        } else {
          searchProps.forEach((srchprop, index) => {
            qryText += `${srchprop}:${finalSearchText}*`;
            if (index !== searchProps.length - 1) qryText += ' OR ';
          });
        }

        const users = await _services.searchUsersNew('', qryText, false);

        const cleaned =
          users && users.PrimarySearchResults
            ? users.PrimarySearchResults
                .filter((u: any) => !shouldHideUserFromSearch(u))
                .filter((u: any) => isAllowedDomain(u))
            : [];

        setstate((prev) => ({
          ...prev,
          searchText,
          indexSelectedKey: '0',
          users: cleaned,
          isLoading: false,
          errorMessage: '',
          hasError: false,
        }));
        setalphaKey('0');
      } else {
        // empty query -> back to alphabet A
        setstate((prev) => ({ ...prev, searchText: '' }));
        await _searchByAlphabets(true);
      }
    } catch (err: any) {
      setstate((prev) => ({ ...prev, errorMessage: err.message, hasError: true }));
    }
  };

  const _searchBoxChanged = (newvalue: string): void => {
    setCurrentPage(1);
    setstate((prev) => ({ ...prev, searchText: newvalue || '' }));
  };

  // Debounce, and only fire when len==0 (reset) or len>=3
  useEffect(() => {
    const t = setTimeout(() => {
      const txt = (state.searchText || '').trim();
      if (txt.length === 0) {
        _searchUsers(''); // reset to alphabet
      } else if (txt.length >= MIN_SEARCH_LEN) {
        _searchUsers(txt);
      } else {
        // ensure no spinner stays on for 1–2 chars
        setstate((prev) => ({ ...prev, isLoading: false }));
      }
    }, 300);
    return () => clearTimeout(t);
  }, [state.searchText]);

  const _sortPeople = async (sortField: string): Promise<void> => {
    let _users = [...state.users];
    _users = _users.sort((a: any, b: any) => {
      switch (sortField) {
        case 'Location': {
          const aVal = (a.BaseOfficeLocation || '').toUpperCase();
          const bVal = (b.BaseOfficeLocation || '').toUpperCase();
          if (aVal < bVal) return -1;
          if (aVal > bVal) return 1;
          return 0;
        }
        default: {
          const aVal = (a[sortField] || '').toUpperCase();
          const bVal = (b[sortField] || '').toUpperCase();
          if (aVal < bVal) return -1;
          if (aVal > bVal) return 1;
          return 0;
        }
      }
    });
    setstate({ ...state, users: _users, searchString: sortField });
  };

  useEffect(() => {
    setPageSize(props.pageSize);
    if (state.users) {
      _onPageUpdate();
    }
  }, [state.users, props.pageSize]);

  useEffect(() => {
    if (alphaKey.length > 0 && alphaKey !== '0') {
      _searchByAlphabets(false);
    }
  }, [alphaKey]);

  useEffect(() => {
    _loadAlphabets();
    _searchByAlphabets(true);
  }, [props]);

  // Enter only triggers when >=3 chars (vagy üres -> reset)
  const onOptionSelect = (ev: any, data: OptionOnSelectData) => {
    _sortPeople(data.optionValue as string);
  };

  const handleSearchKeyPress = React.useCallback(
    (ev: React.KeyboardEvent<HTMLInputElement>) => {
      if (ev.key === 'Enter') {
        const txt = (state.searchText || '').trim();
        if (txt.length >= MIN_SEARCH_LEN || txt.length === 0) {
          _searchUsers(txt);
        }
      }
    },
    [state.searchText]
  );

  const fluentStyles = useFluentStyles();

  return (
    <div className={styles.directory}>
      <WebPartTitle
        displayMode={props.displayMode}
        title={props.title}
        updateProperty={props.updateProperty}
      />
      <div className={styles.searchBox}>
        <SearchBox
          type="search"
          placeholder={strings.SearchPlaceHolder}
          className={styles.searchTextBox}
          value={state.searchText}
          onKeyDown={handleSearchKeyPress}
          onChange={(_, data) => _searchBoxChanged((data?.value as string) ?? '')}
        />
        <div className={mergeClasses(fluentStyles.alphabets, fluentStyles.horizontal)}>
          <Overflow minimumVisible={2}>
            <TabList
              selectedValue={state.indexSelectedKey}
              onTabSelect={_alphabetChange}
              className={fluentStyles.tabList}
            >
              {az.map((index: string) => (
                <OverflowItem key={index} id={index}>
                  <Tab value={index} key={index}>
                    {index}
                  </Tab>
                </OverflowItem>
              ))}
              <OverflowAlphabetsMenu onTabSelect={_alphabetChange} tabs={az} />
            </TabList>
          </Overflow>
        </div>
      </div>

      {state.isLoading ? (
        <div style={{ marginTop: '10px' }}>
          <Shimmer />
        </div>
      ) : (
        <>
          {state.hasError ? (
            <div style={{ marginTop: '10px' }}>
              <MessageBar intent="error">
                <MessageBarBody>
                  <MessageBarTitle>{state.errorMessage}</MessageBarTitle>
                </MessageBarBody>
              </MessageBar>
            </div>
          ) : (
            <>
              {!pagedItems || pagedItems.length === 0 ? (
                <div className={styles.noUsers}>
                  <People48Filled style={{ fontSize: '54px', color: color }} />
                  <Title2 style={{ marginLeft: 5, color: color }}>{strings.DirectoryMessage}</Title2>
                </div>
              ) : (
                <>
                  <div style={{ width: '100%', display: 'inline-block' }}>
                    <Paging
                      totalItems={state.users.length}
                      itemsCountPerPage={pageSize}
                      onPageUpdate={_onPageUpdate}
                      currentPage={currentPage}
                    />
                  </div>
                  <div className={styles.dropDownSortBy}>
                    <Stack horizontal horizontalAlign="center" wrap tokens={wrapStackTokens}>
                      <Field label={strings.DropDownPlaceLabelMessage}>
                        <Dropdown
                          placeholder={strings.DropDownPlaceHolderMessage}
                          value={state.searchString}
                          onOptionSelect={onOptionSelect}
                        >
                          {orderOptions.map((option: any) => (
                            <Option key={option.value} value={option.value}>
                              {option.text}
                            </Option>
                          ))}
                        </Dropdown>
                      </Field>
                    </Stack>
                  </div>
                  <Stack
                    horizontal
                    horizontalAlign={props.useSpaceBetween ? 'space-between' : 'center'}
                    wrap
                    tokens={wrapStackTokens}
                  >
                    {diretoryGrid}
                  </Stack>
                  <div style={{ width: '100%', display: 'inline-block' }}>
                    <Paging
                      totalItems={state.users.length}
                      itemsCountPerPage={pageSize}
                      onPageUpdate={_onPageUpdate}
                      currentPage={currentPage}
                    />
                  </div>
                </>
              )}
            </>
          )}
        </>
      )}
    </div>
  );
};

export default DirectoryHook;
