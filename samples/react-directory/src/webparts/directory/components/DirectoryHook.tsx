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
   Safety-net filter for names
   ========================= */
const EXCLUDED_PREFIXES = ['(X)', '(SZ)','ADM']; // edit as needed

const shouldHideUserFromSearch = (u: any) => {
  const name = ((u?.PreferredName ?? u?.Title ?? '') as string).trim().toLowerCase();
  if (!name) return false;
  return EXCLUDED_PREFIXES.some((p) => name.startsWith(p.toLowerCase()));
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
                WorkPhone: user.WorkPhone,
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
        ? users.PrimarySearchResults.filter((u: any) => !shouldHideUserFromSearch(u))
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

  const _searchUsers = async (searchText: string): Promise<void> => {
    try {
      setstate({ ...state, searchText, isLoading: true });
      if (searchText.length > 0) {
        const searchProps: string[] =
          props.searchProps && props.searchProps.length > 0
            ? props.searchProps.split(',')
            : ['FirstName', 'LastName', 'WorkEmail', 'Department'];

        let qryText = '';
        const finalSearchText: string = searchText ? searchText.replace(/ /g, '+') : searchText;

        if (props.clearTextSearchProps) {
          const tmpCTProps: string[] =
            props.clearTextSearchProps.indexOf(',') >= 0
              ? props.clearTextSearchProps.split(',')
              : [props.clearTextSearchProps];

          if (tmpCTProps.length > 0) {
            searchProps.map((srchprop, index) => {
              const ctPresent: any[] = tmpCTProps.filter((o) => {
                return o.toLowerCase() === srchprop.toLowerCase();
              });
              if (ctPresent.length > 0) {
                if (index === searchProps.length - 1) {
                  qryText += `${srchprop}:${searchText}*`;
                } else {
                  qryText += `${srchprop}:${searchText}* OR `;
                }
              } else {
                if (index === searchProps.length - 1) {
                  qryText += `${srchprop}:${finalSearchText}*`;
                } else {
                  qryText += `${srchprop}:${finalSearchText}* OR `;
                }
              }
            });
          } else {
            searchProps.map((srchprop, index) => {
              if (index === searchProps.length - 1) {
                qryText += `${srchprop}:${finalSearchText}*`;
              } else {
                qryText += `${srchprop}:${finalSearchText}* OR `;
              }
            });
          }
        } else {
          searchProps.map((srchprop, index) => {
            if (index === searchProps.length - 1) {
              qryText += `${srchprop}:${finalSearchText}*`;
            } else {
              qryText += `${srchprop}:${finalSearchText}* OR `;
            }
          });
        }

        const users = await _services.searchUsersNew('', qryText, false);

        const cleaned =
          users && users.PrimarySearchResults
            ? users.PrimarySearchResults.filter((u: any) => !shouldHideUserFromSearch(u))
            : [];

        setstate({
          ...state,
          searchText,
          indexSelectedKey: '0',
          users: cleaned,
          isLoading: false,
          errorMessage: '',
          hasError: false,
        });
        setalphaKey('0');
      } else {
        setstate({ ...state, searchText: '' });
        await _searchByAlphabets(true);
      }
    } catch (err: any) {
      setstate({ ...state, errorMessage: err.message, hasError: true });
    }
  };

  const _searchBoxChanged = (newvalue: string): void => {
    setCurrentPage(1);
    setstate((prev) => ({ ...prev, searchText: newvalue }));
  };

  useEffect(() => {
    const debouncedSearch = setTimeout(() => {
      if (state.searchText !== undefined) {
        _searchUsers(state.searchText);
      }
    }, 300);
    return () => clearTimeout(debouncedSearch);
  }, [state.searchText]);

  const _sortPeople = async (sortField: string): Promise<void> => {
    let _users = [...state.users];
    _users = _users.sort((a: any, b: any) => {
      switch (sortField) {
        case 'Location':
          if ((a.BaseOfficeLocation || '').toUpperCase() < (b.BaseOfficeLocation || '').toUpperCase()) {
            return -1;
          }
          if ((a.BaseOfficeLocation || '').toUpperCase() > (b.BaseOfficeLocation || '').toUpperCase()) {
            return 1;
          }
          return 0;

        default:
          if ((a[sortField] || '').toUpperCase() < (b[sortField] || '').toUpperCase()) {
            return -1;
          }
          if ((a[sortField] || '').toUpperCase() > (b[sortField] || '').toUpperCase()) {
            return 1;
          }
          return 0;
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

  const onOptionSelect = (ev: any, data: OptionOnSelectData) => {
    _sortPeople(data.optionValue as string);
  };

  const handleSearchKeyPress = React.useCallback(
    (ev: React.KeyboardEvent<HTMLInputElement>) => {
      if (ev.key === 'Enter') {
        _searchUsers(state.searchText);
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
          onChange={(_, data) => _searchBoxChanged(data.value as string)}
        />
        <div className={mergeClasses(fluentStyles.alphabets, fluentStyles.horizontal)}>
          <Overflow minimumVisible={2}>
            <TabList
              selectedValue={state.indexSelectedKey}
              onTabSelect={_alphabetChange}
              className={fluentStyles.tabList}
            >
              {az.map((index: string) => {
                return (
                  <OverflowItem key={index} id={index}>
                    <Tab value={index} key={index}>
                      {index}
                    </Tab>
                  </OverflowItem>
                );
              })}
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
