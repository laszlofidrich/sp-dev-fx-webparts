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
    zIndex: 0, //stop the browser resize handle from piercing the overflow menu
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
   ADDED: simple safety-net filter
   ========================= */
const EXCLUDED_PREFIXES = ['(X)', '(SZ)']; // edit this list anytime

const shouldHideUserFromSearch = (u: any) => {
  const name = ((u?.PreferredName ?? u?.Title ?? '') as string).trim().toLowerCase();
  if (!name) return false;
  return EXCLUDED_PREFIXES.some(p => name.startsWith(p.toLowerCase()));
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
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const [pagedItems, setPagedItems] = useState<any[]>([]);
  const [pageSize, setPageSize] = useState<number>(
    props.pageSize ? props.pageSize : 10
  );
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
      ? // eslint-disable-next-line @typescript-eslint/no-explicit-any
        pagedItems.map((user: any, i) => {
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
                Location: user.OfficeNumber
                  ? user.OfficeNumber
                  : user.BaseOfficeLocation,
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
      if (props.searchFirstName)
        users = await _services.searchUsersNew('', `FirstName:a*`, false);
      else users = await _services.searchUsersNew('a', '', true);
    } else {
      if (props.searchFirstName)
        users = await _services.searchUsersNew(
          '',
          `FirstName:${alphaKey}*`,
          false
        );
      else users = await _services.searchUsersNew(`${alphaKey}`, '', true);
    }

    // ADDED: clean results before setting state
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
      setstate({ ...state, searchText: searchText, isLoading: true });
      if (searchText.length > 0) {
        const searchProps: string[] =
          props.searchProps && props.searchProps.length > 0
            ? props.searchProps.split(',')
            : ['FirstName', 'LastName', 'WorkEmail', 'Department'];
        let qryText = '';
        const finalSearchText: string = searchText
          ? searchText.replace(/ /g, '+')
          : searchText;
        if (props.clearTextSearchProps) {
          const tmpCTProps: string[] =
            props.clearTextSearchProps.indexOf(',') >= 0
              ? props.clearTextSearchProps.split(',')
              : [props.clearTextSearchProps];
          if (tmpCTProps.length > 0) {
            searchProps.map((srchprop, index) => {
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              const ctPresent: any[] = tmpCTProps.filter((o) => {
                return o.toLowerCase() === srchprop.toLowerCase();
              });
              if (ctPresent.length > 0) {
                if (index === searchProps.length - 1) {
                  qryText += `${srchprop}:${searchText}*`;
                } else qryText += `${srchprop}:${searchText}* OR `;
              } else {
                if (index === searchProps.length - 1) {
                  qryText += `${srchprop}:${finalSearchText}*`;
                } else qryText += `${srchprop}:${finalSearchText}* OR `;
              }
            });
          } else {
            searchProps.map((srchprop, index) => {
              if (index === searchProps.length - 1)
                qryText += `${srchprop}:${finalSearchText}*`;
              else qryText += `${srchprop}:${finalSearchText}* OR `;
            });
          }
        } else {
          searchProps.map((srchprop, index) => {
            if (index === searchProps.length - 1)
              qryText += `${srchprop}:${finalSearchText}*`;
            else qryText += `${srchprop}:${finalSearchText}* OR `;
          });
        }
        console.log(qryText);
        const users = await _services.searchUsersNew('', qryText, false);

        // ADDED: clean results before setting state
        const cleaned =
          users && users.PrimarySearchResults
            ? users.PrimarySearchResults.filter((u: any) => !shouldHideUserFromSearch(u))
            : [];

        setstate({
          ...state,
          searchText: searchText,
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
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    _users = _users.sort((a: any, b: any) => {
      switch (sortField) {
        // Sort by Location
        case 'Location':
          if (
            (a.BaseOfficeLocation || '').toUpperCase() <
            (b.BaseOfficeLocation || '').toUpperCase()
          ) {
            return -1;
          }
          if (
            (a.BaseOfficeLocation || '').toUpperCase() >
            (b.BaseOfficeLocation || '').toUpperCase()
          ) {
            return 1;
          }
          return 0;

        default:
          if (
            (a[sortField] || '').toUpperCase() <
            (b[sortField] || '').toUpperCase()
          ) {
            return -1;
          }
          if (
            (a[sortField] || '').toUpperCase() >
            (b[sortField] || '').toUpperCase()
          ) {
            return 1;
          }
          return 0;
      }
    });
    setstate({ ...state, users: _users, searchString: sortField });
  };
