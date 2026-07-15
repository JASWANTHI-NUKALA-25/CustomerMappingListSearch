import * as React from 'react';
//import styles from './Customermappinglistsearch.module.scss';
import type { ICustomermappinglistsearchProps } from './ICustomermappinglistsearchProps';
//import { escape } from '@microsoft/sp-lodash-subset';
import SearchComponent from '../../../component/SearchComponent';
import { ListConstants } from '../../../constants/generalConstants';


export default class Advancesearch extends React.Component<ICustomermappinglistsearchProps> {
  public render(): React.ReactElement<ICustomermappinglistsearchProps> {

    return (
      <SearchComponent
        context={this.props.context}
        listName={ListConstants.ListName}
        siteUrl={ListConstants.SiteUrl}
      />
    );
  }
}
