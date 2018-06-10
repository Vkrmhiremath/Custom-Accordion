import * as React from 'react';
import { IAccordionProps, IQueryBuilders } from './IAccordionProps';
import styles from './Accordion.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import CollapsableAccordion from './Collapsable-Accordion/CollapsableAccordion';
require('linqjs');


export interface IAccordionStates {
  listData?: any[];
  listName?: string;
  viewName?: string;
  fieldName?: string;
  keyArray?: string[];
  filteredData?: any[]
}

export interface ListItem {
  Title: string;
}

export default class MyAccordion extends React.Component<IAccordionProps, IAccordionStates> {
  private _clientContext: any;

  public constructor(props: any) {
    super(props);
    this.state = {
      listData: [],
      listName: "",
      viewName: "",
      fieldName: "",
      keyArray: [],
      filteredData: [],
    };
    this._clientContext = this.props.context.spHttpClient;
  }

  private _listData;
  private _selectedKey;
  public render(): React.ReactElement<IAccordionProps> {
    this._listData = this._selectedKey === "All" || this._selectedKey === undefined ? this.state.listData : this.state.filteredData;

    if (this._validateProps() && this._comparePropsAndState()) {
      let _results: any;
      let _results_grouped: any;
      const params: IQueryBuilders = {
        siteUrl: this.props.siteUrl,
        listName: this.props.listName,
        groupByField: this.props.fieldName,
        isDescending: this.props.sortDesc
      };

      this._getListItems(params).then(res => {
        const value: any = res.value;
        let _newArray = [];
        let _keyArray = [];
        _keyArray.push("All");
        if (value.length > 0) {
          _results = this._groupBy(value, this.props.fieldName);
          _results_grouped = _results.forEach((result) => {
            let _key = result.key;
            let _item = [];
            _keyArray.push(_key);
            result.forEach(element => {
              _item.push([element, _key]);
            });
            _newArray.push(_item);
          });
        }
        this.setState({ listData: _newArray, listName: this.props.listName, fieldName: this.props.fieldName, keyArray: _keyArray });
      });
    }

    if (this._listData && this._listData.length > 0) {
      return (
        <div>
          <div className={styles.styledSelect + " " + styles.slate + " " + styles.floatRight}>
            <select onChange={this._onKeySelected.bind(this)}>{
              this.state.keyArray && this.state.keyArray.map((key) => {
                return <option value={key}>{key}</option>;
              })
            }</select>
          </div>
          <br /><br />
          <CollapsableAccordion isOpen={true} title={"List : " + this.props.listName}>
            {this._listData && this._listData.map((listItem, index) => {
              return (
                <CollapsableAccordion title={this.props.fieldName + " : " + listItem[0][1]}>
                  {listItem.map((listItemVal) => {
                    return <CollapsableAccordion title={listItemVal[0].Title}><span>{listItemVal[0].Title}</span> </CollapsableAccordion>;
                  })}
                </CollapsableAccordion>
              );
            })}
          </CollapsableAccordion>
        </div>
      );
    } else {
      return null;
    }

  }

  // on dropdown change event
  private _onKeySelected(e) {
    this._selectedKey = e.target.value;
    this._selectedKey === "" ? this._selectedKey = null : this._selectedKey = this._selectedKey
    this._filter(this.state.listData, this._selectedKey);

  }

  // validates the the webpart properties.
  private _validateProps() {
    return this._isNullOrEmpty(this.props.siteUrl) && this._isNullOrEmpty(this.props.listName) && this._isNullOrEmpty(this.props.fieldName);
  }

  // compares props & state 
  private _comparePropsAndState() {
    return this.state.listName !== this.props.listName && this.state.fieldName != this.props.fieldName;
  }

  // filter the listData with the dropdown filter. 
  private _filter(array, value) {
    if (value === "All") {
      this.setState({ filteredData: this.state.listData });
    } else {
      let _newArray = [];
      let _item = [];
      array.forEach((result) => {
        result.forEach(element => {
          let _key = element[1];
          if (element[1] === value)
            _item.push([element[0], _key]);
        });
      });
      _newArray.push(_item);
      this.setState({ filteredData: _newArray });
    }
  }

  // checks if the property is null or empty.
  private _isNullOrEmpty(propertyValue) {
    return propertyValue !== "" && propertyValue !== null;
  }

  // gets list items using the provided params.
  private _getListItems(params: IQueryBuilders): any {
    return this._clientContext
      .get(
        this._buildListQuery(params),
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  // build the rest API url to get the items from selected list.
  private _buildListQuery(params: IQueryBuilders): any {
    let queryURL: string =
      params.siteUrl +
      "/_api/web/lists/getByTitle('" +
      params.listName +
      "')/Items?";
    queryURL += params.isDescending ? "&$orderby=ID desc" : "";
    return queryURL;
  }

  // groups the array based on the provided fieldname
  private _groupBy(resultArray: any, fieldName: string): any {
    var groupedArray = resultArray.groupBy((field) => { return field[fieldName]; });
    return groupedArray;
  }
}
