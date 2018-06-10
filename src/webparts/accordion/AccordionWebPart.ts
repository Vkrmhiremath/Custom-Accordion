import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
  PropertyPaneCheckbox
} from '@microsoft/sp-webpart-base';

import * as strings from 'AccordionWebPartStrings';
import Accordion from './components/Accordion';
import { IAccordionProps } from './components/IAccordionProps';
import { IAccordionWebPartProps, ISPLists, ISPList } from './IAccordionWebPartProps';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export default class AccordionWebPart extends BaseClientSideWebPart<IAccordionWebPartProps> {


  private _selectedList: string;
  private _selectedSite: string;
  private _selectedColor: string;
  private _selectedField: string;

  private _listOptions: IPropertyPaneDropdownOption[];
  private _siteOptions: IPropertyPaneDropdownOption[];
  private _views: IPropertyPaneDropdownOption[];
  private _fields: IPropertyPaneDropdownOption[];


  public render(): void {
    const element: React.ReactElement<IAccordionProps> = React.createElement(
      Accordion,
      {
        title: this.properties.title,
        listName: this.properties.listName,
        siteUrl: this.properties.siteUrl,
        fieldName: this.properties.fieldName,
        context: this.context,
        sortDesc: this.properties.sortDesc
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    let contentProperty: any;
    return {
      pages: [
        {
          header: {
            description: "Configure  Accordion component"
          },
          groups: [
            {
              groupName: "Data Source",
              groupFields: [
                PropertyPaneDropdown('siteUrl', {
                  label: "Select SharePoint Site",
                  options: this._siteOptions,
                  selectedKey: this._selectedSite
                }),
                PropertyPaneDropdown('listName', {
                  label: "Select SharePoint List",
                  options: this._listOptions,
                  selectedKey: this._selectedList
                }),
                PropertyPaneDropdown('fieldName', {
                  label: "Select Fields to group by",
                  options: this._fields,
                  selectedKey: this._selectedField
                }),
                PropertyPaneCheckbox("sortDesc", {
                  text: "Sort Descending",
                  checked: false
                }),
              ]
            }
          ],
          displayGroupsAsAccordion: true
        }
      ]
    };
  }

  protected onPropertyPaneConfigurationStart(): void {
    this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'Configuring properties...');
    this._loadSites();
    this._loadLists(this.properties.siteUrl != undefined ? this.properties.siteUrl : "");
    this._loadFields(this.properties.siteUrl != undefined ? this.properties.siteUrl : "", this.properties.listName != undefined ? this.properties.listName : "");
    this._selectedList = this.properties.listName;
    this._selectedSite = this.properties.siteUrl;
    this._selectedField = this.properties.fieldName;
    this._refresComponentAndPanePanel();
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'Updating properties...');
    if (propertyPath === 'siteUrl' && newValue) {
      this._selectedSite = newValue;
      this._loadLists(this._selectedSite);
    }
    if (propertyPath === 'listName' && newValue) {
      this._selectedList = newValue;
      this._loadFields(this._selectedSite, this._selectedList);
    }
    if (this._selectedField !== null && this._selectedField !== "") {
      this._refresComponentAndPanePanel();
    } else {
      this._refreshPanePanel();
    }
  }

  private _refresComponentAndPanePanel() {
    this.context.propertyPane.refresh();
    this.context.statusRenderer.clearLoadingIndicator(this.domElement);
    this.render();
  }

  private _refreshPanePanel() {
    this.context.propertyPane.refresh();
    this.context.statusRenderer.clearLoadingIndicator(this.domElement);
  }

  // get all lists specific to the selected site.
  private _getLists(selectedSite): Promise<ISPLists> {
    var queryUrl: string = selectedSite;
    queryUrl += "/_api/lists?$select=Title,id,BaseTemplate";
    queryUrl += "&$orderby=Title";
    queryUrl += "&$filter=BaseTemplate%20eq%20";
    queryUrl += "100";
    queryUrl += "%20and%20Hidden%20eq%20false";

    return this.context.spHttpClient.get(queryUrl, SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {
      return response.json();
    });
  }

  // get all sites corresponding to current site collection.
  private _getSites(): Promise<ISPLists> {
    var queryUrl: string = this.context.pageContext.web.absoluteUrl;
    queryUrl += "/_api/web/webs?$select=Title,Url";
    queryUrl += "&$orderby=Title";
    return this.context.spHttpClient.get(queryUrl, SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {
      return response.json();
    });
  }

  // get all fields corresponding to the selected list .
  private _getFields(selectedSite, selectedList) {
    if (selectedSite && selectedSite) {
      var queryUrl: string = selectedSite;
      queryUrl += "/_api/web/lists/getbytitle('" + encodeURIComponent(selectedList) + "')/fields?$filter=Hidden eq false and ReadOnlyField eq false&$Select=Title";
      return this.context.spHttpClient.get(queryUrl, SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {
        return response.json();
      });
    }
  }

  // load all lists to listOptions 
  private _loadLists(selectedSite) {
    if (selectedSite != null && selectedSite != "") {
      const listsOptions = [];
      this._getLists(selectedSite).then((response: ISPLists) => {
        response.value.map((list: ISPList) => {
          var isSelected: boolean = false;
          if (this._selectedList == list.Id) {
            isSelected = true;
            this._selectedList = list.Id;
          }
          listsOptions.push({
            key: list.Title,
            text: list.Title
          });
        });
        this._listOptions = listsOptions;
      }).then(() => {
        this._refresComponentAndPanePanel();
      });
    }
  }

  // load all fields to fieldOptions 
  private _loadFields(selectedSite: string, selectedList: string): any {
    if (selectedSite !== null && selectedSite !== "" && selectedList !== null && selectedList !== "") {
      const fieldOptions = [];
      this._getFields(selectedSite, selectedList).then((response: any) => {
        response.value.map((field: any) => {
          var isSelected: boolean = false;
          if (this._selectedField === field.Title) {
            isSelected = true;
            this._selectedField = field.Title;
          }
          fieldOptions.push({
            key: field.Title,
            text: field.Title
          });
        });
        this._fields = fieldOptions;
      }).then(() => {
        this._refresComponentAndPanePanel();
      });
    }
  }

  // load all sites to siteOptions 
  private _loadSites() {
    const siteOptions = [];
    var siteText = this.context.pageContext.web.serverRelativeUrl.split('/');

    siteOptions.push({
      key: this.context.pageContext.web.absoluteUrl,
      text: this.context.pageContext.web.serverRelativeUrl.split('/')[siteText.length - 1]
    });
    this._getSites().then((response: ISPLists) => {
      response.value.map((list: ISPList) => {
        var isSelected: boolean = false;
        if (this._selectedSite == list.Url) {
          isSelected = true;
          this._selectedSite = list.Url;
        }
        siteOptions.push({
          key: list.Url,
          text: list.Title
        });
      });
      this._siteOptions = siteOptions;
    }).then(() => {
      this._refresComponentAndPanePanel();
    });
  }

}