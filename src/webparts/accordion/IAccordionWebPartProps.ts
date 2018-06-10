
export interface IAccordionWebPartProps {
  title: string;
  listName: string;
  siteUrl: string;
  fieldName: string;
  listData: any;
  sortDesc: boolean;
}
/**
 * @interface
 * Defines a collection of SharePoint lists
 */
export interface ISPLists {
  value: ISPList[];
}

/**
 * @interface
 * Defines a SharePoint list
 */
export interface ISPList {
  Title: string;
  Id: string;
  BaseTemplate: string;
  Url: string;
}


