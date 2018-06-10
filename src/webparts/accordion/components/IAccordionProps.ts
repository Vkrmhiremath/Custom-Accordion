import { IWebPartContext } from "@microsoft/sp-webpart-base";
import { DisplayMode } from "@microsoft/sp-core-library";

export interface IAccordionProps {
  title: string;
  listName: string;
  siteUrl: string;
  fieldName: string;
  context: IWebPartContext;
  sortDesc: boolean;
}


export interface IQueryBuilders {
  listName?: string;
  siteUrl?: string;
  groupByField?: string;
  isDescending?: boolean;
}

