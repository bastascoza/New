import { BaseWebPartContext, WebPartContext } from "@microsoft/sp-webpart-base";

 
export interface INewProps {
  context: WebPartContext; 
  description: string;
  siteUrl:string;
}
