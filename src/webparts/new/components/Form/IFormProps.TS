import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IFormProps {
  description: string;
  context: WebPartContext;
  siteUrl: string;
}
