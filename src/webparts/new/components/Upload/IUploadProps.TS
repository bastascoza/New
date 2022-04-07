import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IUploadProps {
  description: string;
  context: WebPartContext;
  siteUrl: string;
}
