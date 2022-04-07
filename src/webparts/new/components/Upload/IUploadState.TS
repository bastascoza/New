import { IDropdownOption } from "office-ui-fabric-react";
import { string } from "yup";
interface IdialogContentProps{
  type:any;
  title:string;
  subText:string;
}
export interface IUploadState {
    form?: any;
    file?:any;
    SiteName?: any[];
    projectName?: string;
    dialogContentProps?:IdialogContentProps;
    showDialog?:boolean;
  }
  
