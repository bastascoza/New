import { IDropdownOption } from "office-ui-fabric-react";
interface IdialogContentProps {
  type: any;
  title: string;
  subText: string;
}
export interface IFormState {
  form?: any;
  templateFile?: any;
  SiteName?: any;
  projectName?: string;
  dialogContentProps?: IdialogContentProps;
  showDialog?: boolean;

}
