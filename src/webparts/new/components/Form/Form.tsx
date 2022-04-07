import * as React from "react";
import styles from "./Form.module.scss";
import { IFormProps } from "./IFormProps";
import { IFormState } from "./IFormState";
import { SPService } from "../../shared/service/SPService";
import { TextField, MaskedTextField } from "@fluentui/react/lib/TextField";
import { Stack, IStackProps, IStackStyles } from "@fluentui/react/lib/Stack";
import { Form, Formik, Field, FormikProps } from "formik";
import { Label } from "office-ui-fabric-react/lib/Label";
import * as yup from "yup";
import { Dropdown } from "semantic-ui-react";
import "semantic-ui-css/semantic.min.css";
import { Dialog, DialogType, DialogFooter } from "@fluentui/react/lib/Dialog";
import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import {
  DatePicker,
  mergeStyleSets,
  PrimaryButton,
  IIconProps,
  IDropdownOption,
  Grid,
  VirtualizedComboBox,
  IComboBoxOption,
} from "office-ui-fabric-react";
import { DefaultButton, MessageBar, MessageBarType } from "@fluentui/react";
import { sp } from "@pnp/sp";
import { ValueLabelProps } from "@material-ui/core";

const stackTokens = { childrenGap: 50 };
const iconProps = { iconName: "Calendar" };
const stackStyles: Partial<IStackStyles> = { root: { width: 650 } };
const columnProps: Partial<IStackProps> = {
  tokens: { childrenGap: 15 },
  styles: { root: { width: 300 } },
};
const controlClass = mergeStyleSets({
  control: {
    margin: "0 0 15px 0",
    maxWidth: "300px",
  },
});
const modalPropsStyles = { main: { maxWidth: 450 } };
const modalProps = {
  isBlocking: true,
  styles: modalPropsStyles,
};
export default class ReactFormik extends React.Component<
  IFormProps,
  IFormState
> {
  private cancelIcon: IIconProps = { iconName: "Cancel" };
  private saveIcon: IIconProps = { iconName: "Save" };
  private _services: SPService = null;

  constructor(props: Readonly<IFormProps>) {
    super(props);
    this.state = {
      templateFile: [],
      SiteName: [],
    };
    sp.setup({
      spfxContext: this.props.context,
    });

    this._services = new SPService();
  }

  private getFieldProps = (formik: FormikProps<any>, field: string) => {
    return {
      ...formik.getFieldProps(field),
      errorMessage: formik.errors[field] as string,
    };
  }

  public componentDidMount(): void {
    this._services.getListSite("Templates").then((result) => {
      this.setState({
        templateFile: result,
      });
    });
    this._services.getListSite("Sites").then((result) => {
      this.setState({
        SiteName: result,
      });
    });
  }
  public render(): React.ReactElement<IFormProps> {
    const validate = yup.object().shape({
      Template: yup.string().required("Please select a template file"),
      Site: yup.string().required("Please select a site name"),
      Name: yup.string().required("file name is required"),
    });

    return (
      <Formik
        enableReinitialize
        validateOnChange={false}
        validateOnBlur={false}
        initialValues={{
          Template: "",
          Site: "",
          Name: "",
        }}
        validationSchema={validate}
        onSubmit={(values, helpers) => {
          /* const Template = values.Template;
          const Site = values.Site;
          const Name = values.Name.concat(" "+this.formatDate(new Date()));
          let body = {
            Template:Template,
            Site: Site,
            Name: Name
          }; */
         
           this._services._fileExists(values).then((exists) => {
            //console.log('submit value',values);
            if (!exists) {
              this._services.copyFile(false,values);
            } else {
             // this._services.copyFile(true);
               this.setState({
                form: values,
                showDialog: true,
                dialogContentProps: {
                  type: DialogType.normal,
                  title: "File exists!",
                  subText:
                    'File with name "' +
                    values.Name +
                    "." +
                    values.Template.split(".").pop() +
                    '" already exists !',
                },
              });
            }
          }); 
        }}
      >
        {(formik) => (
          <div className={styles.reactFormik}>
            <Stack>
              <Label className={styles.lblForm}>Current User</Label>
              <PeoplePicker
                context={this.props.context as any}
                personSelectionLimit={1}
                showtooltip={true}
                showHiddenInUI={false}
                principalTypes={[PrincipalType.User]}
                ensureUser={true}
                disabled={true}
                defaultSelectedUsers={[
                  this.props.context?.pageContext?.user.email as any,
                ]}
              />

              <Label className={styles.lblForm}>Template File</Label>
              <Dropdown
                fluid
                placeholder="select a file template"
                search
                name="Template"
                selection
                options={this.state.templateFile}
                {...this.getFieldProps(formik, "Template")}
                onChange={(event, option) => {
                  formik.setFieldValue("Template", option.value);
                }}
              />
              <div>
                {formik.touched.Site && formik.errors.Template?.length > 0 ? (
                  <MessageBar
                    messageBarType={MessageBarType.error}
                    isMultiline={false}
                    onDismiss={() => {
                      formik.errors.Template = "";
                    }}
                  >
                    {formik.errors.Template}
                  </MessageBar>
                ) : null}
              </div>
              <Label className={styles.lblForm}>Site</Label>
              <Dropdown
                placeholder="select a site"
                required
                fluid
                search
                selection
                options={this.state.SiteName}
                {...this.getFieldProps(formik, "Site")}
                onChange={(event, option) => {
                  formik.setFieldValue("Site", option.value);
                }}
              />
              <div>
                {formik.touched.Site && formik.errors.Site?.length > 0 ? (
                  <MessageBar
                    messageBarType={MessageBarType.error}
                    isMultiline={false}
                    onDismiss={() => {
                      formik.errors.Site = "";
                    }}
                  >
                    {formik.errors.Site}
                  </MessageBar>
                ) : null}
              </div>
              <Label className={styles.lblForm}>File Name</Label>
              <TextField
                autoComplete={"off"}
                {...this.getFieldProps(formik, "Name")}
              />
            </Stack>
            <PrimaryButton
              type="submit"
              text="Save"
              iconProps={this.saveIcon}
              className={styles.btnsForm}
              onClick={formik.handleSubmit as any}
            />
            <PrimaryButton
              text="Cancel"
              iconProps={this.cancelIcon}
              className={styles.btnsForm}
              onClick={formik.handleReset as any}
            />
            <div>
              <Dialog
                hidden={!this.state.showDialog}
                dialogContentProps={this.state.dialogContentProps}
                modalProps={modalProps}
                onDismiss={() => this.setState({ showDialog: false })}
              >
                <DialogFooter>
                  <PrimaryButton
                    onClick={() => {
                      this._handleReplace(this.state.form);
                    }}
                    text="replace"
                  />
                  <DefaultButton
                    onClick={() => {
                      this._hadnleFileExists(this.state.form);
                    }}
                    text="keep them both"
                  />
                </DialogFooter>
              </Dialog>
            </div>
          </div>
        )}
      </Formik>
    );
  }
  private async _handleReplace(body?: any) {
    this.setState({
      showDialog: false,
    });
    this._services.copyFile(true,body);
  }
  private async _hadnleFileExists(body?: any) {
    this.setState({
      showDialog: false,
    });
    const filesName = await this._services._getSameFileName();
    //console.log("fileName", filesName);
    const Name = body.Name.concat(" "+ this.formatDate(new Date()));
    let waitMinute : boolean;
     filesName.map(x=>{
     if(x.File.Name.split('.').shift().indexOf(Name) > -1 ) {
       waitMinute = true;
     }
     else {waitMinute=false;}
     return waitMinute;
    });
    //console.log('typof',typeof(waitMinute),'\n value',waitMinute);
    if(waitMinute === true){
      this._services.copyFile(true);
    }else{
      let newbody = {
      Template: body.Template,
      Name: Name,
      Site: body.Site,
    };
    this._services.copyFile(false,newbody);
    /* let replaceName: string = "";
    let y: number = 0;
    await this._services
      ._getSameFileName()
      .then((result) => {
        result.map((item) => {
          if (
            body.Name.trim() === item.File.Name.split("(").shift().trim() &&
            item.File.Name.split("(").length > 1
          ) {
            let nameNoExtension = item.File.Name.split(".").shift();
            let numChar = nameNoExtension.split("(").pop();
            y = 1;
            if (item.File.Name.split(".").pop() === body.Template.split(".").pop()) {
              let identity: number = parseInt(numChar);
              let index: number[] = [];
              index.push(identity);
              y +=   Math.max(...index) ;
            }
          }
          else {
            y = 1;
          }
        });
      })
      .then(() => {
        replaceName += replaceName + body.Name + " (" + y + ")";
        let newbody = {
          Template: body.Template,
          Name: replaceName,
          Site: body.Site,
        };
       // this._services.copyFile(newbody,false);
      }); */
    }
  }
  private formatDate = (date): string => {
    var date1 = new Date(date);
    var year = date1.getFullYear().toString();
    var month = (1 + date1.getMonth()).toString();
    month = month.length > 1 ? month : "0" + month;
    var day = date1.getDate().toString();
    day = day.length > 1 ? day : "0" + day;
    let hours = (date1.getHours() ).toString();
    hours = hours.length > 1 ? hours : "0" + hours;
    let minutes = date1.getMinutes().toString();
    minutes = minutes.length > 1 ? minutes : "0" + minutes;
    return month + "-" + day + "-" + year + " " + hours + "-" + minutes;
  }
}
