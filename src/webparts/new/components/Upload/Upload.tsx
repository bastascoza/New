import * as React from "react";
import styles from "./Upload.module.scss";
import { IUploadProps } from "./IUploadProps";
import { IUploadState } from "./IUploadState";
import { SPService } from "../../shared/service/SPService";
import { TextField, MaskedTextField } from "@fluentui/react/lib/TextField";
import { Stack, IStackProps, IStackStyles } from "@fluentui/react/lib/Stack";
import { Form, Formik, Field, FormikProps } from "formik";
import { Label } from "office-ui-fabric-react/lib/Label";
import * as yup from "yup";
import "semantic-ui-css/semantic.min.css";
import { Dropdown } from "semantic-ui-react";
import { Dialog, DialogType, DialogFooter } from "@fluentui/react/lib/Dialog";
import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import {
  DatePicker,
  // Dropdown,
  mergeStyleSets,
  PrimaryButton,
  IIconProps,
} from "office-ui-fabric-react";
import { sp } from "@pnp/sp";
import { DefaultButton, MessageBar, MessageBarType } from "@fluentui/react";
import { extendWith } from "lodash";

const modalPropsStyles = { main: { maxWidth: 450 } };
const modalProps = {
  isBlocking: true,
  styles: modalPropsStyles,
};

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

export default class Upload extends React.Component<
  IUploadProps,
  IUploadState
> {
  private cancelIcon: IIconProps = { iconName: "Clear" };
  private saveIcon: IIconProps = { iconName: "Save" };
  private _services: SPService = null;

  constructor(props: Readonly<IUploadProps>) {
    super(props);
    this.state = {
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
    //console.log('date',new Date().getDate().toString()); 
    this._services.getListSite("Sites").then((result) => {
      this.setState({
        SiteName: result,
      });
    });
  }

  public render(): React.ReactElement<IUploadProps> {
    const validate = yup.object().shape({
      file: yup.mixed().required("Please provide a file"),
      Site: yup.string().required("Please select a site name"),
      Name: yup.string().required("file name is required"),
    });

    return (
      <Formik
        validateOnChange={false}
        validateOnBlur={false}
        initialValues={{
          file: undefined,
          Site: "",
          Name: "",
        }}
        validationSchema={validate}
        onSubmit={(values, helpers) => {
          /* console.log("submit values", values);
          console.log('date',new Date().getTime());
          const file = values.file;
          const Site = values.Site;
          const Name = values.Name.concat(" "+this.formatDate(new Date())); */
          /* let body = {
            file:file,
            Site: Site,
            Name: Name
          }; */
          this._services._fileExists(values).then((exists) => {
            //console.log("values", values);
            //console.log("file exists", exists);
           if (!exists) {
              this._services.CreateFile(false, values);
            } else {
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
                    values.file.name.split(".").pop() +
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

              <Label className={styles.lblForm}>Upload File</Label>
              <input
                type="file"
                multiple={false}
                name="file"
                onChange={(event) => {
                  formik.setFieldValue("file", event.currentTarget.files[0]);
                }}
              />
              <div>
                {formik.touched.file && formik.errors.file ? (
                  <MessageBar
                    messageBarType={MessageBarType.error}
                    isMultiline={false}
                    onDismiss={() => {
                      formik.errors.file = undefined;
                    }}
                  >
                    {formik.errors.file}
                  </MessageBar>
                ) : undefined}
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
                ) : undefined}
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
              text="Reset"
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
    this._services.CreateFile( true,body);
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
      this._services.CreateFile(true);
    }else{
      let newbody = {
      file: body.file,
      Name: Name,
      Site: body.Site,
    };
    this._services.CreateFile(false,newbody);
  }    
}
  private formatDate = (date): string => {
    var date1 = new Date(date);

    var year = date1.getFullYear().toString();
    var month = (1 + date1.getMonth()).toString();
    month = month.length > 1 ? month : "0" + month;
    var day = date1.getDate().toString();
    day = day.length > 1 ? day : "0" + day;
    let hours = (date1.getHours()).toString();
    hours = hours.length > 1 ? hours : "0" + hours;
    let minutes = date1.getMinutes().toString();
    minutes = minutes.length > 1 ? minutes : "0" + minutes;
    return month + "-" + day + "-" + year + " " + hours + "-" + minutes;
  }
  
}
