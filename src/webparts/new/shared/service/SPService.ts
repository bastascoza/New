import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import "@pnp/sp/files";

import "@pnp/sp/folders";
import { Web } from "@pnp/sp/webs";
import { IDropdownOption } from "office-ui-fabric-react";
import { Dialog } from "@microsoft/sp-dialog";

export class SPService {
  //const items = await sp.web.lists.getByTitle("Sites").items.select("Title", "Lookup/Title", "Lookup/ID").expand("Lookup")();

  public async getListSite(list: string): Promise<IDropdownOption[]> {
    let docSites: any[] = [];
    let listSites: any[] = [{ key: "", value: "", text: "" }];
    let Items: any[] = [];
    return new Promise<IDropdownOption[]>(async (resolve, reject) => {
      if (list.toLowerCase().trim() === "templates") {
        Items = await sp.web.lists
          .getByTitle(list)
          .items.select("File/Name, File/ServerRelativeUrl")
          .expand("File")
          .get();

        Items.map((item) => {
          docSites.push({
            key: item.File.ServerRelativeUrl,
            value: item.File.ServerRelativeUrl,
            text: item.File.Name,
          });
        });

        //console.log("Templates =>", docSites);
        resolve(docSites);
      } else {
        Items = await sp.web.lists
          .getByTitle(list)
          .items.select("Werfnaam, WerfId, FullWerfNummer")
          .get();

        Items.map((item) => {
          listSites.push({
            key: item.WerfId,
            value: item.WerfId,
            text: item.FullWerfNummer,
          });
        });
        //console.log("Werven =>", listSites);
        resolve(listSites);
      }
    });
  }
  public async copyFile( replace: boolean, body?: any) {
    if(replace && body === undefined){
      Dialog.alert('Please try after a minute!');
    }
    else{
    let destUrl = "/sites/DemoPO/Site";
    const _SiteId: number[] = [];
    await _SiteId.push(body.Site);
    let _fileCreatedUrl = "";
    let _redirectUrl: string = "";
    await sp.web
      .getFileByServerRelativePath(body.Template)
      .copyByPath(
        destUrl + "/" + body.Name + "." + body.Template.split(".").pop(),
        replace,
        false
      )
      .then(() => {
        sp.web.lists
          .getByTitle("Site")
          .items.orderBy("Created", false)
          .expand("File")
          .select("EncodedAbsUrl")
          .top(1)
          .get()
          .then((x) => {
            x.map((item) => {
              _fileCreatedUrl += item.File.ServerRelativeUrl;
              _redirectUrl += item.EncodedAbsUrl;
            });
          })
          .then(() => {
            sp.web
              .getFileByServerRelativeUrl(_fileCreatedUrl)
              .getItem()
              .then((y) => {
                y.update({
                  SiteId: {
                    results: _SiteId,
                  },
                });
              });
          })
          .then(() => {
            Dialog.alert("File Added Succesfully !").then(() => {
                window.location.href = _redirectUrl + "?web=1";
            });
          });
      })
      .catch((err) => {
        console.error(err);
      });
    }
  }
  public async CreateFile( replace: boolean, body?: any) {
    let destUrl = "/sites/DemoPO/Site/";
    if (replace === true && body === undefined){
      Dialog.alert('Please try after a minute');
    }
    else{
    const _SiteId: number[] = [];
    await _SiteId.push(body.Site);
    const file = await sp.web
      .getFolderByServerRelativeUrl(destUrl)
      .files.add(
        body.Name + "." + body.file.name.split(".").pop(),
        body.file,
        replace
      );
    const item = await file.file.getItem();
    await item
      .update({
        SiteId: {
          results: _SiteId,
        },
      })
      .then(() => {
        Dialog.alert("Added Successfully !").then(() => {
           window.location.href =
           "https://spobuild.sharepoint.com/sites/DemoPO/SitePages/Werfleiders.aspx";
        });
      });
    }
  }
  public _fileExists = async (body: any): Promise<Boolean> => {
    let exists: Boolean;
    let fileName = body.file?.name ? body.file.name : body.Template;
    return new Promise<Boolean>(async (resolve, reject) => {
      exists = await sp.web
        .getFolderByServerRelativePath("/sites/DemoPO/Site")
        .files.getByName(body.Name + "." + fileName.split(".").pop())
        .exists();
      resolve(exists);
    });
  }
  public async _getSameFileName(): Promise<any[]> {
    let filesName: any[] = [];
    return new Promise<any[]>(async (resolve, reject) => {
      let Items = await sp.web.lists
        .getByTitle("Site")
        .items.expand("File")
        .select("File/Name")
        .get();

      Items.map((item) => {
        filesName.push(item);
      });
      resolve(filesName);
    });
  }
}
