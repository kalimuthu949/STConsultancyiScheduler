import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './JobdetailsgridWebPart.module.scss';
import * as strings from 'JobdetailsgridWebPartStrings';



import { sp } from "@pnp/sp/presets/all";
import "jquery";
import "../../ExternalRef/css/jobdetailsgrid.css";
var alertify: any = require("../../ExternalRef/js/alertify.min.js");

declare var $;

import "../../../node_modules/datatables/media/js/jquery.dataTables.min.js";
import "../../../node_modules/datatables/media/css/jquery.dataTables.min.css";

import { SPComponentLoader } from "@microsoft/sp-loader";
SPComponentLoader.loadCss(
  "https://cdn.jsdelivr.net/npm/bootstrap@5.0.0-beta2/dist/css/bootstrap.min.css"
);

export interface IJobdetailsgridWebPartProps {
  description: string;
}


export default class JobdetailsgridWebPart extends BaseClientSideWebPart<IJobdetailsgridWebPartProps> {
  public onInit(): Promise<void> {
    return super.onInit().then(() => {
      sp.setup({
         spfxContext: this.context
         });

    });
  }
  public render(): void {
    this.domElement.innerHTML = `
     <div class=container>
     <h3 class="text-center">Job Details </h3>
  <div class="btntext text-end py-2">
    <button class="btn  buttoncolor" id="btnCreate" type="submit">Create</button>
  </div>

     <table class="table table-bordered" id="tableForIScheduleJoblist">
    <thead>
      <tr>
        <th>S.No</th>
        <th>Title</th>
        <th>Client</th>
        <th>SiteName</th>
        <th>NodeID</th>
        <th>SiteType</th>
        <th>Projects</th>
        <th>VersionID</th>
        <th>Action</th>
      </tr>
    </thead>
    <tbody id="IScheduleJoblist">
      <tr>
        <td></td>
        <td></td>
        <td></td>
      </tr>

      <tr>
      <td></td>
      <td></td>
      <td></td>
      </tr>
      <tr>
      <td></td>
      <td></td>
      <td></td>
      </tr>
    </tbody>
  </table>
     </div>`;
  
     getIScheduleJoblist();
    }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    }; 
  }
}
async function getIScheduleJoblist(){
  var html ="";
  var htmldata = "";
  
  
  
  await sp.web.lists.getByTitle("IscheduleJobList").items.get().then((items: any[]) => {
    // console.log(items);
    for (var i=0; i<items.length; i++)
    {
     var element = items[i].Projects.split(";");
     console.log(element);
     if(element.length > 1){
      for(var j = 0; j < element.length - 1;j++){
        console.log(htmldata)
        htmldata +=  `<div> ${element[j]}</div>`;
      }
     }
     else{
      htmldata = items[i].Projects
    }
    
      html += `<tr><td>${i+1}<td>${items[i].Title}</td><td>${items[i].Client}</td><td>${items[i].SiteName}<td>${items[i].NodeID}</td>${htmldata}<td>${items[i].SiteType}<td>${htmldata}</td> <td>${items[i].VersionID}</td><td><a href="#"><span class="icon-img icon-view"></span></a></td></tr>`

    }
    $("#IScheduleJoblist").html("");
    $("#IScheduleJoblist").html(html);
    $("#tableForIScheduleJoblist").DataTable();
})
.catch(function (error) {
  ErrorCallBack(error, "getIScheduleJoblist");
});
}

/* This is place for common  functionalities start*/

async function ErrorCallBack(error, methodname) {
  console.log(error);
  AlertMessage("Something went wrong.please contact system admin");
}

function AlertMessage(Message) {
  alertify
    .alert()
    .setting({
      label: "OK",

      message: Message,

      onok: function () {
        // window.location.href = siteURL + "/SitePages/ConfigurationGrid.aspx";
        //window.location.href = "#";
      },
    })
    .show()
    .setHeader("<em>Confirmation</em> ")
    .set("closable", false);
}