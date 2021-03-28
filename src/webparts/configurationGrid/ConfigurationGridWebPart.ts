import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPComponentLoader } from "@microsoft/sp-loader";
import styles from './ConfigurationGridWebPart.module.scss';
import * as strings from 'ConfigurationGridWebPartStrings';

import "jquery";
import { sp } from "@pnp/sp/presets/all";
//import "datatables";
SPComponentLoader.loadCss(
  "https://stackpath.bootstrapcdn.com/bootstrap/3.4.1/css/bootstrap.min.css"
);
SPComponentLoader.loadScript(
  "https://maxcdn.bootstrapcdn.com/bootstrap/3.4.1/js/bootstrap.min.js"
);

import "../../../node_modules/datatables/media/js/jquery.dataTables.min.js";
import "../../ExternalRef/css/alertify.min.css";
import "../../../node_modules/datatables/media/css/jquery.dataTables.min.css";
import "../../ExternalRef/css/ConfigurationStyle.css";
var alertify: any = require("../../ExternalRef/js/alertify.min.js");

declare var $;
var siteURL = "";

export interface IConfigurationGridWebPartProps {
  description: string;
}
export default class ConfigurationGridWebPart extends BaseClientSideWebPart<IConfigurationGridWebPartProps> {
  protected onInit(): Promise<void> {
    return super.onInit().then((_) => {
      siteURL = this.context.pageContext.web.absoluteUrl;
      sp.setup({
        spfxContext: this.context,
      });
    });
  }
  public render(): void {

    this.domElement.innerHTML = `
    <div class="container">
    <ul class="nav nav-tabs">
      <li class="active"><a data-toggle="tab" href="#div-division">Division</a></li>
      <li>
      <a data-toggle="tab" href="#div-business">BusinessDivision</a></li>
      <li><a data-toggle="tab" href="#div-client">Client</a></li>
      <li><a data-toggle="tab" href="#div-project">Project</a></li>
      <li><a data-toggle="tab" href="#div-task">TaskMasterList</a></li>
    </ul>
    <div class="tab-content">
      <div id="div-division" class="tab-pane fade in active">
      <div class="head-btn-sec">
      <input class="btn btn-primary" type="button" id="btndiv" value="Create">
      </div>
        <table id="tblForDivision">
        <thead>
          <tr>
            <th>S.No</th>
            <th>Title</th> 
            <th>Action</th>
          </tr>
          </thead>
        <tbody id="tblBodyForDivision">
        <tr>
        <td>1</td>
        <td>Ram</td>
        <td>KKDI</td>
      </tr>
        </tbody>
        </table>
      </div>
      <div id="div-business"  class="tab-pane fade">
      <div class="head-btn-sec">
      <input class="btn btn-primary" type="button" id="btnbusdiv" value="Create">
      </div>
          <table id="tblForBusinessDivision">
          <thead>
          <tr>
            <th>S.No</th>
            <th>Title</th>
            <th>Action</th>
          </tr>
          </thead>
          <tbody id="tblBodyForBusinessDivision">
          <tr>
          <td>1</td>
          <td>Ram</td>
          <td>KKDI</td>
        </tr>
          </tbody>
            </table>
      </div>
      <div id="div-client"  class="tab-pane fade">
      <div class="head-btn-sec">
      <input class="btn btn-primary" type="button" id="btncli" value="Create">
      </div>
      <table id="tblForClient">
      <thead>
      <tr>
        <th>S.No</th>
        <th>Title</th>
        <th>Action</th>
      </tr>
      </thead>
      <tbody id="tblBodyForClient">
      <tr>
      <td>1</td>
      <td>Ram</td>
      <td>KKDI</td>
    </tr>
      </tbody>
        </table>
      </div>
      <div id="div-project"  class="tab-pane fade">
      <div class="head-btn-sec">
      <input class="btn btn-primary" type="button" id="btnpro" value="Create">
      </div>
          <table id="tblForProjects">
          <thead>
          <tr>
            <th>S.No</th>
            <th>Title</th>
            <th>Action</th>
          </tr>
          </thead>
          <tbody id="tblBodyForProjects">
          <tr>
          <td>1</td>
          <td>Ram</td>
          <td>KKDI</td>
        </tr>
          </tbody>
            </table>
      </div>
      <div id="div-task"  class="tab-pane fade">
      <div class="head-btn-sec">
      <input class="btn btn-primary" type="button" id="btntask" value="Create">
      </div>
      <table id="tblForTaskMasterList">
      <thead>
      <tr>
        <th>S.No</th>
        <th>Division</th>
        <th>BusinessDivision</th>
        <th>Project</th>
        <th>Tasks</th>
        <th>Priority</th>
        <th>Action</th>
      </tr>
      </thead>
      <tbody id="tblBodyForTaskMasterList">
      <tr>
      <td>1</td>
      <td>-</td>
      <td>-</td>
      <td>-</td>
      <td>-</td>
      <td>-</td>
      <td>-</td>
      
    </tr>
      </tbody>
        </table>
      </div>
      </div>
    </div>`;


    getDivisions();
    getBusinessDivisions();
    getProjects();
    getClient();
    getTaskMasterList();
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
   

async function getDivisions() 
{
  var html = "";
  await sp.web.lists
    .getByTitle("Division")
    .items.get()
    .then(async (items: any[]) => {
      // console.log(items);
      for (var i = 0; i < items.length; i++) 
      {
        html+='<tr><td>'+(i+1)+'</td><td>'+items[i].Title+'</td><td><a href="#"><span class="icon-img icon-view"></span></a><a href="#"><span class="icon-img icon-edit"></span></a></td></tr>';
      }
      $("#tblBodyForDivision").html("");
      await $("#tblBodyForDivision").html(html);
      
    })
    .catch(function (error) 
    {
      ErrorCallBack(error, "getDivisions");
    });
}
async function getBusinessDivisions() {
  var html = "";
  await sp.web.lists
    .getByTitle("BusinessDivision")
    .items.get()
    .then(async (items: any[]) => {
      for (var i = 0; i < items.length; i++) {
        html+='<tr><td>'+(i+1)+'</td><td>'+items[i].Title+'</td><td><a href="#"><span class="icon-img icon-view"></span></a><a href="#"><span class="icon-img icon-edit"></span></a></td></tr>';
      }
      $("#tblBodyForBusinessDivision").html("");
      await $("#tblBodyForBusinessDivision").html(html);
    })
    .catch(function (error) {
      ErrorCallBack(error, "getBusinessDivisions");
    });
}
async function getClient(){
  var html ="";
  await sp.web.lists.getByTitle("Client").items.get().then(async (items: any[]) => {
    // console.log(items);
    for (var i = 0; i < items.length; i++){
      html+='<tr><td>'+(i+1)+'</td><td>'+items[i].Title+'</td><td><a href="#"><span class="icon-img icon-view"></span></a><a href="#"><span class="icon-img icon-edit"></span></a></td></tr>';
    }
    $("#tblBodyForClient").html("");
    await $("#tblBodyForClient").html(html);

})
.catch(function (error) {
  ErrorCallBack(error, "getClients");
});
}
async function getProjects() {
  var html = "";
  await sp.web.lists
    .getByTitle("Projects")
    .items.get()
    .then(async (items: any[]) => {
      for (var i = 0; i < items.length; i++) {
        
        html+='<tr><td>'+(i+1)+'</td><td>'+items[i].Title+'</td><td><a href="#"><span class="icon-img icon-view"></span></a><a href="#"><span class="icon-img icon-edit"></span></a></td></tr>';
      }
      $("#tblBodyForProjects").html("");
      await $("#tblBodyForProjects").html(html);
    })
    .catch(function (error) {
      ErrorCallBack(error, "getProjects");
    });
}
async function getTaskMasterList() {
  var html = "";
  await sp.web.lists
    .getByTitle("TaskMasterList")
    .items.get()
    .then(async (items: any[]) => {
      for (var i = 0; i < items.length; i++) {
        
        html+='<tr><td>'+(i+1)+'</td><td>'+items[i].Division+'</td><td>'+items[i].BusinessDivision+'</td><td>'+items[i].Projects+'</td><td>'+items[i].Tasks+'</td><td>'+items[i].Priority+'</td><td><a href="#"><span class="icon-img icon-view"></span></a><a href="#"><span class="icon-img icon-edit"></span></a></td></tr>';
      }
      $("#tblBodyForTaskMasterList").html("");
      await $("#tblBodyForTaskMasterList").html(html);
      $('#tblForDivision').DataTable();
      $('#tblForBusinessDivision').DataTable();
      $('#tblForClient').DataTable();
      $('#tblForProjects').DataTable();
      $('#tblForTaskMasterList').DataTable();
    })
    .catch(function (error) {
      ErrorCallBack(error, "getTaskMasterList");
    });
}
/* This is place for common  functionalities start*/
function AlertMessage(Message) {
  alertify
    .alert()
    .setting({
      label: "OK",

      message: Message,

      onok: function () {
        //window.location.href = siteURL + "/SitePages/RequestDashboard.aspx";
        window.location.href = "#";
      },
    })
    .show()
    .setHeader("<em>Confirmation</em> ")
    .set("closable", false);
}


async function ErrorCallBack(error, methodname) {
  console.log(error);
  AlertMessage("Something went wrong.please contact system admin");
}

function getUrlParameter(param) {
  var url = window.location.href
    .slice(window.location.href.indexOf("?") + 1)
    .split("&");
  for (var i = 0; i < url.length; i++) {
    var urlparam = url[i].split("=");
    if (urlparam[0] == param) {
      return urlparam[1];
    }
  }
}

/* This is place for common  functionalities end*/