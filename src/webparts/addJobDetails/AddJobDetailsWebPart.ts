import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './AddJobDetailsWebPart.module.scss';
import * as strings from 'AddJobDetailsWebPartStrings';
import { SPComponentLoader } from "@microsoft/sp-loader";


import "jquery";
import * as moment from "moment";
import "datatables";
import { sp } from "@pnp/pnpjs";
import "../../ExternalRef/css/StyleJob.css";
SPComponentLoader.loadCss(
  "https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css"
);

var alertify: any = require("../../ExternalRef/js/alertify.min.js");
declare var $;



export interface IAddJobDetailsWebPartProps {
  description: string;
}

export default class AddJobDetailsWebPart extends BaseClientSideWebPart <IAddJobDetailsWebPartProps> {

  public onInit(): Promise<void> {
    return super.onInit().then((_) => {
      sp.setup({
        spfxContext: this.context,
      });
    });
  }


  public render(): void {
    this.domElement.innerHTML = `
    <div class="loading-modal"> 
    <div class="spinner-border" role="status"> 
    <span class="sr-only"></span>
  </div></div>
    <div class="container"><label class="Heading">Site Details</label>
        <div class="row clsRowDiv">
          <div class="column col-xl-3 col-lg-3 col-md-12 col-sm-12 col-12">
            <label>Node ID</label>
            <input type="text" id="txtNode">
          </div>
          <div class="column col-xl-3 col-lg-3 col-md-12 col-sm-12 col-12">
            <label>SiteName</label>
            <input type="text" id="txtSiteName">
          </div>
          <div class="column col-xl-3 col-lg-3 col-md-12 col-sm-12 col-12">
            <label>Client</label>
            <input type="text" id="txtClient">
          </div>
          <div class="column col-xl-3 col-lg-3 col-md-12 col-sm-12 col-12">
            <label>Version</label>
            <input type="text" id="txtVersion">
          </div> 
        </div>
        <label class="Heading divProjectdetails" style="display:none">Project Details</label>
        <div class="row clsRowDiv divProjectdetails" style="display:none">
          <div class="column col-xl-3 col-lg-3 col-md-12 col-sm-12 col-12">
            <label id="lblProjects">Selected Projects</label>
            <ul id="selectedProjects">
            </ul>
          </div>
        </div>

        <label class="Heading" style="display:none">Project Details</label>
        <div class="row clsRowDiv" style="display:none">
          <div class="column col-xl-3 col-lg-3 col-md-12 col-sm-12 col-12">
            <label id="lblProjects">Projects</label>
            <select id="drpProjects">
            <option value="Select">Select</option>
            </select>
          </div>
          <div class="column col-xl-3 col-lg-3 col-md-12 col-sm-12 col-12">
          <label id="lblTasks">Tasks</label>
          <select id="drpTasks">
          <option value="Select">Select</option>
          </select>
        </div>
        </div>
<div class="row clsRowDiv" id="tblForTasks" style="display:none">
  <table>
  <thead>
  <tr>
    <th>Project Name</th>
    <th>Task Name</th>
    <th>Assignee</th>
    <th>Active</th>
  </tr>
  </thead>
  <tbody id="tbodyForTaskDetails">
  <tr>
    <td>MTX Design</td>
    <td>MTX Design</td>
    <td><input type="text"></td>
    <td><input type="checkbox" checked></td>
  </tr>
  </tbody>
</table>
</div>
</div>`;



$("#txtNode").blur(function()
{
  getSiteDetails($("#txtNode").val());
});
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

function getSiteDetails(NodeID)
{
  sp.web.lists.getByTitle("SiteList").items.select("*","Client/Title,Category/Title").expand("Client,Category").filter("NodeId eq '"+NodeID+"'").get().then((item)=>
  {
      if(item.length>0)
      {
          console.log(item);  
          $("#txtNode").val(item[0].NodeId);
          $("#txtSiteName").val(item[0].SiteName);
          $("#txtVersion").val(item[0].Version_x0023_);
          $("#txtClient").val(item[0].Client.Title);


          if(item[0].Category.length>0)
          {
              var html='';
              for(var i=0;i<item[0].Category.length;i++)
              {
                  html+="<li>"+item[0].Category[i].Title+"</li>";
              }

              $("#selectedProjects").html('');
              $("#selectedProjects").html(html);

              $(".divProjectdetails").show();
              $("#tblForTasks").show();
              
          }
          else
          {
            $(".divProjectdetails").hide();
            $("#tblForTasks").hide();
          }
      }
      else
      {
        alert("Can't Find Site")
      }
  }).catch((error)=>
  {
    ErrorCallBack(error, "getSiteDetails");
  });
}

async function ErrorCallBack(error, methodname) 
{
  try {
    var errordata = {
      Error: error.message,
      MethodName: methodname,
    };
    await sp.web.lists
      .getByTitle("ErrorLog")
      .items.add(errordata)
      .then(function (data) 
      {
        $(".loading-modal").removeClass("active");
        $("body").removeClass("body-hidden");
        AlertMessage("Something went wrong.please contact system admin");
      });
  } catch (e) {
    //alert(e.message);
    $(".loading-modal").removeClass("active");
    $("body").removeClass("body-hidden");
    AlertMessage("Something went wrong.please contact system admin");
  }
}
function AlertMessage(strMewssageEN) {
  alertify
    .alert()
    .setting({
      label: "OK",

      message: strMewssageEN,

      onok: function () {
        window.location.href = "#";
      },
    })
    .show()
    .setHeader("<em>Confirmation</em> ")
    .set("closable", false);
}
