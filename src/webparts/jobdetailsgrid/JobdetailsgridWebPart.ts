import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { escape } from "@microsoft/sp-lodash-subset";

import styles from "./JobdetailsgridWebPart.module.scss";
import * as strings from "JobdetailsgridWebPartStrings";

import { sp } from "@pnp/sp/presets/all";
import "jquery";
import "../../ExternalRef/css/jobdetailsgrid.css";
import "../../ExternalRef/css/loader.css";
var alertify: any = require("../../ExternalRef/js/alertify.min.js");


declare var $;

import "../../../node_modules/datatables/media/js/jquery.dataTables.min.js";
import "../../../node_modules/datatables/media/css/jquery.dataTables.min.css";

import { SPComponentLoader } from "@microsoft/sp-loader";
SPComponentLoader.loadCss(
  "https://cdn.jsdelivr.net/npm/bootstrap@5.0.0-beta2/dist/css/bootstrap.min.css"
);

var siteURL = "";
var FilteredManager =[];
var currentuser = "";
var YesChecked="Yes";
var TaskRefId=[];

export interface IJobdetailsgridWebPartProps {
  description: string;
}

export default class JobdetailsgridWebPart extends BaseClientSideWebPart<IJobdetailsgridWebPartProps> {
  public onInit(): Promise<void> {
    return super.onInit().then(() => {
      sp.setup({
        spfxContext: this.context,
      });
    });
  }
  public render(): void {
    currentuser = this.context.pageContext.user.email;
    siteURL = this.context.pageContext.web.absoluteUrl;
    this.domElement.innerHTML = `
    <span style="display:none" class="loader">
<img class="loader-spin"/>
</span>
     <div class=container>
  <div class="btntext text-end py-2">
    <button class="btn  buttoncolor" id="btnCreate" type="submit" style="display:none">Create</button>
  </div>

     <table class="table table-bordered" id="tableForIScheduleJoblist">
    <thead>
      <tr>
        <th>S.No</th>
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
    $(".loader").show();
    getIscheduletaskList();
    

    $("#btnCreate").click(function () {
      location.href = `${siteURL}/SitePages/AddJob.aspx`;
    });
    $(document).on("click", ".viewjob", function () {
      location.href = `${siteURL}/SitePages/ViewJob.aspx?Itemid=${$(this).attr(
        "data-id"
      )}`;
    });
    $(document).on("click", ".editjob", function () {
      location.href = `${siteURL}/SitePages/EditJob.aspx?Itemid=${$(this).attr(
        "data-id"
      )}`;
    });
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}

async function getmanagerfromsite() {
  var ManagerInfo = [];
  await sp.web.siteGroups
    .getByName("ITimeSheetManagers")
    .users.get()
    .then(function (result) {
      for (var i = 0; i < result.length; i++) {
        ManagerInfo.push({
          Title: result[i].Title,
          ID: result[i].Id,
          Email: result[i].Email,
        });
      }
      FilteredManager = ManagerInfo.filter((manager)=>{return (manager.Email == currentuser)});
      console.log(FilteredManager);
        if (FilteredManager.length>0) 
        {
          getIScheduleJoblist(true);
        } else {
          getIScheduleJoblist(false);
        }
    })
    .catch(function (err) {
      alert("Group not found: " + err);
    });
}

async function getIScheduleJoblist(flag) {
  var html = "";
    await sp.web.lists
      .getByTitle("SiteMasterList")
      .items.top(5000).get()
      .then((items: any[]) => {
        for (var i = 0; i < items.length; i++) 
        {
 
          ///var element = items[i].Projects.split(";");
          var element = items[i].Category.split(";");

          var htmldata = "";
          if (element.length > 1) {
            for (var j = 0; j < element.length - 1; j++) {
              console.log(htmldata);
              htmldata += `<div> ${element[j]}</div>`;
            }
          } else {
            ///htmldata = items[i].Projects;
            htmldata = items[i].Category;
          }
          if(flag)
          {
            html += `<tr><td>${i + 1}</td><td>${items[i].Client}</td><td>${
              items[i].SiteName
            }<td>${items[i].NodeID}</td>${htmldata}<td>${
              items[i].SiteType
            }<td>${htmldata}</td> <td>${
              items[i].VersionID
            }</td><td><a href="#" class="viewjob" data-id=${
              items[i].ID
            }><span class="icon-img icon-view"></a><a href="#" class="editjob" data-id=${
              items[i].ID
            }><span class="icon-img icon-edit"></a></td></tr>`;
          }
          else
          {
            $("#btnCreate").hide();
            for(var j=0;j<TaskRefId.length;j++)
            {
              if(items[i].ID==TaskRefId[j].ReferenceNumber)
              {
            html += `<tr><td>${i + 1}</td><td>${items[i].Client}</td><td>${
              items[i].SiteName
            }<td>${items[i].NodeID}</td>${htmldata}<td>${
              items[i].SiteType
            }<td>${htmldata}</td> <td>${
              items[i].VersionID
            }</td><td><a href="#" class="viewjob" data-id=${
              items[i].ID
            }><span class="icon-img icon-view"></a></td></tr>`;
              }
            }
          }
        }
        $("#IScheduleJoblist").html("");
        $("#IScheduleJoblist").html(html);
        $("#tableForIScheduleJoblist").DataTable();
        $(".loader").hide();
        
      })
      .catch(function (error) {
        ErrorCallBack(error, "getIScheduleJoblist");
      });
  
}

async function getIscheduletaskList()
{
  await sp.web.lists.getByTitle("IscheduletaskList").items.select("*").filter("Active eq '"+YesChecked+"' and AssignedToEmail eq '"+currentuser+"'").top(5000).get().then(async (item)=>
  {
      var ItemInfo=[];
      if(item.length>0)
      {
          for(var i=0; i<item.length; i++)
          {
            if(currentuser==item[i].AssignedToEmail)
                {
                  ItemInfo.push({"ReferenceNumber":item[i].ReferenceNumber,"AssigneeName":item[i].AssigneeName,"AssignedToEmail":item[i].AssignedToEmail});
              }
          }

          TaskRefId = ItemInfo.reduce(function (item, e1) {  
            var matches = item.filter(function (e2)  
            { return e1.ReferenceNumber == e2.ReferenceNumber});  
            if (matches.length == 0) {  
                item.push(e1);  
            }  
            return item;  
        }, []);  
        
      
      }

      getmanagerfromsite();
      
  }).catch((error)=>
  {
    ErrorCallBack(error, "getIscheduletaskList");
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
        window.location.href = siteURL+"/SitePages/JobDetails.aspx";
      },
    })
    .show()
    .setHeader("<em>Confirmation</em> ")
    .set("closable", false);
}
