//import { Version } from '@microsoft/sp-core-library';
import { UrlQueryParameterCollection, Log,Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './ViewJobDetailsWebPart.module.scss';
import * as strings from 'ViewJobDetailsWebPartStrings';
import { SPComponentLoader } from "@microsoft/sp-loader";
import { MSGraphClient } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

import "jquery";
import * as moment from "moment";
import "datatables";
import "moment";

import { sp } from "@pnp/pnpjs";
import "../../ExternalRef/css/StyleJob.css";
import "../../ExternalRef/css/loader.css";

import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";

SPComponentLoader.loadCss(
  "https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css"
);
SPComponentLoader.loadCss(
  "https://maxcdn.bootstrapcdn.com/bootstrap/3.4.1/css/bootstrap.min.css"
);
SPComponentLoader.loadScript(
  "https://maxcdn.bootstrapcdn.com/bootstrap/3.4.1/js/bootstrap.min.js"
);

import "../../ExternalRef/css/alertify.min.css";
var alertify: any = require("../../ExternalRef/js/alertify.min.js");
declare var $;

var that;
var Itemid;
var taskdetails=[];
var tval='';
var YesChecked="Yes";
var FilteredManager = [];
var currentuser = "";

var flagmangerornot=false;
var siteURL="";

export interface IViewJobDetailsWebPartProps {
  description: string;
}

export default class ViewJobDetailsWebPart extends BaseClientSideWebPart <IViewJobDetailsWebPartProps> {

  public onInit(): Promise<void> {
    return super.onInit().then((_) => {
      sp.setup({
        spfxContext: this.context,
      });
    });
  }

  public render(): void {
    that=this;
    siteURL=this.context.pageContext.web.absoluteUrl;
    currentuser = this.context.pageContext.user.email;
    console.log(currentuser);
    this.domElement.innerHTML = `
    <span style="display:none" class="loader">
<img class="loader-spin"/>
</span>
    <div class="loading-modal"> 
    <div class="spinner-border" role="status"> 
    <span class="sr-only"></span>
  </div></div>
    <div class="container ischedule"><label class="Heading">Site Details</label>
        <div class="row clsRowDiv">
          <div class="column col-xl-3 col-lg-3 col-md-12 col-sm-12 col-12">
            <label>Node ID</label>
            <input type="text" id="txtNode"> 
          </div>
          <div class="column col-xl-3 col-lg-3 col-md-12 col-sm-12 col-12">
            <label>Site Name</label>
            <input type="text" id="txtSiteName">
          </div>
          <div class="column col-xl-3 col-lg-3 col-md-12 col-sm-12 col-12">
            <label>Site Type</label>
            <input type="text" id="txtSiteType">
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
        <label class="Heading divProjectdetails">Project Details</label>
        <div class="row clsRowDiv divProjectdetails">
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
<div class="row clsRowDiv" id="tblForTasks">
  <table>
  <thead>
  <tr>
    <th>Project Name</th>
    <th>Task Name</th>
    <th>Assignee</th>
    <th>Due Date</th>
    <th>View</th>
  </tr>
  </thead>
  <tbody id="tbodyForTaskDetails">
  <tr>
  <td colspan="5">No Tasks</td>
  </tr>
  </tbody>
</table>
</div>

<!-- Modal -->
<div class="new-container">
  <!-- Modal -->
  <div class="modal fade" id="myModal" role="dialog">
    <div class="modal-dialog">
      <!-- Modal content-->
      <div class="modal-content">
        <div class="modal-header">
          <button type="button" class="close" data-dismiss="modal">&times;</button>
          <h4 class="modal-title"><label>Task Details</label></h4>
        </div>
        <div class="modal-body" id="selectedtaskdetails">
          <p></p>
        </div>
        <div class="modal-footer">
          <button type="button" class="btn btn-default" data-dismiss="modal">Close</button>
        </div>
      </div>
    </div>
  </div>
</div>

<label class="Heading">Action Details</label>
<div class="row clsRowDiv" id="tblForaction">
        <table>
        <thead>
        <tr>
          <th>File Name</th>
          <th>Title</th>
          <th>Comments</th>
          <th>Assignee</th>
          <th>Due Date</th>
          <th>Status</th>
        </tr>
        </thead>
        <tbody id="tbodyForactionDetails">
        <tr>
        <td colspan="6">No Actions</td>
      </tr>
        </tbody>
      </table>
      </div>

      <div class="btnsubmit">
      <input class="submit" type="button" id="btnClose" value="Close">
      </div>
</div>
`;
$(".loader").show();
Itemid = getUrlParameter("Itemid");
getmanagerfromsite();



$(document).on('click','#btnClose',function()
{
     location.href=`${siteURL}/SitePages/JobDetails.aspx`;
});

$(document).on('click','#icon-view',async function()
{
  $(".loader").show();
  var viewdata='';
viewdata=$(this).attr("data-index");
console.log(viewdata);
  await viewtaskdetails(viewdata);
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

async function getIschedulejobList(Itemid)
{
  await sp.web.lists.getByTitle("SiteMasterList").items.select("*").filter("ID eq '"+Itemid+"'").get().then(async (item)=>
  {
    taskdetails=[];
    if(item.length > 0)
    {
        console.log(item);  
        $("#txtNode").val(item[0].NodeID);
        $("#txtSiteName").val(item[0].SiteName);
        $("#txtSiteType").val(item[0].SiteType);
        $("#txtVersion").val(item[0].VersionID);
        $("#txtClient").val(item[0].Client);
        //$("#selectedProjects").val(item[0].Projects);
        ///if(item[0].Projects)
        if(item[0].Category)
        {
            var html='';
            ///tval=item[0].Projects;
            tval=item[0].Category;
            var val=tval.split(";");
            if(val.length>1)
            {
            for(var i=0;i<val.length-1;i++)
            {
                html+="<li>"+val[i]+"</li>";
                await getIscheduletaskList(val[i]);
            }
          } else{
            
                html+="<li>"+tval+"</li>";
                await getIscheduletaskList(val[i]);
          }
              var htmlfortask='';
              var isChecked  ="checked";
              
              for(var i=0;i<taskdetails.length;i++)
              {
                // if(taskdetails[i].Active=="No")
                // isChecked  ="";
                // else
                // isChecked  ="checked";
                if(flagmangerornot||currentuser==taskdetails[i].AssignedToEmail)
                {
                htmlfortask += `<tr><td>${taskdetails[i].Project}</td><td>${taskdetails[i].TaskName}</td><td>${taskdetails[i].AssigneeName}</td><td>${taskdetails[i].DueDate}</td><td><a href="#"><span class="btn btn-info btn-lg" data-toggle="modal" data-target="#myModal" id="icon-view" data-index=${i}></span></a></td></tr>`;
                //<td><input type="checkbox" ${isChecked} class="clsactive" data-index=${i}></td>
              }
              }

          $("#selectedProjects").html('');
          $("#selectedProjects").html(html);

          $("#tbodyForTaskDetails").html('');
          $("#tbodyForTaskDetails").html(htmlfortask);

          if(!htmlfortask)
          $("#tbodyForTaskDetails").html(`<tr><td colspan="5">No Tasks</td></tr>`);

          disableallfields();

          $('.loader').hide();
        }
    }
    else
    {
      ErrorCallBack("No data","getIschedulejobList")
    }
}).catch((error)=>
{
  ErrorCallBack(error, "IscheduleJobList");
});
}

async function getIscheduletaskList(Projects)
{
  await sp.web.lists.getByTitle("IscheduletaskList").items.select("*").filter("Project eq '"+Projects+"' and ReferenceNumber eq '"+Itemid+"' and Active eq '"+YesChecked+"'").get().then(async (item)=>
  {
      if(item.length>0)
      {
          await console.log(item);  
          //taskdetails.push(item);
          for(var i=0;i<item.length;i++)
          {
          taskdetails.push({"Project":item[i].Project,"Priority":item[i].Priority,"TaskName":item[i].TaskName,"AssigneeName":item[i].AssigneeName,"AssignedToEmail":item[i].AssignedToEmail,"DueDate": moment(item[i].DueDate).format("DD-MM-YYYY"),/*"Active":item[i].Active,*/"Startdate": moment(item[i].Startdate).format("DD-MM-YYYY"),"EndDate": moment(item[i].EndDate).format("DD-MM-YYYY"),"HoldStartDate": moment(item[i].HoldStartDate).format("DD-MM-YYYY"),"HoldEndDate": moment(item[i].HoldEndDate).format("DD-MM-YYYY"),"CompletionDate": moment(item[i].CompletionDate).format("DD-MM-YYYY")});
        }
        getJobAction();
      }
      else
      {
        getJobAction();
      }
      
  }).catch((error)=>
  {
    ErrorCallBack(error, "getIscheduletaskList");
  });
}

async function getJobAction()
{
  var htmlforaction='';
  await sp.web.lists.getByTitle("JobAction").items.select("*").filter("ReferenceNumber eq '"+Itemid+"' and Active eq '"+YesChecked+"'").get().then(async (item)=>
  {
      if(item.length>0)
      {
          await console.log(item);  
          //taskdetails.push(item);
          for(var i=0;i<item.length;i++)
          {
            var Refnum=item[i].Id.toString();
            const itemval = await sp.web.lists.getByTitle("JobAction").items.getById(Refnum);
            const Info = await itemval.attachmentFiles();
            console.log(Info); 
            if(flagmangerornot||currentuser==item[i].AssignedToEmail)
            {
            htmlforaction += `<tr><td><a href="${Info[0].ServerRelativeUrl}" target="_blank">${item[i].Filename}</a></td><td>${item[i].Title}</td><td>${item[i].Comments}</td><td>${item[i].AssigneeName}</td><td>${moment(item[i].DueDate).format("DD-MM-YYYY")}</td><td>${item[i].Status}</td></tr>`; 
            }
        }
        $("#tbodyForactionDetails").html('');
        $("#tbodyForactionDetails").html(htmlforaction);
      }
      
  }).catch((error)=>
  {
    ErrorCallBack(error, "getJobAction");
  });
}

function viewtaskdetails(viewdata)
{
  var htmlfortaskdetails='';
  for(var i=0;i<taskdetails.length;i++)
  {
    var Sdate,Edate,HoldSdate,HoldEdate,Completiondate;

    Sdate=taskdetails[viewdata].Startdate;
    if(!Sdate)
    Sdate=taskdetails[viewdata].Startdate;
    else
    Sdate="NA";

    Edate=taskdetails[viewdata].EndDate;
    if(!Edate)
    Edate=taskdetails[viewdata].EndDate;
    else
    Edate="NA";

    HoldSdate=taskdetails[viewdata].HoldStartDate;
    if(!HoldSdate)
    HoldSdate=taskdetails[viewdata].HoldStartDate;
    else
    HoldSdate="NA";
    
    HoldEdate=taskdetails[viewdata].HoldEndDate;
    if(!HoldEdate)
    HoldEdate=taskdetails[viewdata].HoldEndDate;
    else
    HoldEdate="NA";

    Completiondate=taskdetails[viewdata].CompletionDate;
    if(!Sdate)
    Completiondate=taskdetails[viewdata].CompletionDate;
    else
    Completiondate="NA";

    htmlfortaskdetails = `<label>Project Name</label> : <p>${taskdetails[viewdata].Project}</p><label>Task Name</label> : <p>${taskdetails[viewdata].TaskName}</p><label>Assignee Name</label> : <p>${taskdetails[viewdata].AssigneeName}</p><label>Due Date</label> : <p>${taskdetails[viewdata].DueDate}</p><label>Start Date</label> : <p>${Sdate}</p><label>End Date</label> : <p>${Edate}</p><label>Hold Start Date</label> : <p>${HoldSdate}</p><label>Hold End Date</label> : <p>${HoldEdate}</p><label>Completion Date</label> : <p>${Completiondate}</p>`;
    
  } 

$("#selectedtaskdetails").html('');
$("#selectedtaskdetails").html(htmlfortaskdetails);
$('.loader').hide();
}

async function getmanagerfromsite() {
  var ManagerInfo = [];
  await sp.web.siteGroups
    .getByName("ITimeSheetManagers")
    .users.get()
    .then(function (result) {
      for (var i = 0; i < result.length; i++) 
      {
        
          if(result[i].Email == currentuser)
          {
            flagmangerornot=true;
          }
      }
      getIschedulejobList(Itemid);
    })
    .catch(function (err) {
      alert("Group not found: " + err);
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
        $('.loader').hide();
        AlertMessage("Something went wrong.please contact system admin");
      });
  } catch (e) {
    //alert(e.message);
    $('.loader').hide();
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


function disableallfields()
{
  $("#txtNode").prop('disabled',true);
  $("#txtSiteName").prop('disabled',true);
  $("#txtSiteType").prop('disabled',true);
  $("#txtClient").prop('disabled',true);
  $("#txtVersion").prop('disabled',true);
  $(".clsactive").prop('disabled',true);
}
