import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './EditJobDetailsWebPart.module.scss';
import * as strings from 'EditJobDetailsWebPartStrings';
import { SPComponentLoader } from "@microsoft/sp-loader";
import { MSGraphClient } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

import "jquery";
import * as moment from "moment";
import "datatables";
import { sp } from "@pnp/pnpjs";


import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";

// import "../../ExternalRef/js/sp.peoplepicker.js";
SPComponentLoader.loadCss("https://cdn.jsdelivr.net/npm/select2@4.1.0-beta.1/dist/css/select2.min.css");
SPComponentLoader.loadCss(
  "https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css"
);
import "../../ExternalRef/css/alertify.min.css";
import "../../ExternalRef/css/StyleJob.css";
import "../../ExternalRef/css/loader.css";
import "../../ExternalRef/js/select2.min.js";
var alertify: any = require("../../ExternalRef/js/alertify.min.js");
declare var $;

var siteURL="";
var that;
var options='';
var Itemid;
var taskdetails=[];
var tval='';

export interface IEditJobDetailsWebPartProps {
  description: string;
}

export default class EditJobDetailsWebPart extends BaseClientSideWebPart <IEditJobDetailsWebPartProps> {

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
    <th>Active</th>
  </tr>
  </thead>
  <tbody id="tbodyForTaskDetails">
  <tr>
    <td> </td>
  </tr>
  </tbody>
</table>
</div>

<label class="Heading Actiondetails"">Action Details</label>
        <div class="row clsRowDiv divforaction Actiondetails">
        <div class="column col-xl-4 col-lg-4 col-md-12 col-sm-12 col-12" class="fileupload">
            <label>Upload File</label>
            <input type="file" id="fileupload"> 
          </div>
          <div class="column col-xl-3 col-lg-3 col-md-12 col-sm-12 col-12">
            <label>Comments</label>
            <input type="text" id="txtcmd">
          </div>
          <div class="column col-xl-3 col-lg-3 col-md-12 col-sm-12 col-12">
            <label>Assignee</label>
            <select id="actionassignee"></select>
          </div>
          <div class="column col-xl-3 col-lg-3 col-md-12 col-sm-12 col-12">
            <label>Due Date</label>
            <input type="date" id="datedue">
          </div>
          <div class="column col-xl-3 col-lg-3 col-md-12 col-sm-12 col-12">
          <input class="btnsave" type="button" id="btnsave" value="Save">
          </div>
          </div> 

<div class="row clsRowDiv" id="tblForaction">
        <table>
        <thead>
        <tr>
          <th>File Name</th>
          <th>Comments</th>
          <th>Assignee</th>
          <th>Due Date</th>
          <th>Status</th>
        </tr>
        </thead>
        <tbody id="tbodyForactionDetails">
        <tr>
          <td> </td>
        </tr>
        </tbody>
      </table>
      </div>

<div class="btnsubmit">
<input class="submit" type="button" id="btnUpdate" value="Update">
<input class="submit" type="button" id="btnClose" value="Close">
</div>
</div>
`;
$(".loader").show();
Itemid = getUrlParameter("Itemid");


$(document).on('click','#btnClose',function()
{
     location.href=`${siteURL}/SitePages/JobDetails.aspx`;
});

$(document).on('click','#btnUpdate',function()
{
  $(".loader").show();
  UpdateJobDetails();
});

getusersfromsite();
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

async function getusersfromsite()
{
  await that.context.msGraphClientFactory.getClient().then( async client => {
     client.api("/users").select("*").get(async ( error, response) => 
    {
    var getResponse = response;
    console.log("My Details: ");
    console.log(response);
    for(var i=0;i<response.value.length;i++)
  {
    if(response.value[i].mail) 
    options += `<option data-name="${response.value[i].displayName}" value="${response.value[i].mail}">${response.value[i].displayName}</option>`;
  }
  await console.log(options);
  await getIschedulejobList(Itemid); 
    });
  }).catch(function (error) {
    ErrorCallBack(error, "getusersfromsite");
  });
}

async function getIschedulejobList(Itemid)
{
  $(".loader").show();
  await sp.web.lists.getByTitle("IscheduleJobList").items.select("*").filter("ID eq '"+Itemid+"'").get().then(async (item)=>
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
        if(item[0].Projects)
        {
            var html='';
            tval=item[0].Projects;
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
                await getIscheduletaskList(tval);
          }
              var htmlfortask='';
              var isChecked  ="checked";
              
              for(var i=0;i<taskdetails.length;i++)
              {
                if(taskdetails[i].Active=="No")
                isChecked  ="";
                else
                isChecked  ="checked";

                htmlfortask += `<tr><td>${taskdetails[i].Project}</td><td>${taskdetails[i].TaskName}</td><td><select class="clsassign" data-index=${i}>${options}</select></td><td><input type="date" class="clsduedate" data-index=${i}></td><td><input type="checkbox" ${isChecked} class="clsactive" data-index=${i}></td></tr>`;
                
              }

          $("#selectedProjects").html('');
          $("#selectedProjects").html(html);

          $("#tbodyForTaskDetails").html('');
          $("#tbodyForTaskDetails").html(htmlfortask);

          disableallfields();
          
          $('.clsassign').each(function()
          {
            $(this).val(taskdetails[$(this).attr('data-index')].AssignedToEmail);
         });

         $('.clsduedate').each(function()
         {
           var Date=taskdetails[$(this).attr('data-index')].DueDate;
          $(this).val(Date);
        });

         $(".clsassign").select2();

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
  $(".loader").show();
  await sp.web.lists.getByTitle("IscheduletaskList").items.select("*").filter("Project eq '"+Projects+"' and ReferenceNumber eq '"+Itemid+"'").get().then(async (item)=>
  {
      if(item.length>0)
      {
          await console.log(item);  
          //taskdetails.push(item);
          for(var i=0;i<item.length;i++)
          {
          taskdetails.push({"Project":item[i].Project,"Priority":item[i].Priority,"ID":item[i].ID,"TaskName":item[i].TaskName,"AssigneeName":item[i].AssigneeName,"AssignedToEmail":item[i].AssignedToEmail,"DueDate": moment(item[i].DueDate).format("YYYY-MM-DD"),"Active":item[i].Active});
        }
        getJobAction();
      }
      
  }).catch((error)=>
  {
    ErrorCallBack(error, "IscheduletaskList");
  });
}

async function getJobAction()
{
  var htmlforaction='';
  await sp.web.lists.getByTitle("JobAction").items.select("*").filter("ReferenceNumber eq '"+Itemid+"'").get().then(async (item)=>
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
            htmlforaction += `<tr><td><a href="${Info[0].ServerRelativeUrl}" target="_blank">${item[i].Filename}</a></td><td>${item[i].Comments}</td><td>${item[i].AssigneeName}</td><td>${moment(item[i].DueDate).format("DD-MM-YYYY")}</td><td>${item[i].Status}</td></tr>`; 
        }
        $("#tbodyForactionDetails").html('');
        $("#tbodyForactionDetails").html(htmlforaction);
      }
      
  }).catch((error)=>
  {
    ErrorCallBack(error, "insert getJobAction");
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
}

   async function UpdateJobDetails() {

  $('.clsduedate').each(function()
  {
  if($(this).val()!="")
  {
  taskdetails[$(this).attr('data-index')].DueDate=$(this).val();
  }
  });
  $('.clsactive').each(function()
  {
  taskdetails[$(this).attr('data-index')].Active=($(this).is(':checked')? "Yes" : "No");
  });
  $('.clsassign').each(function()
  {
      taskdetails[$(this).attr('data-index')].AssignedToEmail=$(this).val();
      taskdetails[$(this).attr('data-index')].AssigneeName=$(this).find("option:selected").attr("data-name");
  });

    var count=1;
    var requesttaskdata = {};
                for(var i=0;i<taskdetails.length;i++)
                {
                  var Id=taskdetails[i].ID;
                  requesttaskdata = {
                      DueDate: taskdetails[i].DueDate,
                      Active: taskdetails[i].Active,
                      AssignedToEmail: taskdetails[i].AssignedToEmail,
                      AssigneeName: taskdetails[i].AssigneeName
                    };
                    await sp.web.lists
                      .getByTitle("ischeduletasklist")
                       .items.getById(Id)
                       .update(requesttaskdata).then(function (data) {
                      count++;
  
                      if(count==taskdetails.length)
                      {
                        $(".loader").hide();
                        AlertMessage("Job Updated successfully");
                      }
                      
                    })
                    .catch(function (error) {
                      ErrorCallBack(error, "insert ischeduletasklist");
                    });
                    
                  }
                
  }

  
