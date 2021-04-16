//import { Log, Version } from '@microsoft/sp-core-library';
import { UrlQueryParameterCollection, Log,Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './AddJobDetailsWebPart.module.scss';
import * as strings from 'AddJobDetailsWebPartStrings';
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

// import "../../ExternalRef/js/sp.peoplepicker.js";
var taskdetails=[];
var projects='';
var options='';
var that;
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
    that=this;
    siteURL=this.context.pageContext.web.absoluteUrl;
    this.domElement.innerHTML = `
    <span style="display:none" class="loader">
<img class="loader-spin"/>
</span>
    <div class="container ischedule" ><label class="Heading">Site Details</label>
        <div class="row clsRowDiv">
          <div class="column col-xl-3 col-lg-3 col-md-12 col-sm-12 col-12">
            <label>Node ID</label>
            <input type="text" id="txtNode"> 
            <div class ="generate-fields" id="generateFields"></div>
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
<div class="btnsubmit"><input class="submit" type="button" id="btnsubmit" value="Submit" style="display:none">
<input class="submit" type="button" id="btnClose" value="Close">
</div>
</div>
`;

$(".loader").show();

$(document).on('click','#btnClose',function()
{
     location.href=`${siteURL}/SitePages/JobDetails.aspx`;
});


getusersfromsite();
// $("#txtNode").blur(function()
// {
//   getSiteDetails($("#txtNode").val());
// });

$("#generateFields").click(async function()
{
  $(".loader").show();
  await getSiteDetails($("#txtNode").val());
});

$("#btnsubmit").click(async function()
{
  $(".loader").show();
  await insertischedulejoblist();
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

function test()
{
}

async function getSiteDetails(NodeID)
{
  await sp.web.lists.getByTitle("SiteMasterList").items.select("*").filter("NodeID eq '"+NodeID+"'").get().then(async (item)=>
  {
    taskdetails=[];
    
      if(item.length > 0)
      {
          // console.log(item);  
          $("#txtNode").val(item[0].NodeID);
          $("#txtSiteName").val(item[0].SiteName);
          $("#txtSiteType").val(item[0].SiteType);
          $("#txtVersion").val(item[0].VersionID);
          $("#txtClient").val(item[0].Client);
          // $("#selectedProjects").val(item[0].Category);
          //console.log(item[0].Category);
          if(item[0].Category)
          {
              var html='';
              projects=item[0].Category;
              var val=projects.split(";");
              if(val.length>1)
              {
              for(var i=0;i<val.length-1;i++)
              {
                  html+="<li>"+val[i]+"</li>";
                  await getTaskDetails(val[i]);
              }
            } else{
              
                  html+="<li>"+projects+"</li>";
                  await getTaskDetails(projects);
              
            }

              console.log(taskdetails);
              var htmlfortask='';
            
              for(var i=0;i<taskdetails.length;i++)
              {
              
                  htmlfortask += `<tr><td>${taskdetails[i].Projects}</td><td>${taskdetails[i].Tasks}</td><td><select class="clsassign" data-index=${i}>${options}</select></td><td><input type="date" class="clsduedate" data-index=${i}></td><td><input type="checkbox" checked class="clsactive" data-index=${i}></td></tr>`;
                
              }
              
              $("#selectedProjects").html('');
              $("#selectedProjects").html(html);

              $(".divProjectdetails").show();
              $("#tblForTasks").show();
              $("#btnsubmit").show();

              $("#tbodyForTaskDetails").html('');
              $("#tbodyForTaskDetails").html(htmlfortask);

              $('.loader').hide();

              $(".clsassign").select2();
              
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

async function getTaskDetails(Projects)
{
  
  await sp.web.lists.getByTitle("TaskMasterList").items.select("*").filter("Projects eq '"+Projects+"'").get().then(async (item)=>
  {
      if(item.length>0)
      {
          await console.log(item);  
          //taskdetails.push(item);
          for(var i=0;i<item.length;i++)
          {
          taskdetails.push({"Projects":item[i].Projects,"Priority":item[i].Priority,"Tasks":item[i].Tasks,"Assignee":"","AssigneeName":"","DueDate": null,"Active":""});
        }
      }
      
  }).catch((error)=>
  {
    ErrorCallBack(error, "getSiteDetails");
  });
}

//insert work
async function insertischedulejoblist() {

  // $('.clsassign').each(function()
  // {
  // taskdetails[$(this).attr('data-index')].Assignee=$(this).val();
  
  // });
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
      taskdetails[$(this).attr('data-index')].Assignee=$(this).val();
      taskdetails[$(this).attr('data-index')].AssigneeName=$(this).find("option:selected").attr("data-name");
  });

  

    var requestdata = {}; 
    requestdata = {
      NodeID: $("#txtNode").val(),
      SiteName: $("#txtSiteName").val(),
      SiteType: $("#txtSiteType").val(),
      Client: $("#txtClient").val(),
      VersionID: $("#txtVersion").val(),
      Projects: projects
    };

    await sp.web.lists
    .getByTitle("ischedulejoblist")
    .items.add(requestdata)
    .then(function (data) {
      console.log(data);
      var strRefnumber=data.data.Id.toString();
      insertischeduletasklist(strRefnumber);
      //AlertMessage("Record created successfully");
    })
    .catch(function (error) {
      ErrorCallBack(error, "insert ischedulejoblist");
    });
    
}

async function insertischeduletasklist(RefNum) {
  var count=1;
  var requesttaskdata = {};
              for(var i=0;i<taskdetails.length;i++)
              {
                
                requesttaskdata = {

                    ReferenceNumber :RefNum,
                    Project: taskdetails[i].Projects,
                    TaskName: taskdetails[i].Tasks,
                    Priority: taskdetails[i].Priority,
                    DueDate: taskdetails[i].DueDate,
                    Active: taskdetails[i].Active,
                    AssignedToEmail: taskdetails[i].Assignee,
                    AssigneeName: taskdetails[i].AssigneeName
                  };
                  await sp.web.lists
                  .getByTitle("ischeduletasklist")
                  .items.add(requesttaskdata)
                  .then(function (data) 
                  {
                    count++;

                    if(count=taskdetails.length)
                    {
                      $(".loader").hide();
                      AlertMessage("Job created successfully");
                    }
                    
                  })
                  .catch(function (error) {
                    ErrorCallBack(error, "insert ischeduletasklist");
                  });
                  
                }
              
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
  $('.loader').hide();
    });
  }).catch(function (error) {
    ErrorCallBack(error, "getusersfromsite");
  });
}

