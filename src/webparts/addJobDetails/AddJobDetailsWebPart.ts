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
var actiondetails=[];
var htmlfordate='';
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
<div class="container ischedule" ><label class="Heading">Ischedule</label>
        <div class="row clsRowDiv">
          <div class="column col-xl-3 col-lg-3 col-md-12 col-sm-12 col-12">
            <label>Node ID</label>
            <input type="text" id="clkNode"> 
            <div class ="generate-fields" id="generateFields"></div>
          </div>
          <div class="column col-xl-6 col-lg-6 col-md-12 col-sm-12 col-12">
            <label>Job Detail</label>
            <select id="drpCategory" class="drpCategory">
            </select>
          </div>
          </div>
          </div>

    <div class="container ischedule"><label class="Heading divsitedetails" style="display:none">Site Details</label>
        <div class="row clsRowDiv divsitedetails" style="display:none">
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
    <th>Action</th>
  </tr>
  </thead>
  <tbody id="tbodyForTaskDetails">
  <tr>
    <td> </td>
  </tr>
  </tbody>
</table>
</div>

<label class="Heading Actiondetails" style="display:none">Action Details</label>
        <div class="row clsRowDiv divforaction Actiondetails" style="display:none">
        <div class="column col-xl-4 col-lg-4 col-md-12 col-sm-12 col-12" class="fileupload">


        <label for="file-upload" class="custom-file-upload">
        <i class="fa fa-cloud-upload"></i> Upload File
      </label>
      <input id="file-upload" name='upload_cont_img' type="file" style="display:none;">

          </div>
          <div class="column col-xl-3 col-lg-3 col-md-12 col-sm-12 col-12">
            <label>Title</label>
            <input type="text" id="txttitle">
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
          <input type="button" class="btnUpdate" id="btnUpdate" data-id="" value="Update" style="display:none">
          </div>
          </div> 

        <div class="row clsRowDiv" id="tblForaction" style="display:none">
        <table>
        <thead>
        <tr>
          <th>File Name</th>
          <th>Title</th>
          <th>Comments</th>
          <th>Assignee</th>
          <th>Due Date</th>
          <th>Edit</th>
        </tr>
        </thead>
        <tbody id="tbodyForactionDetails">
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


$("#generateFields").click(async function()
{
  $(".loader").show();
  await getIscheduledetails($("#clkNode").val());
});

$("#drpCategory").change(async function()
{
var Siteid='';
Siteid=$("#drpCategory option:selected").attr("data-id");
await getSiteDetails(Siteid);
});

$("#btnsubmit").click(async function()
{
  $(".loader").show();
  await insertischedulejoblist();
});

$("#btnsave").click(async function()
{
  
  if(mandatoryforaddaction())
  {
    $(".loader").show();
    await getactiondetails(true);
  }
  else
  {
    console.log("All fileds not filled");
    $(".loader").hide()
  }

});

$('#file-upload').change(function() {
  var i = $(this).prev('label').clone();
  var file = $('#file-upload')[0].files[0].name;
  $(this).prev('label').text(file);
});

$(document).on('click','.icon-edit',async function()
{
  //$(".loader").show();

  $("#btnsave").hide();
  $("#btnUpdate").show();
  $("#btnUpdate").attr("data-id",$(this).attr('data-index'))
var editdata='';
editdata=$(this).attr("data-index");
console.log(editdata);
  await editactiondetails(editdata);
});

$("#btnUpdate").click(async function()
{
    if($(this).attr('data-id'))
    {
        
        if(!$("#txtcmd").val())
        {
          alertify.error("Please Enter Comments");
          return false;
        }
      
        if($("#file-upload")[0].files.length>0)
        {
          actiondetails[$(this).attr('data-id')].Filename=$("#file-upload")[0].files[0].name;
          actiondetails[$(this).attr('data-id')].FileContent=$("#file-upload")[0].files[0];
        }

        actiondetails[$(this).attr('data-id')].Comments=$("#txtcmd").val();
        actiondetails[$(this).attr('data-id')].AssignedToEmail=$("#actionassignee option:selected").val();
        actiondetails[$(this).attr('data-id')].AssigneeName=$("#actionassignee option:selected").attr("data-name");
        actiondetails[$(this).attr('data-id')].DueDate=$("#datedue").val();

        
  $("#btnsave").show();
  $("#btnUpdate").hide();
  $("#btnUpdate").attr('data-id','');

  await getactiondetails(false);

    }
    else
    {

    }
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
async function getIscheduledetails(NodeID)
{
  await sp.web.lists.getByTitle("SiteMasterList").items.select("*").filter("NodeID eq '"+NodeID+"'").get().then(async (item)=>
  { 
    var htmlforCategory="";
    if(item.length > 0)
    {
      htmlforCategory=`<option>Select</option>`;
    for(var i=0;i<item.length;i++)
    {
    
        htmlforCategory += `<option data-name="${item[i].Category}" data-id="${item[i].ID}">${item[i].NodeID} - Version${item[i].VersionID} (${(item[i].Category)})</option>`;
    }
  }
  $("#drpCategory").html('');
  $("#drpCategory").html(htmlforCategory);
  $(".loader").hide()
  }).catch((error)=>
  {
    ErrorCallBack(error, "getSiteDetails");
  });
}

async function getSiteDetails(SiteId)
{
  await sp.web.lists.getByTitle("SiteMasterList").items.select("*").filter("ID eq '"+SiteId+"'").get().then(async (item)=>
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
              
                  htmlfortask += `<tr><td>${taskdetails[i].Projects}</td><td>${taskdetails[i].Tasks}</td><td><select class="clsassign" data-index=${i}>${options}</select></td><td><input type="date" class="clsduedate" value="${moment().format("YYYY-MM-DD")}" data-index=${i}></td><td><input type="checkbox" checked class="clsactive" data-index=${i}></td></tr>`;
                
              }
              $(".divsitedetails").show();
              $("#selectedProjects").html('');
              $("#selectedProjects").html(html);

              $(".divProjectdetails").show();
              $("#tblForTasks").show();
              $("#btnsubmit").show();
              $(".Actiondetails").show();

              $("#tbodyForTaskDetails").html('');
              $("#tbodyForTaskDetails").html(htmlfortask);
              disableallfields();

              $("#actionassignee").html('');
              $("#actionassignee").html(options);

              $("#datedue").val('');
              $("#datedue").val(moment().format("YYYY-MM-DD"));

              $('.loader').hide();

              $(".clsassign").select2();
              $("#actionassignee").select2();
              
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
          taskdetails.push({"ID":item[i].ID,"Division":item[i].Division,"BusinessDivision":item[i].BusinessDivision,"Projects":item[i].Projects,"Priority":item[i].Priority,"Tasks":item[i].Tasks,"Assignee":"","AssigneeName":"","DueDate": null,"Active":""});
        }
      }
      
  }).catch((error)=>
  {
    ErrorCallBack(error, "getTaskDetails");
  });
}

//insert work
async function insertischedulejoblist() {

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
      Projects: projects,

    };

    await sp.web.lists
    .getByTitle("ischedulejoblist")
    .items.add(requestdata)
    .then(async function (data) {
      console.log(data);
      var strRefnumber=data.data.Id.toString();
      await insertischeduletasklist(strRefnumber);
      //AlertMessage("Record created successfully");
    })
    .catch(function (error) {
      ErrorCallBack(error, "insertischedulejoblist");
    });
    
}

async function insertischeduletasklist(RefNum) {
  var count=0;

  var requesttaskdata = {};
              for(var i=0;i<taskdetails.length;i++)
              {
                var Percentage="0";
                var BalancePercent="100";
                var taskid=taskdetails[i].ID.toString();
                requesttaskdata = {

                    ReferenceNumber :RefNum,
                    Project: taskdetails[i].Projects,
                    TaskName: taskdetails[i].Tasks,
                    Priority: taskdetails[i].Priority,
                    DueDate: taskdetails[i].DueDate,
                    Active: taskdetails[i].Active,
                    AssignedToEmail: taskdetails[i].Assignee,
                    AssigneeName: taskdetails[i].AssigneeName,
                    TaskID: taskid,
                    Division: taskdetails[i].Division,
                    BusinessDivision: taskdetails[i].BusinessDivision,
                    Percentage:Percentage,
                    BalancePercent:BalancePercent
                  };
                  await sp.web.lists
                  .getByTitle("ischeduletasklist")
                  .items.add(requesttaskdata)
                  .then(function (data) 
                  {
                    count++;

                    if(count==taskdetails.length)
                    {
                      
                      if(actiondetails.length>0)
                      Insertactiondetails(RefNum);
                      else
                      {
                        $(".loader").hide();
                        AlertMessage("Record created successfully");
                      }
                      //AlertMessage("Record created successfully");
                    }
                    
                  })
                  .catch(function (error) {
                    ErrorCallBack(error, "insertischeduletasklist");
                  });
                  
                }
              
}

async function getactiondetails(flagforupdate)
{
  $("#tblForaction").show();
  var requestactiondata = {}; 

  if(flagforupdate)
  {
   requestactiondata = {
    //ReferenceNumber :strRefnumber,
    Filename:$("#file-upload")[0].files[0].name,
    FileContent:$("#file-upload")[0].files[0],
    Title:$("#txttitle").val(),
    Comments:$("#txtcmd").val(),
    AssignedToEmail:$("#actionassignee option:selected").val(),
    AssigneeName:$("#actionassignee option:selected").attr("data-name"),
    DueDate:$("#datedue").val()
  }
  actiondetails.push(requestactiondata);
}
  console.log(actiondetails);

  $('#txtcmd').val("");
  $('#txttitle').val("");
  $('#file-upload').val("");
  $(".custom-file-upload").text('Upload File');
  $('#datedue').val(moment().format("YYYY-MM-DD"));

  var htmlforaction="";
  for(var i=0;i<actiondetails.length;i++)
              {
    var Ddate;
    Ddate=actiondetails[i].DueDate;
    if(!Ddate)
    Ddate=actiondetails[i].DueDate;
    else
    Ddate="NA";
              
                  htmlforaction += `<tr><td>${actiondetails[i].Filename}</td><td>${actiondetails[i].Title}</td><td>${actiondetails[i].Comments}</td><td>${actiondetails[i].AssigneeName}</td><td>${moment(actiondetails[i].DueDate).format("DD-MM-YYYY")}</td><td><a href="#"><span class="icon-edit" data-index=${i}></span></a></td></tr>`; 
              }

              $("#tbodyForactionDetails").html('');
              $("#tbodyForactionDetails").html(htmlforaction);
              $(".loader").hide();
}

async function Insertactiondetails(RefNum)
              {
            var count=0;
                for(var i=0;i<actiondetails.length;i++)
            {
              await sp.web.lists
              .getByTitle("JobAction")
              .items.add({"ReferenceNumber":RefNum,"Filename":actiondetails[i].Filename,"Title":actiondetails[i].Title,"Comments":actiondetails[i].Comments,"AssignedToEmail":actiondetails[i].AssignedToEmail,"AssigneeName":actiondetails[i].AssigneeName,"DueDate":actiondetails[i].DueDate})
              .then(async function (data) 
              {
                count++;
                var Refnum=data.data.Id.toString();
                const Item = await sp.web.lists.getByTitle("JobAction").items.getById(Refnum);
                await Item.attachmentFiles.add(actiondetails[i].Filename, actiondetails[i].FileContent);
                if(count==actiondetails.length)
                {
                  $(".loader").hide();
                  AlertMessage("Record created successfully");
                }
                
              })
              .catch(function (error) {
                ErrorCallBack(error, "getactiondetails");
              });
            }
              }

function editactiondetails(editdata)
{

  //$("#file-upload").files[0].name.val(actiondetails[editdata].FileContent);
  $('.custom-file-upload').text(actiondetails[editdata].Filename);
  $("#txttitle").val(actiondetails[editdata].Title);
  $("#txtcmd").val(actiondetails[editdata].Comments);
  //$("#actionassignee").val(actiondetails[editdata].AssigneeName);
  $("#actionassignee").val(actiondetails[editdata].AssignedToEmail);
  $("#datedue").val(actiondetails[editdata].DueDate);
  $("#actionassignee").select2();

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

function disableallfields()
{
  $("#txtNode").prop('disabled',true);
  $("#txtSiteName").prop('disabled',true);
  $("#txtSiteType").prop('disabled',true);
  $("#txtClient").prop('disabled',true);
  $("#txtVersion").prop('disabled',true);
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

function mandatoryforaddaction()
{
      var isAllvalueFilled=true;

      if($("#file-upload")[0].files.length==0)
      {
        alertify.error("Please Select file");
        isAllvalueFilled=false;
        
      }
      else if(!$("#txtcmd").val())
      {
        alertify.error("Please enter comments");
        isAllvalueFilled=false;
        
      }
      else if(!$("#txttitle").val())
      {
        alertify.error("Please enter title");
        isAllvalueFilled=false;
        
      }

      return isAllvalueFilled;
}

