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
var actiondetails=[];
var FilteredManager =[];
var currentuser = "";
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
    currentuser = this.context.pageContext.user.email;
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
    <th>Action</th>
  </tr>
  </thead>
  <tbody id="tbodyForTaskDetails">
  <tr>
  <td colspan="5">No Tasks</td>
  </tr>
  </tbody>
</table>
</div>

<label class="Heading Actiondetails"">Action Details</label>
        <div class="row clsRowDiv divforaction Actiondetails">
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
          <th>Action</th>
          <th>Edit</th>
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
<input class="submit" type="button" id="MbtnUpdate" value="Update">
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

$(document).on('click','#MbtnUpdate',function()
{
  $(".loader").show();
  UpdateIscheduleLists();
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
        actiondetails[$(this).attr('data-id')].Title=$("#txttitle").val();
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

getmanagerfromsite();
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
          getusersfromsite();
        } 
        else{
          ErrorCallBack("Data Not Found","getmanagerfromsite")
        }
    })
    .catch(function (err) {
      alert("Group not found: " + err);
    });
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

              $("#actionassignee").html('');
              $("#actionassignee").html(options);

              $("#datedue").val('');
              $("#datedue").val(moment().format("YYYY-MM-DD"));

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
         $("#actionassignee").select2();

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
    ErrorCallBack(error, "getIscheduletaskList");
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
            
            const itemval = await sp.web.lists.getByTitle("JobAction").items.getById(Refnum).attachmentFiles().then(async function(data)
            {

              var fileURL="";

              for(var j=0;j<data.length;j++)
              {
                fileURL=data[j].ServerRelativeUrl;
              }
              
              var requestactiondata = 
              {
                Filename:item[i].Filename,
                FileContent:"",
                Title:item[i].Title,
                Comments:item[i].Comments,
                AssignedToEmail:item[i].AssignedToEmail,
                AssigneeName:item[i].AssigneeName,
                DueDate:moment(item[i].DueDate).format("YYYY-MM-DD"),
                Active:item[i].Active,
                ID:Refnum,
                FileURL:fileURL,
                Status:item[i].Status
              }
              var isChecked  ="checked";
                if(item[i].Active=="No")
                isChecked  ="";
                else
                isChecked  ="checked";

              htmlforaction += `<tr><td><a href="${fileURL}" target="_blank">${item[i].Filename}</a></td><td>${item[i].Title}</td><td>${item[i].Comments}</td><td>${item[i].AssigneeName}</td><td>${moment(item[i].DueDate).format("DD-MM-YYYY")}</td><td>${item[i].Status}</td><td><input type="checkbox" ${isChecked} class="clsaction" data-index=${i}></td><td><a href="#"><span class="icon-edit" data-index=${i}></span></a></td></tr>`;
  
              await actiondetails.push(requestactiondata);
            }).catch((error)=>
            {
              ErrorCallBack(error, "getJobAction");
            });

        }
        
        $("#tbodyForactionDetails").html('');
        $("#tbodyForactionDetails").html(htmlforaction);

        var result = [];
        $.each(actiondetails, function (i, e) {
            var matchingItems = $.grep(result, function (item) {
              return item.ID === e.ID;
            });
            if (matchingItems.length === 0){
                result.push(e);
            }
        });

        actiondetails=result;
      }
      
  }).catch((error)=>
  {
    ErrorCallBack(error, "insert getJobAction");
  });
}

function editactiondetails(editdata)
{
  $('.custom-file-upload').text(actiondetails[editdata].Filename);
  $("#txttitle").val(actiondetails[editdata].Title);
  $("#txtcmd").val(actiondetails[editdata].Comments);
  //$("#actionassignee").val(actiondetails[editdata].AssigneeName);
  $("#actionassignee").val(actiondetails[editdata].AssignedToEmail);
  $("#datedue").val(actiondetails[editdata].DueDate);
  $("#actionassignee").select2();
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
    DueDate:$("#datedue").val(),
    Active:$("#clsaction").val(),
    ID:"",
    FileURL:"",
    Status:"N/A"
  } 
  actiondetails.push(requestactiondata);
}

  console.log(actiondetails);

  $('#txttitle').val("");
  $('#txtcmd').val("");
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
    var isChecked  ="checked";
    if(actiondetails[i].Active=="No")
    isChecked  ="";
    else
    isChecked  ="checked";
                  htmlforaction += `<tr><td>${actiondetails[i].Filename}</td><td>${actiondetails[i].Title}</td><td>${actiondetails[i].Comments}</td><td>${actiondetails[i].AssigneeName}</td><td>${moment(actiondetails[i].DueDate).format("DD-MM-YYYY")}</td><td>${actiondetails[i].Status}</td><td><input type="checkbox" ${isChecked} class="clsaction" data-index=${i}></td><td><a href="#"><span class="icon-edit" data-index=${i}></span></a></td></tr>`; 
              }

              $("#tbodyForactionDetails").html('');
              $("#tbodyForactionDetails").html(htmlforaction);
              $(".loader").hide();
}

async function Insertactiondetails(RefNum)
{
  $('.clsaction').each(function()
  {
    actiondetails[$(this).attr('data-index')].Active=($(this).is(':checked')? "Yes" : "No");
  });
            var count=0;
            for(var i=0;i<actiondetails.length;i++)
            {
              
              if(actiondetails[i].ID!="")
              {
                await sp.web.lists
                .getByTitle("JobAction")
                .items.getById(actiondetails[i].ID).update({"ReferenceNumber":RefNum,"Filename":actiondetails[i].Filename,"Title":actiondetails[i].Title,"Comments":actiondetails[i].Comments,"AssignedToEmail":actiondetails[i].AssignedToEmail,"AssigneeName":actiondetails[i].AssigneeName,"DueDate":actiondetails[i].DueDate,"Active":actiondetails[i].Active})
                .then(async function (data) 
                {
                  count++;
                  var Refnum=actiondetails[i].ID;
  
                  if(actiondetails[i].FileContent!="")
                  {
                    const Item = await sp.web.lists.getByTitle("JobAction").items.getById(Refnum);
                    await Item.attachmentFiles.add(actiondetails[i].Filename, actiondetails[i].FileContent);
                  }              
                  if(count==actiondetails.length)
                  {
                    $(".loader").hide();
                    AlertMessage("Record Updated successfully");
                  }
                  
                })
                .catch(function (error) {
                  ErrorCallBack(error, "Insertactiondetails");
                });
              }
              else
              {
                await sp.web.lists
                .getByTitle("JobAction")
                .items.add({"ReferenceNumber":RefNum,"Filename":actiondetails[i].Filename,"Title":actiondetails[i].Title,"Comments":actiondetails[i].Comments,"AssignedToEmail":actiondetails[i].AssignedToEmail,"AssigneeName":actiondetails[i].AssigneeName,"DueDate":actiondetails[i].DueDate,"Active":actiondetails[i].Active})
                .then(async function (data) 
                {
                  count++;
                  var Refnum=data.data.Id.toString();
  
                  if(actiondetails[i].FileContent!="")
                  {
                    const Item = await sp.web.lists.getByTitle("JobAction").items.getById(Refnum);
                    await Item.attachmentFiles.add(actiondetails[i].Filename, actiondetails[i].FileContent);
                  }              
                  if(count==actiondetails.length)
                  {
                    $(".loader").hide();
                    AlertMessage("Record Updated successfully");
                  }
                  
                })
                .catch(function (error) {
                  ErrorCallBack(error, "actionlistdetails");
                });
              }

            }
}

async function UpdateIscheduleLists() {

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
                        
                        //AlertMessage("Job Updated successfully");
                        if(actiondetails.length>0)
                        Insertactiondetails(Itemid);
                        else{
                          $(".loader").hide();
                          AlertMessage("Record Updated successfully");
                        }
                        
                      }
                      
                    })
                    .catch(function (error) {
                      ErrorCallBack(error, "UpdateIscheduleLists");
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
        window.location.href = siteURL+"/SitePages/JobDetails.aspx";
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
        alertify.error("Please enter comments");
        isAllvalueFilled=false;
        
      }

      return isAllvalueFilled;
}