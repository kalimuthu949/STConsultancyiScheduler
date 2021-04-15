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
import "../../ExternalRef/css/alertify.min.css";
var alertify: any = require("../../ExternalRef/js/alertify.min.js");
declare var $;

var that;
var Itemid;
var taskdetails=[];
var tval='';

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
<div class="btnsubmit">
<input class="submit" type="button" id="btnClose" value="Close">
</div>
</div>
`;
$(".loader").show();
Itemid = getUrlParameter("Itemid");
getIschedulejobList(Itemid);

$(document).on('click','#btnClose',function()
{
     location.href=`${siteURL}/SitePages/JobDetails.aspx`;
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

                htmlfortask += `<tr><td>${taskdetails[i].Project}</td><td>${taskdetails[i].TaskName}</td><td>${taskdetails[i].AssigneeName}</td><td>${taskdetails[i].DueDate}</td><td><input type="checkbox" ${isChecked} class="clsactive" data-index=${i}></td></tr>`;
                
              }
          $("#selectedProjects").html('');
          $("#selectedProjects").html(html);

          $("#tbodyForTaskDetails").html('');
          $("#tbodyForTaskDetails").html(htmlfortask);

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
  
  await sp.web.lists.getByTitle("IscheduletaskList").items.select("*").filter("Project eq '"+Projects+"'").get().then(async (item)=>
  {
      if(item.length>0)
      {
          await console.log(item);  
          //taskdetails.push(item);
          for(var i=0;i<item.length;i++)
          {
          taskdetails.push({"Project":item[i].Project,"Priority":item[i].Priority,"TaskName":item[i].TaskName,"AssigneeName":item[i].AssigneeName,"DueDate": moment(item[i].DueDate).format("DD-MM-YYYY"),"Active":item[i].Active});
        }
      }
      
  }).catch((error)=>
  {
    ErrorCallBack(error, "IscheduletaskList");
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