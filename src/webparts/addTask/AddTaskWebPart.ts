import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './AddTaskWebPart.module.scss';
import * as strings from 'AddTaskWebPartStrings';
import { SPComponentLoader } from '@microsoft/sp-loader';


import * as $ from "jquery"; 
import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import * as moment from 'moment';  
import { readItems,readItem, addItems,formatDate } from "../../commonJS";
SPComponentLoader.loadCss('//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css')
SPComponentLoader.loadScript('//cdn.jsdelivr.net/npm/alertifyjs@1.13.1/build/alertify.min.js');
SPComponentLoader.loadScript('https://code.jquery.com/ui/1.12.1/jquery-ui.js');
import "../../ExternalRef/css/StyleSchedule.css";
import 'alertifyjs';
import "../../ExternalRef/css/alertify.min.css";
var alertify: any = require("../../ExternalRef/js/alertify.min.js"); 
export interface IAddTaskWebPartProps {
  description: string;
}


declare var SP:any;
declare var SPClientPeoplePicker_InitStandaloneControlWrapper:any;
declare var SPClientPeoplePicker:any;
declare var datepicker:any
var NodeID="";
var siteURL="";
var publicHolidays=[];

export default class AddTaskWebPart extends BaseClientSideWebPart <IAddTaskWebPartProps> {

  public onInit(): Promise<void> {
    return super.onInit().then(() => {
      sp.setup({
         spfxContext: this.context
         });
    });
  }

  public render(): void {
    this.domElement.innerHTML = `
    <div class="section" >
    <h2>SCHEDULING DETAILS</h2>

    <div class="container">
    <div class="nid">
        <div class="labelDiv">            
        <label class="nodeid" for="">Node ID</label>
</div>
<div class="inputDiv">
<input type="text" id="txtNodeID" value="">
</div>
    </div>
    <div class="client">
        <div class="labelDiv">            
            <label for="">Client</label>
        </div>
        <div class="inputDiv"><select  name="Select" id="drpClient">
          
        </select></div>
       
    </div>
</div>

<div class="container">
<div class="sitename">
    <div class="labelDiv">
        <label for="">SiteName</label>
    </div>
    <div class="inputDiv test">
        <input class="sname" type="text" id="txtSiteName" value="">
    </div>
</div>
</div>

<div class="container">
<div class="version">
    <div class="labelDiv">            
    <label  for="">Version</label>
</div>
<div class="inputDiv">
<input type="text" id="txtVersion" value="">
</div>
</div>
<div class="sitetype">
    <div class="labelDiv">            
        <label for="">Site Type</label>
    </div>
    <div class="inputDiv"><select id="drpSiteType" name="Select">
        
    </select></div>
   
</div>
</div>

<div class="container">
<div class="priority">
    <div class="labelDiv">            
    <label  for="">Priority</label>
</div>
<div class="inputDiv">
<input type="text" id="txtPriority" value="">
</div>
</div>
<div class="architecture">
    <div class="labelDiv">            
        <label for="">Architecture</label>
    </div>
    <div class="inputDiv"><select  id="drpArchitecture" name="Select">
      
    </select></div>
   
</div>
</div>


<div class="container"> 
<div class="category">
    <div class="labelDiv">            
    <label  for="">Category</label>
</div>
<div class="inputDiv"><select id="drpCategory" name="Select">

</select></div>

</div>
<div class="sub-category" style="display:none">
    <div class="labelDiv">            
        <label for="">Sub-Category</label>
    </div>
    <div class="inputDiv"><select id="drpSubCategory" name="Select">

    </select></div>
   
</div>
</div>

    
<div class="task-section"></div>
    

        </div> 
        </div>

        <div class="btn-sec"><input type="button" value="submit" name="submit" id="btnSubmit" class="submit-btn"/></div></div> 
          `;
    getClients();
    getSiteType();
    getArchitecture();
    getCategory();
    // getSubCategory();
    getPublicHolidays();

    $("#txtNodeID").blur((e)=>{
      NodeID=e.target["value"];
      getMetaData(NodeID)
    });

    $("#btnSubmit").click((e)=>{
  //  if(this.MandatoryValidation())
  //  {
  //   let saveData = {
  //     Nodeid: $("#txtNodeID").val(),       
  //     SiteName: $("#txtSiteName").val(),
  //     ClientId:Number($('#drpClient option:selected').attr("id")),
  //     Version_x0023_:$("#txtVersion").val(),
  //     SiteType:Number($('#drpSiteType option:selected').attr("id")),
  //     Priority:$("#txtPriority").val(),
  //     Architecture:Number($('#drpArchitecture option:selected').attr("id")),
  //     Category:Number($('#drpCategory option:selected').attr("id"))
  //  }
  // }
  
  getUserID()


 
    });

    $("#drpCategory").change((e)=>{
     var SelectedItem=e.target["value"];
      if(SelectedItem!="Select")
      {
        getSubCategory(SelectedItem);
      }
    });


  }
 


    
 

  MandatoryValidation(): any {
    if (!$("#txtNodeID").val()) {
      alertify.set("notifier", "position", "top-right");
      alertify.error("Please Enter Node ID");
      return false;
    } else if (!$("#drpClient").val()||$("#drpClient").val()=="Select") {
      alertify.set("notifier", "position", "top-right");
      alertify.error("Please Select the Client");
      return false;
    } else if (!$("#txtSiteName").val()) {
      alertify.set("notifier", "position", "top-right");
      alertify.error("Please Enter Site Name");
      return false;
    }else if (!$("#txtVersion").val()) {
      alertify.set("notifier", "position", "top-right");
      alertify.error("Please Enter Version");
      return false;
    } else if (!$("#drpSiteType").val()||$("#drpSiteType").val()=="Select") {
      alertify.set("notifier", "position", "top-right");
      alertify.error("Please Select Site Type");
      return false;
    }else if (!$("#txtPriority").val()) {
      alertify.set("notifier", "position", "top-right");
      alertify.error("Please Enter Priority");
      return false;
    } else if (!$("#drpArchitecture").val()||$("#drpArchitecture").val()=="Select") {
      alertify.set("notifier", "position", "top-right");
      alertify.error("Please Select Architecture");
      return false;
    }
    else if (!$("#drpCategory").val()||$("#drpCategory").val()=="Select") {
      alertify.set("notifier", "position", "top-right");
      alertify.error("Please Select Category");
      return false;
    } else if (!$("#drpSubCategory").val()||$("#drpSubCategory").val()=="Select") {
      alertify.set("notifier", "position", "top-right");
      alertify.error("Please Select SubCategory");
      return false;
    }
    return true;

   
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

async function getUserID()
{
  var allTasks=$('.task-section .hop-review-section');
  allTasks.map((i,val)=>{
    var crtItem=val.childNodes[2].childNodes
    var data={};
    var objects = {};
    crtItem.forEach((innerItem,idx)=>{
      
      if(idx==1)
      {
      var getUser=innerItem["id"]+"_TopSpan_HiddenInput";
      var assigneID =  $('#'+getUser).val().toString();
      if(assigneID)
      {
        var finalJSON= JSON.parse(assigneID)
        var reviewerID=finalJSON[0]["Key"]
         data[innerItem["id"]+"Id"]=reviewerID
      }
       

      }
      else if(idx==2)
      {
       var splittedval= innerItem["id"].split("Assigneddatepicker");
        var duedateval=innerItem["id"];
        var duedate;
        if(duedateval)
        {
          data[splittedval[0]+"DueDate"]= $('#'+innerItem["id"]).val();
          duedate= $('#'+innerItem["id"]).val();
        }
      }
      if(idx==2 )
      {
        objects[innerItem["id"]]={"reviewer":assigneID,"duedate":duedate}
         
      }
      if(Object.keys(data).length>1)
      addItemsToList(data,Object.values(data)[0],Object.keys(data)[0]);

    })
  });


 
}

function addItemsToList(data,reviewer,ID) {
  sp.web.ensureUser(reviewer).then((result)=>{
    var itemID=result.data["Id"];
    if(itemID)
    data[ID]=itemID
    // console.log(itemID,dueDate);

addItems("iScheduleTasks",data)
      });
}

function PeoplePopulate(SelectedItem)
  {
    SPComponentLoader.loadCss('/_layouts/15/1033/styles/corev15.css');  
    
    SPComponentLoader.loadScript('/_layouts/15/init.js', {
      globalExportsName: '$_global_init'
    })
    .then((): Promise<{}> => {
      return SPComponentLoader.loadScript('/_layouts/15/MicrosoftAjax.js', {
        globalExportsName: 'Sys'
      });
    })
    .then((): Promise<{}> => {
      return SPComponentLoader.loadScript('/_layouts/15/ScriptResx.ashx?name=sp.res&culture=en-us', {
        globalExportsName: 'Sys'
      });
    })
    .then((): Promise<{}> => {
      return SPComponentLoader.loadScript('/_layouts/15/SP.Runtime.js', {
        globalExportsName: 'SP'
      });
    })
    .then((): Promise<{}> => {
      return SPComponentLoader.loadScript('/_layouts/15/SP.js', {
        globalExportsName: 'SP'
      });
    })            
    .then((): Promise<{}> => {
      return SPComponentLoader.loadScript('/_layouts/15/sp.init.js', {
        globalExportsName: 'SP'
      });
    })  
    .then((): Promise<{}> => {
      return SPComponentLoader.loadScript('/_layouts/15/1033/strings.js', {
        globalExportsName: 'Strings'
      });
    })      
    .then((): Promise<{}> => {
      return SPComponentLoader.loadScript('/_layouts/15/sp.ui.dialog.js', {
        globalExportsName: 'SP'
      });
    })
    .then((): Promise<{}> => {
      return SPComponentLoader.loadScript('/_layouts/15/clienttemplates.js', {
        globalExportsName: 'SP'
      });
    })
    .then((): Promise<{}> => {
      return SPComponentLoader.loadScript('/_layouts/15/clientforms.js', {
        globalExportsName: 'SP'
      });
    })
    .then((): Promise<{}> => {
      return SPComponentLoader.loadScript('/_layouts/15/clientpeoplepicker.js', {
        globalExportsName: 'SP'
      });
    })
    .then((): Promise<{}> => {
      return SPComponentLoader.loadScript('/_layouts/15/autofill.js', {
        globalExportsName: 'SP'
      });
    })
    .then((): Promise<{}> => {
      return SPComponentLoader.loadScript('/_layouts/15/sp.core.js', {
        globalExportsName: 'SP'
      });
    })
    .then(function(){      
      SP.SOD.executeOrDelayUntilScriptLoaded(function() {                
        var schema = {};  
        schema['PrincipalAccountType'] = 'User,DL';  
        schema['SearchPrincipalSource'] = 15;  
        schema['ResolvePrincipalSource'] = 15;  
        schema['AllowMultipleValues'] = false;  
        schema['MaximumEntitySuggestions'] = 50;  
        schema['Width'] = '280px';  
        SPClientPeoplePicker_InitStandaloneControlWrapper(SelectedItem, null, schema); 
     }, 'clientpeoplepicker.js');      
    }).then(function(){
      var picker = SPClientPeoplePicker.SPClientPeoplePickerDict[ SelectedItem+"_TopSpan"];
      GetPeoplePickerIds(SelectedItem);
      // picker.OnValueChangedClientScript = function (elementId, userInfo) {
      //   GetPeoplePickerIds(SelectedItem);
      // }; 
    });
  }

async function GetPeoplePickerIds(eleId) {
  var peoplePicker = GetPeoplePicker(eleId);
  if (peoplePicker != null) {
      // Get information about all users.
      var users = peoplePicker.GetAllUserInfo();
      var userInfo = [];
      for (var i = 0; i < users.length; i++) {
          var user = users[i];
          //userInfo += user['DisplayText'] + ";#";
          var userID = await sp.web.ensureUser(user.Key).then((result)=>{
            var itemID=result.data["Id"]
            userInfo.push(itemID);

          });

      console.log(userInfo)
      }
      return userInfo;
  } else
      return '';
}

function GetPeoplePicker(eleId) {
if(eleId != undefined){
    var toSpanKey = eleId + "_TopSpan";
    var peoplePicker = null;
    var ClientPickerDict = SPClientPeoplePicker.SPClientPeoplePickerDict;
    for (var propertyName in ClientPickerDict) {
        if (propertyName == toSpanKey) {
            peoplePicker = ClientPickerDict[propertyName];
            break;
        }
    }
    return peoplePicker;
}
}


async function getPublicHolidays()
{
  let objResults = readItems("PublicHolidays",["Date", "Title"],5000,"Modified").then((items: any) => {
    items.forEach(item => {
    publicHolidays.push(moment(item.Date).format('MM/DD/YYYY'))      
    });
  });
}


async function  getClients()
{
  let objResults = readItems("Client",["ID", "Title"],5000,"Modified").then((items: any) => {
    var drphtml = "<option value='Select'>Select</option>";
    items.forEach(element => {
      drphtml+=" <option value="+element.Title+" id="+element.ID+">"+element.Title+"</option>"
    });
   $("#drpClient").append(drphtml);
  });
}

async function  getSiteType()
{
  readItems("SiteType",["ID", "Title"],5000,"Modified").then((items: any) => {
    var drphtml = "<option value='Select'>Select</option>";
    items.forEach(element => {
      drphtml+=" <option value='"+element.Title+"' id="+element.ID+">"+element.Title+"</option>"
    });
   $("#drpSiteType").append(drphtml);
  });
}
async function  getArchitecture()
{
  readItems("ArchitectureType",["ID", "Title"],5000,"Modified").then((items: any) => {
    var drphtml = "<option value='Select'>Select</option>";
    items.forEach(element => {
      drphtml+=" <option value='"+element.Title+"' id="+element.ID+">"+element.Title+"</option>"
    });
   $("#drpArchitecture").append(drphtml);
  });
}

async function  getCategory()
{
  readItems("Category",["ID", "Title"],5000,"Modified").then((items: any) => {
    var drphtml = "<option value='Select'>Select</option>";
    items.forEach(element => {
      drphtml+=" <option value='"+element.Title+"' id="+element.ID+">"+element.Title+"</option>"
    });
   $("#drpCategory").append(drphtml);
  });
} 
async function  getSubCategory(SelectedItem) 
{
  var html="";
  $("#drpSubCategory").empty();
  readItem("SubCategory",["ID", "Title","Category/Title","Category/Title"],5000,"Created","Category/Title",SelectedItem,"Category").then((items: any) => {
    // var drphtml = "<option value='Select'>Select</option>";
    items.forEach(element => {
      // drphtml+=" <option value='"+element.Title+"'>"+element.Title+"</option>";

      var formattedSelectedItem = element.Title.replace(/\s|&/g, ""); 

      html+='<div class="hop-review-section"><h3>'+element.Title+'</h3><i class="item-show arrow right"></i><div class="parentDiv master-div"><div class="labelDiv"><label>Reviewer</label></div><div id="'+formattedSelectedItem+'" class="inputDiv"></div><div class="parentDiv"><div class="labelDiv"><label>DueDate:</label></div><input type="text" id="'+formattedSelectedItem+'datepicker" class="datepicker"></div></div></div>';
      $('#heading-2').text(element.Title);       
      
      

      PeoplePopulate(formattedSelectedItem)

    });
  //  $("#drpSubCategory").append(drphtml);
   $('.task-section').append(html);
   $('.master-div').hide();
   $('.hop-review-section').show();
   $('.item-show').click(function(){
   if($(this).hasClass( "right" ))
   {
    $(this).removeClass('right');
    $(this).addClass('down');
   }
   else{
    $(this).removeClass('down');
    $(this).addClass('right');
   }

     $(this).next().toggle();
   })

   let elem: any;
      
     elem = $(".datepicker");
     elem.datepicker({ dateFormat: 'dd/mm/yy', beforeShowDay: function(crtdate) {
     var currentDate=  moment(crtdate).format('MM/DD/YYYY')
       var show = true;
       if(crtdate.getDay()==6||crtdate.getDay()==0|| publicHolidays.indexOf(currentDate) != -1 ) show=false
       return [show];

   } });

  });
}

async function getMetaData(NodeID) {
  readItem("SiteList",["*,Client/Title,SiteType/Title,Architecture/Title"],5000,"Modified","NodeId",NodeID,"Client,SiteType,Architecture").then((items: any) => {

    if(items.length>0)
    {
      $('#txtSiteName').val(items[0].SiteName);
      $('#txtSiteName').val(items[0].SiteName);
      $('#drpClient').val(items[0].Client.Title);
      $('#txtVersion').val(items[0].Version_x0023_);
      $('#drpSiteType').val(items[0].SiteType.Title);
      $('#drpArchitecture').val(items[0].Architecture.Title);
    }
    else{
      alertify.set("notifier", "position", "top-right");
      alertify.error("No Sites in this Node ID");
      $('#txtSiteName').val("");
      $('#txtSiteName').val("");
      $('#drpClient').val("Select");
      $('#txtVersion').val("");
      $('#drpSiteType').val("Select");
      $('#drpArchitecture').val("Select");
    }

  });
}



