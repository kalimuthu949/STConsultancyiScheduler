import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './EditTaskListWebPart.module.scss';
import * as strings from 'EditTaskListWebPartStrings';




import "jquery";
import { sp } from "@pnp/sp/presets/all";
import "../../ExternalRef/css/addTaskList.css";
import "../../ExternalRef/css/jquery.multiselect.css";
import "../../ExternalRef/js/jquery.multiselect.js";


import { SPComponentLoader } from "@microsoft/sp-loader";
SPComponentLoader.loadCss(
  "https://cdn.jsdelivr.net/npm/bootstrap@5.0.0-beta2/dist/css/bootstrap.min.css"
);
import "../../ExternalRef/css/alertify.min.css";



var alertify: any = require("../../ExternalRef/js/alertify.min.js");



declare var $;
var siteURL = "";
var pagename = "";
var itemid;

export interface IEditTaskListWebPartProps {
  description: string;
}

export default class EditTaskListWebPart extends BaseClientSideWebPart<IEditTaskListWebPartProps> {
  protected onInit(): Promise<void> {
    return super.onInit().then((_) => {
      siteURL = this.context.pageContext.web.absoluteUrl;
      sp.setup({
        spfxContext: this.context,
      });
    });
  }

  public render(): void {

    var pagehtmlfortask = `<div class="row rowpadding">
    <div class="row justify-content-around">
 <div class="col-md-4  m-2">
     <label for="Division" class="form-label">Division</label>
     <select class="form-select" id="Division">
         <option selected>Choose...</option>
         <option value="1">One</option>
         <option value="2">Two</option>
         <option value="3">Three</option>
       </select>
   </div>
   <div class="col-md-4 m-2">
     <label for="BusinessDivision" class="form-label">Business Division</label>
     <select class="form-select" id="BusinessDivision">
         <option selected>Choose...</option>
         <option value="1">One</option>
         <option value="2">Two</option>
         <option value="3">Three</option>
       </select>
   </div>
 </div>
   <div class="row justify-content-around">
   <div class="col-md-4  m-2">
     <label for="Projects" class="form-label">Projects</label>
     <select class="form-select" id="Projects">
         <option selected>Choose...</option>
         <option value="1">One</option>
         <option value="2">Two</option>
         <option value="3">Three</option>
       </select>
   </div>
   <div class="col-md-4  m-2">
     <label for="Tasks" class="form-label">Tasks</label>
     <input type="Tasks" class="form-control" id="Tasks">
   </div>
 </div>
 
 <div class="row justify-content-around">
   <div class="col-md-4  m-2">
     <label for="Priority" class="form-label">Priority</label>
     <input type="number" class="form-control"  id ="Priority">
   </div>
   <div class="col-md-4  m-2">
   <label for="Client" class="form-label">Client</label>
   <select name="Client[]" multiple class="form-select" id="Client">
         <option selected>Choose...</option>
         <option value="1">One</option>
         <option value="2">Two</option>
         <option value="3">Three</option>
       </select>
   
 </div>
 </div>
 <div class="row justify-content-around">
   <div class="col-9 text-end">
     <button class="btn  buttoncolor" id="btnUpdate" type="Update">Update</button>
   </div>
 </div>
</div>`;

    var pagehtmldivision = `<div class="row rowpadding justify-content-around">
<div class="col-md-4  m-2">
    <label for="Division" class="form-label">Division</label>
    <input type="text" class="form-control"  id ="txtDivision">
  </div>
  </div>
  <div class="row justify-content-around">
  <div class="col-4 text-end">
    <button class="btn  buttoncolor" id="btnUpdate" type="Update">Update</button>
  </div>
</div>`;

    var pagehtmlbusinessdivision = `<div class="row rowpadding justify-content-around">
<div class="col-md-4  m-2">
    <label for="BusinessDivision" class="form-label">Business Division</label>
    <input type="text" class="form-control"  id ="txtBusinessDivision">
  </div>
  </div>
  <div class="row justify-content-around">
  <div class="col-4 text-end">
    <button class="btn  buttoncolor" id="btnUpdate" type="Update">Update</button>
  </div>
</div>`;

    var pagehtmlprojects = `<div class="row rowpadding justify-content-around">
 <div class="col-md-4  m-2">
     <label for="Projects" class="form-label">Projects</label>
     <input type="text" class="form-control"  id ="txtProjects">
   </div>
   </div>
   <div class="row justify-content-around">
  <div class="col-4 text-end">
    <button class="btn  buttoncolor" id="btnUpdate" type="Update">Update</button>
  </div>
</div>`;

    var pagehtmlpriority = `<div class="row rowpadding justify-content-around">
 <div class="col-md-4  m-2">
   <label for="Priority" class="form-label">Priority</label>
   <input type="text" class="form-control"  id ="txtPriority">
 </div>
 </div>
 <div class="row justify-content-around">
  <div class="col-4 text-end">
    <button class="btn  buttoncolor" id="btnUpdate" type="Update">Update</button>
  </div>
</div>`;

var pagehtmlclient = ` <div class="row rowpadding justify-content-around">
<div class="col-md-4  m-2">
  <label for="Client" class="form-label">Client</label>
  <input type="text" class="form-control"  id ="txtClient">
</div>
</div>
<div class="row justify-content-around">
 <div class="col-4 text-end">
   <button class="btn  buttoncolor" id="btnUpdate" type="Update">Update</button>
 </div>
</div> `;

    this.domElement.innerHTML = ` <div class="container " id="divcontainer"> 
    </div>`;

    pagename = getUrlParameter("pagename");
    itemid=getUrlParameter("itemid");

    if (pagename == "Task") {
      $("#divcontainer").html(pagehtmlfortask);
    } else if (pagename == "Division") {
      $("#divcontainer").html(pagehtmldivision);
    } else if (pagename == "BusinessDivision") {
      $("#divcontainer").html(pagehtmlbusinessdivision);
    } else if (pagename == "Projects") {
      $("#divcontainer").html(pagehtmlprojects);
    } else if (pagename == "Priority") {
      $("#divcontainer").html(pagehtmlpriority);
    }
    else if (pagename == "Client"){
      $("#divcontainer").html(pagehtmlclient);
    }
    getDivisions();
    getBusinessDivisions();
    getProjects();
    getClient();



    // fetch
    if (pagename == "Task") {
    FetchTaskMasterList();
    }
    else if(pagename=="Division"){

      FetchDivision();
    }
    else if(pagename=="BusinessDivision"){
      FetchBusinessDivision();
    }
    else if(pagename=="Client"){
      FetchClient();
    }
    else if(pagename=="Projects"){
      FetchProjects();
    }
    
    // click

    $("#btnUpdate").click(async function (){
      if (pagename == "Task") 
      UpdateTaskMasterList();
      else if(pagename == "Projects")
      UpdateProjects();
      else if(pagename == "Client")
      UpdateClient();
      else if(pagename == "BusinessDivision")
      UpdateBusinessDivision();
      else if(pagename == "Division")
      UpdateDivision();
    });
  }
  



  protected getdataVersion(): Version {
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
/* This is place for onload functionalities start*/

async function getDivisions() {
  var html = "";
  await sp.web.lists
    .getByTitle("Division")
    .items.get()
    .then((items: any[]) => {
      // console.log(items);

      for (var i = 0; i < items.length; i++) {
        html +=
          '<option value="' + items[i].Title + '">' + items[i].Title + "</option>";
      }
      $("#Division").html("");
      $("#Division").html(html);
    })
    .catch(function (error) {
      ErrorCallBack(error, "getDivisions");
    });
}
 async function getBusinessDivisions() {
  var html = "";
  await sp.web.lists
    .getByTitle("BusinessDivision")
    .items.get()
    .then((items: any[]) => {
      for (var i = 0; i < items.length; i++) {
        html +=
          '<option value="' + items[i].Title + '">' + items[i].Title + "</option>";
      }
      $("#BusinessDivision").html("");
      $("#BusinessDivision").html(html);
    })
    .catch(function (error) {
      ErrorCallBack(error, "getBusinessDivisions");
    });
}
async function getProjects() {
  var html = "";
  await sp.web.lists
    .getByTitle("Projects")
    .items.get()
    .then((items: any[]) => {
      for (var i = 0; i < items.length; i++) {
        html +=
          '<option value="' + items[i].Title + '">' + items[i].Title + "</option>";
      }
      $("#Projects").html("");
      $("#Projects").html(html);
    })
    .catch(function (error) {
      ErrorCallBack(error, "getProjects");
    });
}

async function getClient(){
  var html ="";
  await sp.web.lists.getByTitle("Client").items.get().then((items: any[]) => {
    // console.log(items);
    for (var i=0; i<items.length; i++){
      html += '<option value="' + items[i].Title + '">' + items[i].Title + "</option>";
    }
    $("#Client").html("");
    $("#Client").html(html);

    $("#Client").multiselect({
      columns: 1,
      placeholder: "Select",
      search: true,
    });

})
.catch(function (error) {
  ErrorCallBack(error, "getClients");
});
}

 /*This place is for Fetch Details*/

 
 async function FetchDivision(){
  await sp.web.lists.getByTitle("Division").items.getById(itemid).get().then((items: any[]) => 
  {
  if(items)
  {
    $("#txtDivision").val(items['Title']);
  }
  else
  {
    alertify.error("error");
  }
 }).catch(function (error) {
   ErrorCallBack(error, "FetchDivision");
 });
 
}

async function FetchBusinessDivision(){
  await sp.web.lists.getByTitle("BusinessDivision").items.getById(itemid).get().then((items: any[]) => 
  {
  if(items)
  {
    $("#txtBusinessDivision").val(items['Title']);
  }
  else
  {
    alertify.error("error");
  }
 }).catch(function (error) {
   ErrorCallBack(error, "FetchBusinessDivision");
 });
 
}
async function FetchClient(){
  await sp.web.lists.getByTitle("Client").items.getById(itemid).get().then((items: any[]) => 
  {
  if(items)
  {
    $("#txtClient").val(items['Title']);
  }
  else
  {
    alertify.error("error");
  }
 }).catch(function (error) {
   ErrorCallBack(error, "FetchClient");
 });
 
}


async function FetchProjects(){
  await sp.web.lists.getByTitle("Projects").items.getById(itemid).get().then((items: any[]) => 
  {
  if(items)
  {
    $("#txtProjects").val(items['Title']);
  }
  else
  {
    alertify.error("error");
  }
 }).catch(function (error) {
   ErrorCallBack(error, "FetchProjects");
 });
 
}

async function FetchTaskMasterList(){
  await sp.web.lists.getByTitle("TaskMasterList").items.getById(itemid).get().then((items: any[]) => 
  {
  if(items)
  {
    $("#Tasks").val(items['Tasks']);
    $("#Priority").val(items['Priority']);
    $("#Projects").val(items['Projects']);
    $("#Division").val(items['Division']);
    $("#BusinessDivision").val(items['BusinessDivision']);


    var Clientvalue=items['Client'];
    if(Clientvalue)
    {
      var arrClientvalue=Clientvalue.split(";");
      arrClientvalue.pop(arrClientvalue.length-1);
      setdropdownvalues(arrClientvalue, "Client");
    }
  }
  else
  {
    alertify.error("error");
  }
 }).catch(function (error) {
   ErrorCallBack(error, "FetchTaskMasterList");
 });
 
}

/*This place for update*/
async function UpdateDivision(){
if( mandatoryfiledsforUpdateDivision()){

  await sp.web.lists
  .getByTitle("Division")
   .items.getById(itemid)
   .update({
 
     Title: $("#txtDivision").val()
   }).then(function (data) {
     AlertMessage("Record updated successfully");
   }).catch(function (error) {
       ErrorCallBack(error, "UpdateDivision");
     });
 }
}
async function UpdateBusinessDivision(){
  if( mandatoryfiledsforUpdateBusinessDivision() ){

    await sp.web.lists
    .getByTitle("BusinessDivision")
     .items.getById(itemid)
     .update({
   
       Title: $("#txtBusinessDivision").val()
     }).then(function (data) {
       AlertMessage("Record updated successfully");
     }).catch(function (error) {
         ErrorCallBack(error, "UpdateBusinessDivision");
       });
  }
 }

 async function UpdateClient(){
if(mandatoryfiledsforUpdateClient()){
   await sp.web.lists
   .getByTitle("Client")
    .items.getById(itemid)
    .update({
  
      Title: $("#txtClient").val()
    }).then(function (data) {
      AlertMessage("Record updated successfully");
    }).catch(function (error) {
        ErrorCallBack(error, "UpdateClient");
      });
  }
}

async function UpdateProjects(){
      if( mandatoryfiledsforUpdateProjects()) 
      {
     await sp.web.lists
     .getByTitle("Projects")
      .items.getById(itemid)
      .update({
    
        Title: $("#txtProjects").val()
      }).then(function (data) {
        AlertMessage("Record updated successfully");
      }).catch(function (error) {
          ErrorCallBack(error, "UpdateProjects");
        });
    }
}
  async function UpdateTaskMasterList(){
if( mandatoryfieldsforUpdateTaskMasterList() ){
 
   var client = "";
 
   $("#Client option:selected").each(function () {
     client += $(this).val() +";";
   });
 
     var requestdata = {
       Division: $("#Division option:selected").text(),
       BusinessDivision: $("#BusinessDivision option:selected").text(),
       Projects: $("#Projects option:selected").text(),
       Tasks: $("#Tasks").val(),
       Priority: $("#Priority").val(),
       Client:client
     };
 
   await sp.web.lists
   .getByTitle("TaskMasterList")
    .items.getById(itemid)
    .update(requestdata).then(function (data) {
      AlertMessage("Record updated successfully");
    }).catch(function (error) {
        ErrorCallBack(error, "UpdateTaskMasterList");
      });
  }
}



 /* This is place for mandatory functionalities start*/

 function mandatoryfiledsforUpdateDivision() {
  var isAllValueFilled = true;

  if (!$("#txtDivision").val()) {
    alertify.error("Please enter division");
    isAllValueFilled = false;
  }
  return isAllValueFilled;
}

function mandatoryfiledsforUpdateBusinessDivision() {
  var isAllValueFilled = true;

  if (!$("#txtBusinessDivision").val()) {
    alertify.error("Please enter BusinessDivision");
    isAllValueFilled = false;
  }
  return isAllValueFilled;
}


function mandatoryfiledsforUpdateClient() {
  var isAllValueFilled = true;

  if ($("#txtClient").val().length==0) {
    alertify.error("Please enter Client Details");
    isAllValueFilled = false;
  }
  return isAllValueFilled;
}

function mandatoryfiledsforUpdateProjects() {
  var isAllValueFilled = true;

  if (!$("#txtProjects").val()) {
    alertify.error("Please enter Project Details");
    isAllValueFilled = false;
  }
  return isAllValueFilled;
}

function mandatoryfieldsforUpdateTaskMasterList() {
  var isAllValueFilled = true;

  if (!$("#Tasks").val()) {
    alertify.error("Please Enter Task");
    isAllValueFilled = false;
  } else if (!$("#Priority").val()) {
    alertify.error("Please Enter Priority");
    isAllValueFilled = false;
  }
  else if($("#Client").val().length == 0) {
    alertify.error("Please Select Client Details");
    isAllValueFilled = false;
  }
 
  return isAllValueFilled;
}
  /* This is place for common  functionalities start*/

  function AlertMessage(Message) {
    alertify
      .alert()
      .setting({
        label: "OK",
  
        message: Message,
  
        onok: function () {
          window.location.href = siteURL + "/SitePages/ConfigurationGrid.aspx";
          //window.location.href = "#";
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


function setdropdownvalues(selectedOptions, id) {
  for (var i in selectedOptions) {
    var optionVal = selectedOptions[i];
    $("#" + id + "")
      .find("option[value='" + optionVal + "']")
      .prop("selected", "selected");
  }
  $("#" + id + "").multiselect("reload");
}

/* This is place for common  functionalities end*/