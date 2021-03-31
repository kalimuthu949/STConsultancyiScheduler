import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { escape } from "@microsoft/sp-lodash-subset";

import styles from "./AddTaskListWebPart.module.scss";
import * as strings from "AddTaskListWebPartStrings";


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
export interface IAddTaskListWebPartProps {
  description: string;
}

  
export default class AddTaskListWebPart extends BaseClientSideWebPart<IAddTaskListWebPartProps> {
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
    <div class ="row m-2">
    <div class="col">
    <h3 class="text-center">Add / Edit Record</h3>
    </div>
    </div>
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
   <div class="col-9 m-2 text-end">
     <button class="btn  buttoncolor" id="btnSubmit" type="submit">Submit</button>
     <button class="btn  buttoncolor" id="btnClose" type="close">Close</button>
   </div>
 </div>
</div>`;

    var pagehtmldivision = ` <div class ="row m-2">
    <div class="col">
    <h3 class="text-center">Add / Edit Record</h3>
    </div>
    </div>
    <div class="row rowpadding justify-content-around">
<div class="col-md-4  m-2">
    <label for="Division" class="form-label">Division</label>
    <input type="text" class="form-control"  id ="txtDivision">
  </div>
  </div>
  <div class="row m-1 justify-content-around">
  <div class="col-4  text-end">
    <button class="btn  buttoncolor" id="btnSubmit" type="submit">Submit</button>
    <button class="btn  buttoncolor" id="btnClose" type="close">Close</button>
  </div>
</div>`;

    var pagehtmlbusinessdivision = `
    <div class ="row m-2">
    <div class="col">
    <h3 class="text-center">Add / Edit Record</h3>
    </div>
    </div>
    <div class="row rowpadding justify-content-around">
<div class="col-md-4  m-2">
    <label for="BusinessDivision" class="form-label">Business Division</label>
    <input type="text" class="form-control"  id ="txtBusinessDivision">
  </div>
  </div>
  <div class="row m-1 justify-content-around">
  <div class="col-4  text-end">
    <button class="btn  buttoncolor" id="btnSubmit" type="submit">Submit</button>
    <button class="btn  buttoncolor" id="btnClose" type="close">Close</button>
  </div>
</div>`;

    var pagehtmlprojects = ` <div class ="row m-2">
    <div class="col">
    <h3 class="text-center">Add / Edit Record</h3>
    </div>
    </div>

    <div class="row rowpadding justify-content-around">
 <div class="col-md-4  m-2">
     <label for="Projects" class="form-label">Projects</label>
     <input type="text" class="form-control"  id ="txtProjects">
   </div>
   </div>
   <div class="row m-1 justify-content-around">
  <div class="col-4 text-end">
    <button class="btn  buttoncolor" id="btnSubmit" type="submit">Submit</button>
    <button class="btn  buttoncolor" id="btnClose" type="close">Close</button>
  </div>
</div>`;

//     var pagehtmlpriority = `<div class="row rowpadding justify-content-around">
//  <div class="col-md-4  m-2">
//    <label for="Priority" class="form-label">Priority</label>
//    <input type="text" class="form-control"  id ="txtPriority">
//  </div>
//  </div>
//  <div class="row justify-content-around">
//   <div class="col-4 text-end">
//     <button class="btn  buttoncolor" id="btnSubmit" type="submit">Submit</button>
//   </div>
// </div>`;

var pagehtmlclient = ` 
<div class ="row m-2">
<div class="col">
<h3 class="text-center">Add / Edit Record</h3>
</div>
</div>
<div class="row rowpadding justify-content-around">
<div class="col-md-4  m-2">
  <label for="Client" class="form-label">Client</label>
  <input type="text" class="form-control"  id ="txtClient">
</div>
</div>
<div class="row justify-content-around">
 <div class="col-4 text-end">
   <button class="btn  buttoncolor" id="btnSubmit" type="submit">Submit</button>
   <button class="btn  buttoncolor" id="btnClose" type="close">Close</button>
 </div>
</div> `;

    this.domElement.innerHTML = `
    <div class="container " id="divcontainer"> 
</div>`;

    pagename = getUrlParameter("pagename");

    if (pagename == "Task") {
      $("#divcontainer").html(pagehtmlfortask);
    } else if (pagename == "Division") {
      $("#divcontainer").html(pagehtmldivision);
    } else if (pagename == "BusinessDivision") {
      $("#divcontainer").html(pagehtmlbusinessdivision);
    } else if (pagename == "Projects") {
      $("#divcontainer").html(pagehtmlprojects);
    }
    else if (pagename == "Client"){
      $("#divcontainer").html(pagehtmlclient);
    }

    getDivisions();
    getBusinessDivisions();
    getProjects();
    getClient();


    $("#btnClose").click(function() {
      location.href=siteURL+"/SitePages/AddConfigurationGrid.aspx"
    })


    $("#btnSubmit").click(async function () {
      if (pagename == "Task") {
        InsertTasks();
      } else if (pagename == "Division") {
        InsertDivision();
      } else if (pagename == "BusinessDivision") {
        InsertBusinessDivision();
      } else if (pagename == "Projects") {
        InstertProjects();
      }
      else if (pagename == "Client"){
        InstertClient();
      }
    });
  }

  protected getdataVersion(): Version {
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

/* This is place for onload functionalities start*/

function getDivisions() {
  var html = "";
  sp.web.lists
    .getByTitle("Division")
    .items.get()
    .then((items: any[]) => {
      // console.log(items);

      for (var i = 0; i < items.length; i++) {
        html +=
          '<option value="' + items[i].ID + '">' + items[i].Title + "</option>";
      }
      $("#Division").html("");
      $("#Division").html(html);
    })
    .catch(function (error) {
      ErrorCallBack(error, "getDivisions");
    });
}
function getBusinessDivisions() {
  var html = "";
  sp.web.lists
    .getByTitle("BusinessDivision")
    .items.get()
    .then((items: any[]) => {
      for (var i = 0; i < items.length; i++) {
        html +=
          '<option value="' + items[i].ID + '">' + items[i].Title + "</option>";
      }
      $("#BusinessDivision").html("");
      $("#BusinessDivision").html(html);
    })
    .catch(function (error) {
      ErrorCallBack(error, "getBusinessDivisions");
    });
}
function getProjects() {
  var html = "";
  sp.web.lists
    .getByTitle("Projects")
    .items.get()
    .then((items: any[]) => {
      for (var i = 0; i < items.length; i++) {
        html +=
          '<option value="' + items[i].ID + '">' + items[i].Title + "</option>";
      }
      $("#Projects").html("");
      $("#Projects").html(html);
    })
    .catch(function (error) {
      ErrorCallBack(error, "getProjects");
    });
}

function getClient(){
  var html ="";
  sp.web.lists.getByTitle("Client").items.get().then((items: any[]) => {
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

/* This is place for onload functionalities end*/

/* This is place for mandatory functionalities start*/

function mandatoryfieldsforTasks() {
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


async function InsertTasks() {
  if (mandatoryfieldsforTasks()) 
  
  {

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
      .items.add(requestdata)
      .then(function (data) {
        //$(".loading-modal").removeClass("active");
        //$("body").removeClass("body-hidden");
        AlertMessage("Record created successfully");
      })
      .catch(function (error) {
        ErrorCallBack(error, "InsertTasks");
      });
  }
}

function mandatoryfiledsforinsertdivision() {
  var isAllValueFilled = true;

  if (!$("#txtDivision").val()) {
    alertify.error("Please enter division");
    isAllValueFilled = false;
  }
  return isAllValueFilled;
}
/* This is place for mandatory functionalities End*/

/* This is place for submit functionalities start*/

async function InsertDivision() {
  if (mandatoryfiledsforinsertdivision()) {
    var requestdata = {
      Title: $("#txtDivision").val(),
    };
    await sp.web.lists
      .getByTitle("Division")
      .items.add(requestdata)
      .then(function (data) {
        //$(".loading-modal").removeClass("active");
        //$("body").removeClass("body-hidden");
        AlertMessage("Record created successfully");
      })
      .catch(function (error) {
        ErrorCallBack(error, "InsertDivision");
      });
  }
}



/* This is place for submit functionalities end*/

function mandatoryfiledsforinsertBusinessDivision() {
  var isAllValueFilled = true;

  if (!$("#txtBusinessDivision").val()) {
    alertify.error("Please enter Business Division");
    isAllValueFilled = false;
  }
  return isAllValueFilled;
}

async function InsertBusinessDivision() {
  if ( mandatoryfiledsforinsertBusinessDivision()) {
    var requestdata = {
      Title: $("#txtBusinessDivision").val(),
    };
    await sp.web.lists
      .getByTitle("BusinessDivision")
      .items.add(requestdata)
      .then(function (data) {
        //$(".loading-modal").removeClass("active");
        //$("body").removeClass("body-hidden");
        AlertMessage("Record created successfully");
      })
      .catch(function (error) {
        ErrorCallBack(error, "InsertBusinessDivision");
      });
  }
}

function mandatoryfiledsforProjects() {
  var isAllValueFilled = true;

  if (!$("#txtProjects").val()) {
    alertify.error("Please enter Project");
    isAllValueFilled = false;
  }
  return isAllValueFilled;
}


async function InstertProjects() {
  if ( mandatoryfiledsforProjects()) {
    var requestdata = {
      Title: $("#txtProjects").val(),
    };
    await sp.web.lists
      .getByTitle("Projects")
      .items.add(requestdata)
      .then(function (data) {
        //$(".loading-modal").removeClass("active");
        //$("body").removeClass("body-hidden");
        AlertMessage("Record created successfully");
      })
      .catch(function (error) {
        ErrorCallBack(error, "InstertProjects");
      });
  }
}

function mandatoryfiledsforClient() {
  var isAllValueFilled = true;

  if (!$("#txtClient").val()) {
    alertify.error("Please enter Client Details");
    isAllValueFilled = false;
  }
  return isAllValueFilled;
}


async function InstertClient() {
  if ( mandatoryfiledsforClient()) {
    var requestdata = {
      Title: $("#txtClient").val(),
    };
    await sp.web.lists
      .getByTitle("Client")
      .items.add(requestdata)
      .then(function (data) {
        //$(".loading-modal").removeClass("active");
        //$("body").removeClass("body-hidden");
        AlertMessage("Record created successfully");
      })
      .catch(function (error) {
        ErrorCallBack(error, "InstertClient");
      });
  }
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

/* This is place for common  functionalities end*/
