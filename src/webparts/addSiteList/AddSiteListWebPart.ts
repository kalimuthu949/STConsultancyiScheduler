import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPComponentLoader } from '@microsoft/sp-loader';
import styles from './AddSiteListWebPart.module.scss';
import * as strings from 'AddSiteListWebPartStrings';
import * as $ from "jquery"; 
import { sp } from "@pnp/sp/presets/all";
import * as JSZip from 'jszip';
import * as XLSX from 'xlsx';
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import * as moment from 'moment';  
import { readItems,readItem, addItems,formatDate } from "../../commonJS";
SPComponentLoader.loadCss('//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css')
SPComponentLoader.loadScript('//cdn.jsdelivr.net/npm/alertifyjs@1.13.1/build/alertify.min.js');
SPComponentLoader.loadScript('https://code.jquery.com/ui/1.12.1/jquery-ui.js');
// SPComponentLoader.loadScript('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.13.5/xlsx.full.min.js');
// SPComponentLoader.loadScript('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.13.5/jszip.js');
import "../../ExternalRef/css/styleSite.css";
import 'alertifyjs';
import "../../ExternalRef/css/alertify.min.css";
var alertify: any = require("../../ExternalRef/js/alertify.min.js"); 

export interface IAddSiteListWebPartProps {
  description: string;
}
declare var datepicker:any
// declare var XLSX:any;
export default class AddSiteListWebPart extends BaseClientSideWebPart <IAddSiteListWebPartProps> {

  public onInit(): Promise<void> {
    return super.onInit().then(() => {
      sp.setup({
         spfxContext: this.context
         });
    });
  }

  public render(): void {
    this.domElement.innerHTML = `

    <div class="Section">
    <fieldset>
    <legend>Site Details</legend>

    <div class="Container">
    <div class="Client">
      <div class="LabelDiv">
        <label for="">CLIENT</label>
      </div>
      <div class="inputDiv">
        <select name="Select" id="drpClient">

        </select>
      </div>
    </div>
    <div class="Category">
      <div class="LabelDiv">
        <label for="">CATEGORY</label>
      </div>
      <div class="inputDiv">
        <select name="Select" id="drpCategory">

        </select>
      </div>
    </div>
    <div class="Node">
      <div class="LabelDiv">
        <label for="">NODE ID</label>
      </div>
      <div class="inputDiv">
        <input type="text"  id="txtNodeID"/>
      </div>
    </div>
    <div class="Sitename">
      <div class="LabelDiv">
        <label for="">SITE NAME</label>
      </div>
      <div class="inputDiv">
        <input type="text" class="sname"  id="txtSiteName"/>
      </div>
    </div>
    <div class="Sitetype">
      <div class="LabelDiv">
        <label for="">SITE TYPE</label>
      </div>
      <div class="inputDiv">
        <select name="Select" id="drpSiteType">

        </select>
      </div>
    </div>
    <div class="Version">
      <div class="LabelDiv">
        <label for="">Version</label>
      </div>
      <div class="inputDiv">
        <input type="text"  id="txtVersion"/>
      </div>
    </div>
  </div>

  <div class="Container">
  <div class="Mirage">
    <div class="LabelDiv">
      <label for="">MIRAGE#</label>
    </div>
    <div class="inputDiv">
      <input type="text" id="txtMirage"/>
    </div>
  </div>
  <div class="Sp/fo">
    <div class="LabelDiv">
      <label for="">SP/FO</label>
    </div>
    <div class="inputDiv">
      <input type="text"  id="txtSPWO"/>
    </div>
  </div>
  <div class="sp">
    <div class="LabelDiv">
      <label for="">CP</label>
    </div>
    <div class="inputDiv">
      <input type="text" id="txtCP"/>
    </div>
  </div>
  <div class="Architecture">
    <div class="LabelDiv">
      <label for="">ARCHITECTURE</label>
    </div>
    <div class="inputDiv">
      <select name="Select" id="drpArchitecture" class="architecture">

      </select>
    </div>
  </div>
  <div class="canrad">
    <div class="LabelDiv">
      <label for="">CANRAD ID</label>
    </div>
    <div class="inputDiv">
      <input type="text"  id="txtCanradID"/>
    </div>
  </div>
  <div class="rfnsa">
    <div class="LabelDiv">
      <label for="">RFNSA#</label>
    </div>
    <div class="inputDiv">
      <input type="text" id="txtRFNSA"/>
    </div>
  </div>
</div>

<div class="Container1">
<div class="Exsitingtechnology">
  <div class="LabelDiv">
    <label for="">EXSITING TECHNOLOGY</label>
  </div>
  <div class="inputDiv">
    <input class="exsitingtechnology" type="text" id="txtExisting"/>
  </div>
</div>
<div class="Addedtechnology">
  <div class="LabelDiv">
    <label for="">ADDED TECHNOLOGY</label>
  </div>
  <div class="inputDiv">
    <input class="addedtechnology" type="text" id="txtAdded"/>
  </div>
</div>
<div class="Remote">
  <div class="LabelDiv">
    <label for="">REMOTE NEW CELLS</label>
  </div>
  <div class="inputDiv">
    <input class="remote" type="text" id="txtRemoteNewCells"/>
  </div>
</div>
</div>

<div class="Container2">
<div class="Drawing">
  <div class="LabelDiv">
    <label for="">DRAWING #</label>
  </div>
  <div class="inputDiv">
    <input class="drawing" type="text" id="txtDrawing"/>
  </div>
</div>
<div class="State">
  <div class="LabelDiv">
    <label for="">STATE</label>
  </div>
  <div class="inputDiv">
    <input class="state" type="text"  id="txtState"/>
  </div>
</div>
<div class="SoW">
  <div class="Labeldiv">
    <label for="">SOW </label>
  </div>
  <div class="inputdiv">
    <input class="sow" type="text"  id="txtSoW"/>
  </div>
</div>
<div class="SAED">
  <div class="LabelDiv">
    <label for="">SAED/SMR Handler </label>
  </div>
  <div class="inputDiv">
    <input class="saed" type="text" id="txtSAED"/>
  </div>
</div>
</div>

<div class="Container">
<div class="Amendment Name">
  <div class="LabelDiv">
    <label for="">AMENDMENT NAME</label>
  </div>
  <div class="inputDiv">
    <input class="amendment" type="text" id="txtAmendment"/>
  </div>
</div>
<div class="Priority">
  <div class="LabelDiv">
    <label for="">PRIORITY </label>
  </div>
  <div class="inputDiv">
    <input type="text" id="txtpriority" />
  </div>
</div>
<div class="Due Date">
  <div class="LabelDiv">
    <label for="">DUE DATE</label>
  </div>
  <div class="inputDiv">
    <input type="text" id="txtDueDate" class="datepicker"/>
  </div>
</div>
</div>

<div class="Container">
<div class="Comments">
  <div class="LabelDiv">
    <label for="">Comments</label>
  </div>
  <div class="inputDiv">
    <textarea name="nishanth"  id="txtComments" cols="4" rows="4"></textarea>
  </div>
</div>
</div>

    </fieldset>
    </div>
<div class="btn-sec"><input type="button" value="submit" name="submit" id="btnSubmit" class="submit-btn"/></div> 
    `;
var a=new JSZip();
    getClients();
    getSiteType();
    getArchitecture(); 
    getCategory();

    // document.getElementById('upload').addEventListener('change', handleFileSelect, false);

    $("#btnSubmit").click((e)=>{

     Upload();
      
      //   let saveData = {
      //     ClientId:Number($('#drpClient option:selected').attr("id")),
      //     CategoryId:Number($('#drpCategory option:selected').attr("id")),
      //     NodeId: $("#txtNodeID").val(),       
      //     SiteName: $("#txtSiteName").val(),
      //     SiteTypeId:Number($('#drpSiteType option:selected').attr("id")),
      //     Version_x0023_:$("#txtVersion").val(),
      //     WIP_x002f_Mirage_x0023_:$("#txtMirage").val(),
      //     SP_x002f_WO:$("#txtSPWO").val(),
      //     CP:$("#txtCP").val(),
      //     ArchitectureId:Number($('#drpArchitecture option:selected').attr("id")),
      //     CanradAddressId:$("#txtCanradID").val(),
      //     RfnsaSite_x0023_:$("#txtRFNSA").val(),
      //     ExistingTechnology:$("#txtExisting").val(),
      //     AddedTechnology:$("#txtAdded").val(),
      //     RemoteNewCells:$("#txtRemoteNewCells").val(),
      //     DrawingNumbers:$("#txtDrawing").val(),
      //     State:$("#txtState").val(),
      //     SOW:$("#txtSoW").val(),
      //     SAED_x002f_SMRHandler:$("#txtSAED").val(),
      //     AmendmentName:$("#txtAmendment").val(),
      // //    Priority:$("#txtpriority").val(),
      //   //  DueDate:$("#txtDueDate").val(),
      //     Comments:$("#txtComments").val()        
      //  }
      
      

      //  addItems("SiteList",saveData).then(()=>{
      //    alert('Success')
      //  })
    
    
     
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


function Upload(){
  //Reference the FileUpload element.
  var fileUpload = document.getElementById("fileUpload");

  //Validate whether File is valid Excel file.
  var regex = /^([a-zA-Z0-9\s_\\.\-:])+(.xls|.xlsx)$/;
  if (regex.test(fileUpload["value"].toLowerCase())) {
      if (typeof (FileReader) != "undefined") {
          var reader = new FileReader();

          //For Browsers other than IE.
          if (reader.readAsBinaryString) {
              reader.onload = function (e) {
                ProcessExcel(e.target["result"]);
              };
              reader.readAsBinaryString(fileUpload["files"][0]);
          } else {
              //For IE Browser.
              reader.onload = function (e) {
                  var data = "";
                  var bytes = new Uint8Array(e.target["result"]);
                  for (var i = 0; i < bytes.byteLength; i++) {
                      data += String.fromCharCode(bytes[i]);
                  }
                  ProcessExcel(data);
              };
              reader.readAsArrayBuffer(fileUpload["files"][0]);
          }
      } else {
          alert("This browser does not support HTML5.");
      }
  } else {
      alert("Please upload a valid Excel file.");
  }
};

 function ProcessExcel(data) {
  //Read the Excel File data.
  var workbook = XLSX.read(data, {
      type: 'binary'
  });

  //Fetch the name of First Sheet.
  var firstSheet = workbook.SheetNames[0];

  //Read all rows from First Sheet into an JSON array.
  var excelRows = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet]);


};
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
   (<any> $(".datepicker")).datepicker();
  });
} 

