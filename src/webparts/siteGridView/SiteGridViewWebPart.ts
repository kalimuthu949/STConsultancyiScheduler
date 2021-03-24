import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './SiteGridViewWebPart.module.scss';
import * as strings from 'SiteGridViewWebPartStrings';
import "jquery";
import "datatables";
import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import "../../ExternalRef/css/StyleSchedule.css"; 
import "../../../node_modules/datatables/media/css/jquery.dataTables.min.css";
import { readItems,readItem, addItems,formatDate } from "../../commonJS";
export interface ISiteGridViewWebPartProps {  
  description: string;  s
}

export default class SiteGridViewWebPart extends BaseClientSideWebPart <ISiteGridViewWebPartProps> {


  public onInit(): Promise<void> {
    return super.onInit().then(() => {
      sp.setup({
         spfxContext: this.context
         });
    });
  }


  public render(): void {
    this.domElement.innerHTML = `
    <div id='SiteList'>
    <table id="SiteListTable" style="width:100%">
    <thead>    
    <tr>
    <th>Client</th>
    <th>Site Name</th>
    <th>Node ID</th>
    <th>Version #</th>
    <th>Site Type</th>
    <th>Category</th> 

    <th>Architecture</th>
    <th>WIP/Mirage #</th>
    <th>CANRAD ADDRESS ID</th>
    <th>RFNSA Site #</th>
    <th>DONOR Node Code</th>
    <th>SP/WO </th>
    <th>CP</th>
    <th>Existing Technology</th>
    <th>Added Technology</th>
    <th>Remote New Cells</th>
    <th>Donor New Cells</th>
    <th>SOW</th>
    <th>State</th>
    <th>Drawing Number(s)</th>
    <th>SAED/SMR Handler</th>
    <th>AMENDMENT Name</th>
    <th>Comments</th>
    </tr>
    </thead>
    <tbody id='tblbodySiteListTable'>
    </tbody>
    
    </table>
    </div>`;

    getSiteListDetails()
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


function getSiteListDetails() {
  var html="";
  readItem("SiteList",["*,Category/Title,Client/Title,SiteType/Title,Architecture/Title"],5000,"Modified","","","Client,SiteType,Architecture,Category").then((items: any) => {
    console.log(items);

    items.forEach(element => {
      html += "<tr>";
      (element.Client.Title)?html += "<td>" + element.Client.Title + "</td>":html += "<td>-</td>";
      (element.SiteName)? html += "<td>" + element.SiteName + "</td>":html += "<td>-</td>";
      (element.NodeId)?  html += "<td>" + element.NodeId + "</td>":html += "<td>-</td>";
      (element.Version_x0023_)?  html += "<td>" + element.Version_x0023_ + "</td>":html += "<td>-</td>";
      (element.Category.Title)?html += "<td>" + element.Category.Title + "</td>":html += "<td>-</td>";
      (element.Architecture.Title)? html += "<td>" + element.Architecture.Title + "</td>":html += "<td>-</td>";
      (element.SiteType.Title)?  html += "<td>" + element.SiteType.Title+ "</td>":html += "<td>-</td>";
      (element.WIP_x002f_Mirage_x0023_)? html += "<td>" + element.WIP_x002f_Mirage_x0023_ + "</td>":html += "<td>-</td>";
      (element.CanradAddressId)? html += "<td>" + element.CanradAddressId + "</td>":html += "<td>-</td>";
      (element.RfnsaSite_x0023_)? html += "<td>" + element.RfnsaSite_x0023_ + "</td>":html += "<td>-</td>";
      (element.DonorNodeCode)? html += "<td>" + element.DonorNodeCode + "</td>":html += "<td>-</td>";
      (element.SP_x002f_WO)?  html += "<td>" + element.SP_x002f_WO + "</td>":html += "<td>-</td>";
      (element.CP)?html += "<td>" + element.CP + "</td>":html += "<td>-</td>";
      (element.ExistingTechnology)?html += "<td>" + element.ExistingTechnology + "</td>":html += "<td>-</td>";
      (element.AddedTechnology)? html += "<td>" + element.AddedTechnology + "</td>":html += "<td>-</td>";
      (element.RemoteNewCells)? html += "<td>" + element.RemoteNewCells + "</td>":html += "<td>-</td>";
      (element.DonorNewCells)? html += "<td>" + element.DonorNewCells + "</td>":html += "<td>-</td>";
      (element.SOW)? html += "<td>" + element.SOW + "</td>":html += "<td>-</td>";
      (element.State)? html += "<td>" + element.State + "</td>":html += "<td>-</td>";
      (element.DrawingNumbers)? html += "<td>" + element.DrawingNumbers + "</td>":html += "<td>-</td>";
      (element.SAED_x002f_SMRHandler)? html += "<td>" + element.SAED_x002f_SMRHandler + "</td>":html += "<td>-</td>";
      (element.AmendmentName)? html += "<td>" + element.AmendmentName + "</td>":html += "<td>-</td>";
      (element.Comments)? html += "<td>" + element.Comments + "</td>":html += "<td>-</td>";
      html += "</tr>";
    });

    $("#tblbodySiteListTable").empty();
    $("#tblbodySiteListTable").append(html);

   (<any>$("#SiteListTable")).dataTable({
      scrollX: true,
      responsive: true
    });

  });

}
