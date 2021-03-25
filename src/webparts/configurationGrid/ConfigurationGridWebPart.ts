import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPComponentLoader } from "@microsoft/sp-loader";
import styles from './ConfigurationGridWebPart.module.scss';
import * as strings from 'ConfigurationGridWebPartStrings';

import "jquery";
import "../../ExternalRef/css/ConfigurationStyle.css";

SPComponentLoader.loadCss(
  "https://cdn.jsdelivr.net/npm/bootstrap@5.0.0-beta2/dist/css/bootstrap.min.css"
);

export interface IConfigurationGridWebPartProps {
  description: string;
}


export default class ConfigurationGridWebPart extends BaseClientSideWebPart<IConfigurationGridWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
    <div class="container">
    <ul class="nav nav-tabs">
      <li class="active"><a data-toggle="tab" href="#div-division">Division</a></li>
      <li><a data-toggle="tab" href="#div-business">BusinessDivision</a></li>
    </ul>

    <div class="tab-content">
      <div id="div-division">
        <table id="table">
          <tr>
            <th>S.No</th>
            <th>Name</th>
            <th>Place</th>
          </tr>
          <tr>
            <td>1</td>
            <td>Ram</td>
            <td>KKDI</td>
          </tr>
          <tr>
            <td>2</td>
            <td>Arun</td>
            <td>MDU</td>
          </tr>
          <tr>
            <td>3</td>
            <td>Alagu</td>
            <td>TRY</td>
          </tr>
          <tr>
            <td>4</td>
            <td>Karthik</td>
            <td>SVG</td>
          </tr>
        </table>
      </div>
      <div id="div-business">
          <table id="table">
              <tr>
                <th>S.No</th>
                <th>Name</th>
                <th>Place</th>
              </tr>
              <tr>
                <td>1</td>
                <td>Ram</td>
                <td>KKDI</td>
              </tr>
              <tr>
                <td>2</td>
                <td>Arun</td>
                <td>MDU</td>
              </tr>
              <tr>
                <td>3</td>
                <td>Alagu</td>
                <td>TRY</td>
              </tr>
              <tr>
                <td>4</td>
                <td>Karthik</td>
                <td>SVG</td>
              </tr>
            </table>
      </div>
      </div>
    </div>`;
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
