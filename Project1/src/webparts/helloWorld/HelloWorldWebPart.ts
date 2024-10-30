import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'HelloWorldWebPartStrings';

import { ISPHttpClientOptions, SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

export interface IHelloWorldWebPartProps {
  description: string;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div>
        <div>
          <table border='5' bgcolor='aqua'>
           
            <tr>
              <td>Software Title</td>
              <td><input type="text" id="txtSoftwareTitle"></td>
            </tr>
            <tr>
            <td>Software Name</td>
            <td><input type="text" id="txtSoftwareName"></td>
          </tr>

            <tr>
              <td>Software Vendor</td>
              <td>
                <select id="ddlSoftwareVendor">
                  <option value="Microsoft">Microsoft</option>
                  <option value="Sun">Sun</option>
                  <option value="Oracle">Oracle</option>
                  <option value="Google">Google</option>
                </select>
              </td>
            </tr>
            <tr>
              <td>Software Version</td>
              <td><input type="text" id="txtSoftwareVersion"></td>
            </tr>
            <tr>
              <td>Software Description</td>
              <td><textarea rows="5" cols="40" id="txtSoftwareDescription"></textarea></td>
            </tr>
            <tr>
              <td colspan="2" align="center">
                <input type="submit" value="Insert Item" id="btnSubmit">
                <input type="submit" value="Update" id="btnUpdate">
                <input type="submit" value="Delete" id="btnDelete">
                <input type="submit" value="Show All Records" id="btnReadAll">
              </td>
            </tr>
          </table>
        </div>
        <div id="divStatus"></div>
      </div>
    `;

    this._bindEvents();
  }

  private _bindEvents(): void {
    this.domElement.querySelector("#btnSubmit")?.addEventListener("click", () => { this.addListItem(); });
  }

  private addListItem(): void {
    const softwareTitle = (document.getElementById("txtSoftwareTitle") as HTMLInputElement)?.value;
    const softwareName = (document.getElementById("txtSoftwareName") as HTMLInputElement)?.value;
    const softwareVersion = (document.getElementById("txtSoftwareVersion") as HTMLInputElement)?.value;
    const softwareVendor = (document.getElementById("ddlSoftwareVendor") as HTMLSelectElement)?.value;
    const softwareDescription = (document.getElementById("txtSoftwareDescription") as HTMLTextAreaElement)?.value;

    if (!softwareTitle || !softwareName || !softwareVersion || !softwareVendor || !softwareDescription) {
      let statusMessage = this.domElement.querySelector("#divStatus");
      if (statusMessage) {
        statusMessage.innerHTML = "Please fill out all fields.";
      }
      return;
    }

    const siteUrl: string = `${this.context.pageContext.site.absoluteUrl}/_api/web/lists/getbytitle('SoftwareCatalog')/items`;

    const itemBody: any = {
      "Title": softwareTitle,
      "SoftwareName": softwareName,
      "SoftwareVendor": softwareVendor,
      "SoftwareDescription": softwareDescription,
      "SoftwareVersion": softwareVersion
    };

    const spHttpClientOptions: ISPHttpClientOptions = {
      "body": JSON.stringify(itemBody)
    };

    this.context.spHttpClient.post(siteUrl, SPHttpClient.configurations.v1, spHttpClientOptions)
      .then((response: SPHttpClientResponse) => {
        let statusMessage = this.domElement.querySelector("#divStatus");
        if (statusMessage) {
          if (response.status === 201) {
            statusMessage.innerHTML = "List item has been created successfully.";
            this.clear();
          } else {
            statusMessage.innerHTML = `An error has occurred: ${response.status} - ${response.statusText}`;
          }
        }
      });
  }

  private clear(): void {
    (document.getElementById("txtSoftwareTitle") as HTMLInputElement).value = "";
    (document.getElementById("txtSoftwareName") as HTMLInputElement).value = "";
    (document.getElementById("txtSoftwareVersion") as HTMLInputElement).value = "";
    (document.getElementById("ddlSoftwareVendor") as HTMLSelectElement).value = "Microsoft";
    (document.getElementById("txtSoftwareDescription") as HTMLTextAreaElement).value = "";
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
                  label: strings.PropertyPaneDescription // Corrected reference
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
