import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField 
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface IListItem {
  Title: string;
  Id: number;
  SoftwareName: string;
  SoftwareVendor: string;
  SoftwareVersion: string;
  SoftwareDescription: string;
  SoftwareType: { Title: string; Id: number }; // Adjusted to match the lookup field structure
}

export interface IHelloWorldWebPartProps {
  description: string;
  listName: string;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {
  private softwareTypes: { Title: string, Id: number }[] = [];

  public async render(): Promise<void> {
    await this._fetchSoftwareTypes();
    this.domElement.innerHTML = `
      <div>
        <div>
          <table border='5' bgcolor='#FEBE10'>
            <tr>
              <td>Software ID</td>
              <td><input type="text" id="txtSoftwareId"></td>
            </tr>
            <tr>
              <td>Software Title</td>
              <td><input type="text" id="txtSoftwareTitle"></td>
            </tr>
            <tr>
              <td>Software Name</td>
              <td><input type="text" id="txtSoftwareName"></td>
            </tr>
            <tr>
              <td>Software Type</td>
              <td>
                <select id="ddlSoftwareType">
                  ${this.softwareTypes.map(type => `<option value="${type.Id}">${type.Title}</option>`).join('')}
                </select>
              </td>
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
                <input type="submit" value="Read" id="btnRead">
              </td>
            </tr>
          </table>
        </div>
        <div class="status" id="divStatus"></div>
        <div>
          <table id="itemsTable" border="1" cellspacing="0" cellpadding="5">
            <thead>
              <tr>
                <th>ID</th>
                <th>Title</th>
                <th>Software Name</th>
                <th>Software Vendor</th>
                <th>Software Version</th>
                <th>Software Description</th>
                <th>Software Type</th>
              </tr>
            </thead>
            <tbody class="items"></tbody>
          </table>
        </div>
      </div>
    `;

    this._bindEvents();
    this._fetchItems();
  }

  private _bindEvents(): void {
    this.domElement.querySelector("#btnSubmit")?.addEventListener("click", () => this.createItem());
    this.domElement.querySelector("#btnUpdate")?.addEventListener("click", () => this.updateItem());
    this.domElement.querySelector("#btnDelete")?.addEventListener("click", () => this.deleteItem());
    this.domElement.querySelector("#btnReadAll")?.addEventListener("click", () => this._fetchItems());
    this.domElement.querySelector("#btnRead")?.addEventListener("click", () => this.readItem());
  }

  private async _fetchSoftwareTypes(): Promise<void> {
    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('SoftwareLookUp')/items?$select=Id,Title`,
        SPHttpClient.configurations.v1
      );

      if (response.ok) {
        const items: { value: { Title: string, Id: number }[] } = await response.json();
        this.softwareTypes = items.value;
      } else {
        throw new Error(`Failed to fetch software types. Response: ${response.status} - ${response.statusText}.`);
      }
    } catch (error) {
      this.handleRequestError(error);
    }
  }

  private async createItem(): Promise<void> {
    const titleInput = this.domElement.querySelector("#txtSoftwareTitle") as HTMLInputElement;
    const nameInput = this.domElement.querySelector("#txtSoftwareName") as HTMLInputElement;
    const typeSelect = this.domElement.querySelector("#ddlSoftwareType") as HTMLSelectElement;
    const vendorSelect = this.domElement.querySelector("#ddlSoftwareVendor") as HTMLSelectElement;
    const versionInput = this.domElement.querySelector("#txtSoftwareVersion") as HTMLInputElement;
    const descriptionInput = this.domElement.querySelector("#txtSoftwareDescription") as HTMLTextAreaElement;

    if (!titleInput.value || !nameInput.value || !versionInput.value || !descriptionInput.value) {
      this.handleRequestError('Please fill in all the required fields.');
      return;
    }

    const body: string = JSON.stringify({
      'Title': titleInput.value,
      'SoftwareName': nameInput.value,
      'SoftwareTypeId': parseInt(typeSelect.value),
      'SoftwareVendor': vendorSelect.value,
      'SoftwareVersion': versionInput.value,
      'SoftwareDescription': descriptionInput.value
    });

    try {
      const listName = this.properties.listName;
      if (!listName) {
        throw new Error('List name is not defined. Please configure the web part properties.');
      }

      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata',
            'odata-version': ''
          },
          body: body
        }
      );

      if (response.ok) {
        const item: IListItem = await response.json();
        this.updateStatus(`Item '${item.Title}' (ID: ${item.Id}) successfully created`);
        this._fetchItems(); // Refresh the item list after adding a new item
      } else {
        const errorText = await response.text();
        throw new Error(`Failed to create item. Response: ${response.status} - ${response.statusText}. Details: ${errorText}`);
      }
    } catch (error) {
      this.handleRequestError(error);
    }
  }

  private async updateItem(): Promise<void> {
    const idInput = this.domElement.querySelector("#txtSoftwareId") as HTMLInputElement;
    const titleInput = this.domElement.querySelector("#txtSoftwareTitle") as HTMLInputElement;
    const nameInput = this.domElement.querySelector("#txtSoftwareName") as HTMLInputElement;
    const typeSelect = this.domElement.querySelector("#ddlSoftwareType") as HTMLSelectElement;
    const vendorSelect = this.domElement.querySelector("#ddlSoftwareVendor") as HTMLSelectElement;
    const versionInput = this.domElement.querySelector("#txtSoftwareVersion") as HTMLInputElement;
    const descriptionInput = this.domElement.querySelector("#txtSoftwareDescription") as HTMLTextAreaElement;

    if (!idInput.value || !titleInput.value || !nameInput.value || !versionInput.value || !descriptionInput.value) {
      this.handleRequestError('Please fill in all the required fields.');
      return;
    }

    const id = parseInt(idInput.value, 10);
    if (isNaN(id)) {
      this.handleRequestError('Please provide a valid numeric ID.');
      return;
    }

    const body: string = JSON.stringify({
      'Title': titleInput.value,
      'SoftwareName': nameInput.value,
      'SoftwareTypeId': parseInt(typeSelect.value),
      'SoftwareVendor': vendorSelect.value,
      'SoftwareVersion': versionInput.value,
      'SoftwareDescription': descriptionInput.value
    });

    try {
      const listName = this.properties.listName;
      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items(${id})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata',
            'odata-version': '',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'MERGE'
          },
          body: body
        }
      );

      if (response.ok) {
        this.updateStatus(`Item '${titleInput.value}' successfully updated`);
        this._fetchItems(); // Refresh the item list after updating an item
      } else {
        const errorText = await response.text();
        throw new Error(`Failed to update item. Response: ${response.status} - ${response.statusText}. Details: ${errorText}`);
      }
    } catch (error) {
      this.handleRequestError(error);
    }
  }

  private async deleteItem(): Promise<void> {
    const idInput = this.domElement.querySelector("#txtSoftwareId") as HTMLInputElement;

    if (!idInput.value) {
      this.handleRequestError('Please provide the ID of the item to delete.');
      return;
    }

    try {
      const listName = this.properties.listName;
      const id = idInput.value;
      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items(${id})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata',
            'odata-version': '',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'DELETE'
          }
        }
      );

      if (response.ok) {
        this.updateStatus(`Item with ID: ${id} successfully deleted`);
        this._fetchItems(); // Refresh the item list after deleting an item
      } else {
        const errorText = await response.text();
        throw new Error(`Failed to delete item. Response: ${response.status} - ${response.statusText}. Details: ${errorText}`);
      }
    } catch (error) {
      this.handleRequestError(error);
    }
  }

  private async _fetchItems(): Promise<void> {
    try {
      const listName = this.properties.listName;
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=Id,Title,SoftwareName,SoftwareVendor,SoftwareVersion,SoftwareDescription,SoftwareType/Title,SoftwareType/Id&$expand=SoftwareType`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        }
      );

      if (response.ok) {
        const items: { value: IListItem[] } = await response.json();
        this.renderItems(items.value);
      } else {
        const errorText = await response.text();
        throw new Error(`Failed to fetch items. Response: ${response.status} - ${response.statusText}. Details: ${errorText}`);
      }
    } catch (error) {
      this.handleRequestError(error);
    }
  }

  private renderItems(items: IListItem[]): void {
    let html: string = '';
    items.forEach((item: IListItem) => {
      html += `
        <tr>
          <td>${item.Id}</td>
          <td>${item.Title}</td>
          <td>${item.SoftwareName}</td>
          <td>${item.SoftwareVendor}</td>
          <td>${item.SoftwareVersion}</td>
          <td>${item.SoftwareDescription}</td>
          <td>${item.SoftwareType.Title}</td>
        </tr>
      `;
    });

    const itemsContainer: Element | null = this.domElement.querySelector('.items');
    if (itemsContainer) {
      itemsContainer.innerHTML = html;
    }
  }

  private async readItem(): Promise<void> {
    const idInput = this.domElement.querySelector("#txtSoftwareId") as HTMLInputElement;

    if (!idInput.value) {
      this.handleRequestError('Please provide the ID of the item to read.');
      return;
    }

    try {
      const listName = this.properties.listName;
      const id = idInput.value;
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items(${id})?$select=Id,Title,SoftwareName,SoftwareVendor,SoftwareVersion,SoftwareDescription,SoftwareType/Title&$expand=SoftwareType`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        }
      );

      if (response.ok) {
        const item: IListItem = await response.json();
        this.updateStatus(`Item '${item.Title}' (ID: ${item.Id}) successfully fetched: ${item.SoftwareName} - ${item.SoftwareVendor} - ${item.SoftwareVersion} - ${item.SoftwareDescription} - ${item.SoftwareType.Title}`);
      } else {
        const errorText = await response.text();
        throw new Error(`Failed to read item. Response: ${response.status} - ${response.statusText}. Details: ${errorText}`);
      }
    } catch (error) {
      this.handleRequestError(error);
    }
  }

  private handleRequestError(error: any): void {
    const statusDiv = this.domElement.querySelector("#divStatus") as HTMLElement;
    if (statusDiv) {
      statusDiv.innerHTML = `Error: ${error.message || error}`;
      statusDiv.style.color = "red";
    }
  }

  private updateStatus(status: string): void {
    const statusDiv = this.domElement.querySelector("#divStatus") as HTMLElement;
    if (statusDiv) {
      statusDiv.innerHTML = status;
      statusDiv.style.color = "green";
    }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: "Configure your web part"
          },
          groups: [
            {
              groupName: "Basic Settings",
              groupFields: [
                PropertyPaneTextField('description', {
                  label: "Description"
                }),
                PropertyPaneTextField('listName', {
                  label: "List Name"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
