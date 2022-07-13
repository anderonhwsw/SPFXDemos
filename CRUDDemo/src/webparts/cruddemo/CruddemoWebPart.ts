import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

//import styles from './CruddemoWebPart.module.scss';
import * as strings from 'CruddemoWebPartStrings';

import { ISPHttpClientOptions, SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { IEmpleadosListItem } from "./IEmpleadosListItem";

export interface ICruddemoWebPartProps {
  description: string;
}

export default class CruddemoWebPart extends BaseClientSideWebPart<ICruddemoWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

  public render(): void {
    this.domElement.innerHTML = `<div>
    <div>
    <table border='5' bgcolor='aqua'>

      <tr>
        <td>ID Item</td>
        <td><input type='text' id='txtID' />
        <td><input type='submit' id='btnRead' value='Read Details' />
        </td>
      </tr>
      
      <tr>
        <td>Title</td>
        <td><input type='text' id='txtSoftwareTitle'/></td>
      </tr>

      <tr>
        <td>Cedula</td>
        <td><input type='text' id='txtcedula' />
      </tr>

      <tr>
        <td>Cargo</td>
        <td>
          <select id="ddlCargo">
            <option value="Microsoft">Microsoft</option>
            <option value="Sun">Sun</option>
            <option value="Oracle">Oracle</option>
            <option value="Google">Google</option>
          </select>
        </td>
      </tr>

      <tr>
        <td>Nombre</td>
        <td><input type='text' id='txtNombre'/>
      </tr>

      <tr>
        <td>Area Desctiption</td>
        <td><textarea rows='5' cols='40' id='txtAreaDescription'></textarea></td>
      </tr>

      <tr>
        <td colspan='2' align='center'>
          <input type='submit' value='Insert Item' id='btnSubmit' />
          <input type='submit' value='Update' id='btnUpdate' />
          <input type='submit' value='Delete' id='btnDelete' />
          <input type='Submit' value='Show All Records' id='btnReadAll' />
        </td>
      </tr>
    </table>
    </div>
    <div id="divStatus"/>
    </div>`;

    this._binEvents();
    this.readAllItems();
  }

  private readAllItems(): void {

    this._getListItems()
      .then(listItems => {
        let html: string = '<table border=1 width=100% style="border-collapse: collapse;">';
        html += '<th>ID</th> <th>Title</th> <th>Nombre</th> <th>Cedula</th> <th>Cargo</th> <th>Descripción área</th>';
        
        listItems.forEach(listItem => {
          html += `<tr>
          <td>${listItem.Id}</td>
          <td>${listItem.Title}</td>
          <td>${listItem.Nombre}</td>
          <td>${listItem.Cedula}</td>
          <td>${listItem.Cargo}</td>
          <td>${listItem.Area}</td>
          </tr>`;
        });
        html += '</table>';
        const listContainer: Element = this.domElement.querySelector('#divStatus');
        listContainer.innerHTML = html;

      });
  }

  private _getListItems(): Promise<IEmpleadosListItem[]> {
    const url: string = this.context.pageContext.site.absoluteUrl + "/_api/web/lists/getbytitle('Empleados')/items";
    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then(response => {
        return response.json();
      })
      .then(json => {
        return json.value;
      }) as Promise<IEmpleadosListItem[]>;
  }

  private _binEvents(): void {
    this.domElement.querySelector('#btnSubmit').addEventListener('click', () => {
      this.addListItem();
    });
    this.domElement.querySelector('#btnRead').addEventListener('click', () => {
      this.readListItem();
    });
    this.domElement.querySelector('#btnUpdate').addEventListener('click', () => {
      this.updateListItem();
    })
    this.domElement.querySelector('#btnDelete').addEventListener('click', () => {
      this.deleteListItem();
    })
  }

  private deleteListItem(): void {
    
    let id: string = document.getElementById('txtID')["value"];
    
    const url: string = this.context.pageContext.site.absoluteUrl + "/_api/web/lists/getbytitle('Empleados')/items(" + id + ")";
    const headers: any = {
      "X-HTTP-Method": "DELETE",
      "IF-MATCH": "*"
    };
    const spHttpClientOptions: ISPHttpClientOptions = {
      "headers": headers
    };

    this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
      .then((response: SPHttpClientResponse) => {
        if (response.status === 204) {
          let message: Element = this.domElement.querySelector("#divStatus");
          message.innerHTML = "Delete: List item has been deleted successfully.";
        } else {
          let message: Element = this.domElement.querySelector("#divStatus");
          message.innerHTML = "Failed to Delete..." + response.status + " - " + response.statusText;
        }
      })

  }

  private updateListItem(): void {
    
    let title = document.getElementById("txtSoftwareTitle")["value"];
    let cedula = document.getElementById("txtcedula")["value"];
    let Cargo = document.getElementById("ddlCargo")["value"];
    let Nombre = document.getElementById("txtNombre")["value"];
    let Descripcionarea = document.getElementById("txtAreaDescription")["value"];
    
    let id: string = document.getElementById('txtID')["value"];

    const url: string = this.context.pageContext.site.absoluteUrl + "/_api/web/lists/getbytitle('Empleados')/items(" + id + ")";
    console.log(url);
    const itemBody: any = {
      "Title": title,
      "Cedula": cedula,
      "Nombre": Nombre,
      "Cargo": Cargo,
      "Area": Descripcionarea
    };

    const headers: any = {
      "X-HTTP-Method": "MERGE",
      "IF-MATCH": "*"
    };

    const spHttpClientOptions: ISPHttpClientOptions = {
      "headers": headers,
      "body": JSON.stringify(itemBody)
    };

    this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
      .then((response: SPHttpClientResponse) => {
        if (response.status === 204) {
          let message: Element = this.domElement.querySelector('#divStatus');
          message.innerHTML = "List Item has been update successfuly.";
        } else {
          let message: Element = this.domElement.querySelector('#divStatus');
          message.innerHTML = "List Item updation failed. " + response.status + " - " + response.statusText;
        }
      })

  }

  private readListItem(): void {
    let id: string = document.getElementById('txtID')["value"];
    this._getListItemByID(id)
      .then(listItem => {
        document.getElementById("txtSoftwareTitle")["value"] = listItem.Title;
        document.getElementById("txtcedula")["value"] = listItem.Cedula;
        document.getElementById("ddlCargo")["value"] = listItem.Cargo;
        document.getElementById("txtNombre")["value"] = listItem.Nombre;
        document.getElementById("txtAreaDescription")["value"] = listItem.Area;
      })

      .catch(error => {
        let message: Element = this.domElement.querySelector('#divStatus');
        message.innerHTML = "Read: Could not fetch details... " + error.message;
      });
  }

  //Al colocar la Promesa, dea error, se debe imoprtar el modulo
  private _getListItemByID(id: string): Promise<IEmpleadosListItem> {
    const url: string = this.context.pageContext.site.absoluteUrl + "/_api/web/lists/getbytitle('Empleados')/items?$filter=Id eq " + id;
    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      })
      .then((listItems: any) => {
        const untypedItem: any = listItems.value[0];
        const ListItem: IEmpleadosListItem = untypedItem as IEmpleadosListItem;
        return ListItem;
      }) as Promise <IEmpleadosListItem>
  }

  private addListItem(): void {

    let softwaretitle = document.getElementById("txtSoftwareTitle")["value"];
    let cedula = document.getElementById("txtcedula")["value"];
    let Nombre = document.getElementById("txtNombre")["value"];
    let Cargo = document.getElementById("ddlCargo")["value"];
    let Area = document.getElementById("txtAreaDescription")["value"];

    const siteurl: string = this.context.pageContext.site.absoluteUrl + "/_api/web/lists/getbytitle('Empleados')/items";

    const itembody: any = {
      "Title": softwaretitle,
      "Cedula": cedula,
      "Nombre": Nombre,
      "Cargo": Cargo,
      "Area": Area
    };

    console.log('url es: ');
    console.log(siteurl);

    const spHttpClientOptions: ISPHttpClientOptions = {
      "body": JSON.stringify(itembody)
    };

    

    this.context.spHttpClient.post(siteurl, SPHttpClient.configurations.v1, spHttpClientOptions)
      .then((response: SPHttpClientResponse) => {

        if (response.status === 201) {
          let statusmessage: Element = this.domElement.querySelector('#divStatus');
          statusmessage.innerHTML = "List item has been created successfuly.";
          this.clear();
        } else {
          let statusmessage: Element = this.domElement.querySelector('#divStatus');
          statusmessage.innerHTML = "An error has ocurred i.e. " + response.status + " - " + response.statusText;
        }
      });
  }

  private clear = () => {
    document.getElementById("txtSoftwareTitle")["value"] = '';
    document.getElementById("ddlCargo")["value"] = 'Microsoft';
    document.getElementById("txtNombre")["value"] = '';
    document.getElementById("txtcedula")["value"] = '';
    document.getElementById("txtAreaDescription")[""] = '';
  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;
    this.domElement.style.setProperty('--bodyText', semanticColors.bodyText);
    this.domElement.style.setProperty('--link', semanticColors.link);
    this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered);

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
