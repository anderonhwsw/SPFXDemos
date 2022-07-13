import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './SppnpjscruddemoWebPart.module.scss';
import * as strings from 'SppnpjscruddemoWebPartStrings';

import * as pnp from 'sp-pnp-js'; //Se importa para utilizar los metodos pnp
import { functionsIn } from 'lodash';

export interface ISppnpjscruddemoWebPartProps {
  description: string;
}

export default class SppnpjscruddemoWebPart extends BaseClientSideWebPart<ISppnpjscruddemoWebPartProps> {

  //Se agrega el siguiente fragmento de codigo al trabajar con PNP para que el ambiente reconozca las librerias
  public onInit(): Promise<void> {
    return super.onInit()
      .then(_ => {
        pnp.setup({
          spfxContext: this.context
        });
      });
  }

  public render(): void {
    this.domElement.innerHTML = `
    <div>
      <div>
        <table border='5' bgcolor='aqua'>
          <tr>
            <td>Please Enter item ID </td>
            <td><input type='text' id='txtID' /></td>
            <td><input type='submit' id='btnRead' value='Read Details'</td>
          </tr>
          <tr>
            <td>Title</td>
            <td><input type='text' id='txtTitle' /></td>
          </tr>
          <tr>
            <td>Nombre</td>
            <td><input type='text' id='txtNombre' /></td>
          </tr>
          <tr>
            <td>Cedula</td>
            <td><input type='text' id='txtCedula' /></td>
          </tr>
          <tr>
            <td>Cargo</td>
            <td>
              <select id="ddlCargo">
                <option value="Microsoft">Microsoft</option>
                <option value="sun">Sun</option>
                <option value="Oracle">Oracle</option>
                <option value="Google">Google</option>
              </select>
            </td>
          </tr>
          <tr>
            <td>Area description</td>
            <td><textarea rows='5' cols='40' id='txtAreaDescription'></textarea></td>
          </tr>
          <tr>
            <td colpsan='2' align='center'>
              <input type='submit' value='Insert Item' id='btnSubmit' />
              <input type='submit' value='Update' id='btnUpdate' />
              <input type='submit' value='Delete' id='btnDelete' />
              <input type='submit' value='Show All Records' id='btnReadAll' />
            </td>
          </tr>
        </table>
        <div id="divStatus" />

        <h2>Get all List Items</h2>
        <hr/>
        <div id="spListData" />
      </div>
    </div>`;

    this._bindEvents();
    this.ReadAllItems();
  }

  public ReadAllItems(): void {
    let html: string = '<table border=1 width=100% style="bordercollapse: collapse;">';
    html += `<th>ID</th> <th>Title</th> <th>Nombre</th> <th>Cedula</th> <th>Cargo</th> <th>Área</th>`;

    pnp.sp.web.lists.getByTitle("Empleados").items.get()
      .then((items: any[]) => {
        items.forEach(function (item) {
          html += `
          <tr>
            <td>${item["ID"]}</td>
            <td>${item["Title"]}</td>
            <td>${item["Nombre"]}</td>
            <td>${item["Cedula"]}</td>
            <td>${item["Cargo"]}</td>
            <td>${item["Area"]}</td>
          </tr>
          `;
        });
        html += `</table>`;
        const allitems: Element = this.domElement.querySelector("#spListData");
        allitems.innerHTML = html;
      })
  }

  private _bindEvents(): void {
    
    this.domElement.querySelector('#btnSubmit').addEventListener('click', () => {
      this.addListItem();
    });

    this.domElement.querySelector('#btnRead').addEventListener('click', () => {
      this.ReadListItem();
    });

    this.domElement.querySelector("#btnUpdate").addEventListener('click', () => {
      this.UpdateListItem();
    });

    this.domElement.querySelector("#btnDelete").addEventListener('click', () => {
      this.DeleteListItem();
    });
  }

  private DeleteListItem(): void {
    
    const id: number = document.getElementById("txtID")["value"];

    pnp.sp.web.lists.getByTitle("Empleados").items.getById(id).delete()
      .then((response => {
        alert("Item has been Delete....");
      }));
  }

  private UpdateListItem(): void {
    let Title = document.getElementById("txtTitle")["value"];
    let Nombre = document.getElementById("txtNombre")["value"];
    let Cedula = document.getElementById("txtCedula")["value"];
    let Cargo = document.getElementById("ddlCargo")["value"];
    let Area = document.getElementById("txtAreaDescription")["value"];

    let id: number = document.getElementById("txtID")["value"];

    pnp.sp.web.lists.getByTitle("Empleados").items.getById(id).update({
      Title: Title,
      Nombre: Nombre,
      Cedula: Cedula,
      Cargo: Cargo,
      Area: Area
    }).then(response => {
      alert("Details Update");
    });
  }

  private ReadListItem(): void {
    const id = document.getElementById("txtID")["value"];

    pnp.sp.web.lists.getByTitle("Empleados").items.getById(id).get()
      .then((item: any) => {
        document.getElementById("txtTitle")["value"] = item["Title"];
        document.getElementById("txtNombre")["value"] = item["Nombre"];
        document.getElementById("txtCedula")["value"] = item["Cedula"];
        document.getElementById("ddlCargo")["value"] = item["Cargo"];
        document.getElementById("txtAreaDescription")["value"] = item["Area"];
      });
  }

  private addListItem(): void {
    
    let Title = document.getElementById("txtTitle")["value"];
    let Nombre = document.getElementById("txtNombre")["value"];
    let Cedula = document.getElementById("txtCedula")["value"];
    let Cargo = document.getElementById("ddlCargo")["value"];
    let Area = document.getElementById("txtAreaDescription")["value"];

    pnp.sp.web.lists.getByTitle("Empleados").items.add({
      Title: Title,
      Nombre: Nombre,
      Cedula: Cedula,
      Cargo: Cargo,
      Area: Area,
    }).then (response => {
      alert("Success");
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
