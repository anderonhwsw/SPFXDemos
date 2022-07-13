import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'MyCommandSetDemoCommandSetStrings';

import * as pnp from 'sp-pnp-js';
import { values } from 'lodash';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IMyCommandSetDemoCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = 'MyCommandSetDemoCommandSet';

export default class MyCommandSetDemoCommandSet extends BaseListViewCommandSet<IMyCommandSetDemoCommandSetProperties> {

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized MyCommandSetDemoCommandSet');
    return Promise.resolve();
  }

  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.
      compareOneCommand.visible = event.selectedRows.length === 1;
    }

    const compareTwoCommand: Command = this.tryGetCommand('COMMAND_2');
    if (compareTwoCommand) {
      // This command should be hidden unless exactly one row is selected.
      compareTwoCommand.visible = event.selectedRows.length > 1;
    }

    const comparethreeCommand: Command = this.tryGetCommand('COMMAND_3');
    if (comparethreeCommand) {
      // This command should be hidden unless exactly one row is selected.
      comparethreeCommand.visible = event.selectedRows.length > 1;
    }
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'COMMAND_1':
        //Dialog.alert(`${this.properties.sampleTextOne}`);
        let title: string = event.selectedRows[0].getValueByName("Title");
        let Status: string = event.selectedRows[0].getValueByName("Status");
        let User: any = event.selectedRows[0].getValueByName("Usuario");
        let Observacion: string = event.selectedRows[0].getValueByName("Observacion");
        console.log(User);
        Dialog.alert(`Project Name: ${title} - Current Status: ${Status}% done
        <br/>
        Observacion: ${Observacion}
        <br/>
        Assigned: ${User[title]}`);

        break;
      case 'COMMAND_2':
        Dialog.alert(`${this.properties.sampleTextTwo}`);
        break;

      case 'COMMAND_3':
        Dialog.prompt(`Project Status Remarks`)
          .then((value: string) => {
            this.UpdateRemarks(event.selectedRows, value)
          })
        break;

      default:
        throw new Error('Unknown command');
    }
  }

  private UpdateRemarks(items: any, value: string) {
    let batch = pnp.sp.createBatch();
    items.forEach(item => {
      pnp.sp.web.lists.getByTitle("ProjectsStatus").items.
        getById(item.getValueByName("ID")).inBatch(batch).update({ Remarks: value})
          .then(response => {
          });
    });

    batch.execute().then(response => {
      location.reload();
    })
  }
}
