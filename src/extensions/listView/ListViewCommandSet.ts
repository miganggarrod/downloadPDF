import { override } from '@microsoft/decorators';
// import { Log } from '@microsoft/sp-core-library';
import { BaseListViewCommandSet, Command, IListViewCommandSetListViewUpdatedParameters, IListViewCommandSetExecuteEventParameters} from '@microsoft/sp-listview-extensibility';
// import { Dialog } from '@microsoft/sp-dialog';

import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

import jsPDF from 'jspdf';

// import * as strings from 'ListViewCommandSetStrings';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IListViewCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = 'ListViewCommandSet';

export default class ListViewCommandSet extends BaseListViewCommandSet<IListViewCommandSetProperties> {

  @override
  public onInit(): Promise<void> {
    // Log.info(LOG_SOURCE, 'Initialized ListViewCommandSet');
    return Promise.resolve();
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.
      compareOneCommand.visible = event.selectedRows.length === 1;
    }
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'COMMAND_1':
      this.GeneraPdf(event.selectedRows[0].getValueByName('ID'));
        break;
      case 'COMMAND_2':
        console.log(this.properties.sampleTextTwo);
        break;
      default:
        throw new Error('Unknown command');
    }
  }
  private async GeneraPdf(IdItem :number) {
  let item  =  await sp.web.lists.getByTitle('SPFx').items.getById(IdItem).get();
  console.log(item);
  var doc = new jsPDF();
  doc.text('Tipo:',10,10);
  doc.text(item.Type, 50, 10);
  doc.text('Titulo:',10,20);
  doc.text(item.Title, 50, 20);
  doc.text('Descripción:',10,30);
  doc.text(item.Description, 50, 30);
  doc.save('Pdf.pdf');
 
  }
}
