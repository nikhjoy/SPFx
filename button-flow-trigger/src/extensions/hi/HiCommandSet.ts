import {
  BaseListViewCommandSet,
  type Command,
  type IListViewCommandSetExecuteEventParameters,
  type ListViewStateChangedEventArgs
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IHiCommandSetProperties {
  sampleTextOne: string;
  sampleTextTwo: string;
}

export default class HiCommandSet extends BaseListViewCommandSet<IHiCommandSetProperties> {

  public onInit(): Promise<void> {
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    if (compareOneCommand) {
      compareOneCommand.visible = false;
    }

    this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);

    return Promise.resolve(); // ✅ Only return after everything is done
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'MARK_CUSTOMER': {
        const selectedRow = event.selectedRows[0];
        const itemId = selectedRow.getValueByName("ID");

        const flowUrl = "https://prod-20.westus.logic.azure.com:443/workflows/7ecc7479a75a43f18b72b65ec2c7b312/triggers/manual/paths/invoke?api-version=2016-06-01";

        fetch(flowUrl, {
          method: "POST",
          headers: {
            "Content-Type": "application/json"
          },
          body: JSON.stringify({ itemId })
        })
          .then(response => {
            if (!response.ok) {
              throw new Error("Flow call failed");
            }
            return response.text();
          })
          .then(result => {
            Dialog.alert("Flow triggered successfully!");
          })
          .catch(error => {
            Dialog.alert(`Error: ${error.message}`);
          });

        break;
      }
      default:
        throw new Error(`Unknown command: ${event.itemId}`);
    }
  }

  private _onListViewStateChanged = (args: ListViewStateChangedEventArgs): void => {
    const myCommand: Command = this.tryGetCommand('MARK_CUSTOMER');
    if (myCommand) {
      myCommand.visible = this.context.listView.selectedRows?.length === 1;
    }
    this.raiseOnChange();
  }
}
