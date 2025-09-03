import { Log } from '@microsoft/sp-core-library';
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
export interface ISendForSignatureCommandCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = 'SendForSignatureCommandCommandSet';

export default class SendForSignatureCommandCommandSet extends BaseListViewCommandSet<ISendForSignatureCommandCommandSetProperties> {

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized SendForSignatureCommandCommandSet');

    // initial state of the command's visibility
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    compareOneCommand.visible = false;

    this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);

    return Promise.resolve();
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'COMMAND_1':
        const selectedRows = this.context.listView.selectedRows;
        if (selectedRows && selectedRows.length === 1) {
          // Récupérer l'URL du fichier dans la colonne FileRef (chemin relatif)
          const fileRef = selectedRows[0].getValueByName('FileRef') as string;

          // Construire l'URL complète
          const siteUrl = this.context.pageContext.web.absoluteUrl;
          const fileUrl = `${siteUrl}${fileRef}`;

          // Rediriger vers ta page contenant le webpart, en passant fileUrl en query string
          const targetPage = `${siteUrl}/SitePages/Demande-Signature.aspx?fileUrl=${encodeURIComponent(fileUrl)}`;
          window.location.href = targetPage;
        } else {
          Dialog.alert('Veuillez sélectionner un seul fichier.').catch(() => { });
        }
        break;
      // case 'COMMAND_2':
      //   Dialog.alert(`${this.properties.sampleTextTwo}`).catch(() => {
      //     /* handle error */
      //   });
      //   break;
      default:
        throw new Error('Unknown command');
    }
  }

  private _onListViewStateChanged = (args: ListViewStateChangedEventArgs): void => {
    Log.info(LOG_SOURCE, 'List view state changed');

    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.
      compareOneCommand.visible = this.context.listView.selectedRows?.length === 1;
    }

    // TODO: Add your logic here 

    // You should call this.raiseOnChage() to update the command bar
    this.raiseOnChange();
  }
}
