import {
  BaseListViewCommandSet,
  type Command,
  type IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog, BaseDialog } from '@microsoft/sp-dialog';
import { SPHttpClient } from '@microsoft/sp-http';

export interface IButtonDialogBoxCommandSetProperties {
  sampleTextOne: string;
  sampleTextTwo: string;
}

export default class ButtonDialogBoxCommandSet
  extends BaseListViewCommandSet<IButtonDialogBoxCommandSetProperties> {

  public onInit(): Promise<void> {
    const cmd: Command = this.tryGetCommand('SHOW_VERSIONS');
    cmd.visible = false;
    this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);
    return Promise.resolve();
  }

  private _onListViewStateChanged(): void {
    const cmd: Command = this.tryGetCommand('SHOW_VERSIONS');
    cmd.visible = !!this.context.listView.selectedRows?.length;
    this.raiseOnChange();
  }

  public onExecute(e: IListViewCommandSetExecuteEventParameters): void {
    if (e.itemId !== 'SHOW_VERSIONS') { return; }

    const sel = this.context.listView.selectedRows?.[0];
    if (!sel) {
      Dialog.alert('Please select an item.');
      return;
    }

    const list = this.context.pageContext.list;
    if (!list) {
      Dialog.alert('List context not found.');
      return;
    }

    const listId: string = list.id.toString();
    const itemId: number = parseInt(sel.getValueByName('ID'), 10);

    this._showVersionHistory(listId, itemId);
  }

  private async _showVersionHistory(listId: string, itemId: number): Promise<void> {
    const webUrl = this.context.pageContext.web.absoluteUrl;

    const endpoint = `${webUrl}/_api/web/lists(guid'${listId}')/items(${itemId})/versions` +
      `?$select=VersionId,VersionLabel,Modified,Editor/LookupValue,*&$expand=Editor`;

    try {
      const r = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1,
        { headers: { Accept: 'application/json;odata=nometadata', 'odata-version': '' } }
      );

      const versions: any[] = (await r.json()).value;
      if (!versions?.length) {
        Dialog.alert(`No version history found for item ID ${itemId}.`);
        return;
      }

      versions.sort((a, b) => b.VersionId - a.VersionId);
      const dialog = new VersionHistoryDialog(versions, itemId);
      dialog.show().then(() => {
      // Optional: You can log or track when dialog closes
      });

    } catch (err) {
      Dialog.alert(`Error fetching version history: ${(err as Error).message}`);
    }
  }
}

class VersionHistoryDialog extends BaseDialog {

  constructor(private _versions: any[], private _itemId: number) {
    super();
  }

  // Prevent closing by clicking outside
  public get isBlocking(): boolean {
    return true;
  }

protected onAfterClose(): void {
  if (this.domElement) {
    // Remove all child nodes to kill all event listeners
    while (this.domElement.firstChild) {
      this.domElement.removeChild(this.domElement.firstChild);
    }

    // Also remove from DOM tree
    if (this.domElement.parentElement) {
      this.domElement.parentElement.removeChild(this.domElement);
    }
  }

  // Clean internal references
  (this as any)._versions = [];
  (this as any)._itemId = null;

  super.onAfterClose();
}

public render(): void {
  // Clear the previous content explicitly — this ensures the input is truly refreshed
  if (this.domElement) {
    this.domElement.innerHTML = '';
  }

  // Now render freshly
  this._renderDialogContent();
}


  private _renderDialogContent(fromVal?: string, toVal?: string, searchVal?: string): void {
    const html = `
      <div style="width:800px;padding:20px;max-height:600px;overflow-y:auto;font-family:Segoe UI;font-size:14px;">
        <h2>📄 Version History for Item ID ${this._itemId}</h2>

        <div style="margin-bottom:15px;display:flex;gap:10px;align-items:center;">
          <label>From:
            <input type="date" id="fromDate" value="${fromVal || ''}" style="padding:5px;">
          </label>
          <label>To:
            <input type="date" id="toDate" value="${toVal || ''}" style="padding:5px;">
          </label>
          <label>Search:
            <input type="text" id="searchBox" value="${searchVal || ''}" placeholder="Type to filter..." style="padding:5px;width:200px;">
          </label>
          <button id="resetButton" style="padding:5px 12px;">Reset</button>
          <button id="closeButton" style="padding:5px 12px;">Close</button>
        </div>

        <table style="width:100%;border-collapse:collapse;border:1px solid #ddd;">
          <thead>
            <tr style="background:#f4f4f4;">
              <th style="padding:8px;border:1px solid #ddd;">No.</th>
              <th style="padding:8px;border:1px solid #ddd;">Modified</th>
              <th style="padding:8px;border:1px solid #ddd;">Modified By</th>
            </tr>
          </thead>
          <tbody></tbody>
        </table>
      </div>`;

    this.domElement.innerHTML = html;

    this._updateTableBody(this._versions);

    this.domElement.querySelector('#fromDate')
      ?.addEventListener('change', this._handleFiltersChange.bind(this));
    this.domElement.querySelector('#toDate')
      ?.addEventListener('change', this._handleFiltersChange.bind(this));
    this.domElement.querySelector('#searchBox')
      ?.addEventListener('input', this._handleFiltersChange.bind(this));
    this.domElement.querySelector('#resetButton')
      ?.addEventListener('click', () => this._renderDialogContent());
    this.domElement.querySelector('#closeButton')
      ?.addEventListener('click', () => this.close());
  }

  private _handleFiltersChange(): void {
    const fInput = this.domElement.querySelector('#fromDate') as HTMLInputElement;
    const tInput = this.domElement.querySelector('#toDate') as HTMLInputElement;
    const searchInput = this.domElement.querySelector('#searchBox') as HTMLInputElement;

    const f = fInput?.value;
    const t = tInput?.value;
    const searchText = searchInput?.value.toLowerCase();

    const from = f ? new Date(f) : null;
    const to = t ? new Date(+new Date(t) + 86_399_999) : null;

    const subset = this._versions.filter(v => {
      const modified = new Date(v.Modified);
      const dateMatch = (!from || modified >= from) && (!to || modified <= to);
      if (!dateMatch) return false;

      if (!searchText) return true;

      const label = (v.VersionLabel || '').toString().toLowerCase();
      const editor = (v.Editor?.LookupValue || '').toLowerCase();
      const modifiedStr = modified.toLocaleString().toLowerCase();
      const otherValues = Object.keys(v)
        .map(key => v[key])
        .filter((val: unknown) => typeof val === 'string' || typeof val === 'number')
        .map(val => val!.toString().toLowerCase())
        .join(' ');

      return (
        label.includes(searchText) ||
        editor.includes(searchText) ||
        modifiedStr.includes(searchText) ||
        otherValues.includes(searchText)
      );
    });

    this._updateTableBody(subset);
  }

  private _updateTableBody(filtered: any[]): void {
    const tbody = this.domElement.querySelector('tbody');
    if (!tbody) return;

    const excluded: string[] = [
      'ItemChildCount','FolderChildCount','NoExecute','ContentVersion','VersionLabel','Editor','__metadata',
      'ContentTypeId','GUID','Attachments','FileRef','FileDirRef','FileLeafRef','Created','Author',
      'OData__UIVersionString','OData__ModerationStatus','FileSystemObjectType','ID','ScopeId','UniqueId',
      'ParentUniqueId','FSObjType','Order','WorkflowVersion','owshiddenversion','VersionId','IsCurrentVersion',
      'SMLastModifiedDate','SMTotalFileStreamSize','Last_x005f_x0020_x005f_Modified'
    ];

    let html = '';

    filtered.forEach(v => {
      const label = v.VersionLabel;
      const modified = new Date(v.Modified).toLocaleString();
      const editor = v.Editor?.LookupValue ?? '—';

      html += `
        <tr>
          <td style="padding:8px;border:1px solid #ddd;">${label}</td>
          <td style="padding:8px;border:1px solid #ddd;">${modified}</td>
          <td style="padding:8px;border:1px solid #ddd;">${editor}</td>
        </tr>
        <tr>
          <td colspan="3" style="padding:8px;border:1px solid #ddd;">`;

      Object.keys(v).forEach(key => {
        if (excluded.includes(key) || key.startsWith('OData_') || key.startsWith('ows')) return;
        const val = v[key];
        if (val !== null && val !== '' && typeof val !== 'object') {
          const displayKey = (key === 'Created_x005f_x0020_x005f_Date') ? 'Created' : key;
          html += `• <strong>${displayKey}:</strong> ${val}<br>`;
        }
      });

      html += `</td></tr>`;
    });

    tbody.innerHTML = html;
  }
}
