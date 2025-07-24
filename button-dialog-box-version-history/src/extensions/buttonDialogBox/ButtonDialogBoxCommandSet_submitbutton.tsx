// IMPORTS
import {
  BaseListViewCommandSet,
  type Command,
  type IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog, BaseDialog } from '@microsoft/sp-dialog';
import { SPHttpClient } from '@microsoft/sp-http';

// PROPERTIES
export interface IButtonDialogBoxCommandSetProperties {
  sampleTextOne: string;
  sampleTextTwo: string;
}

// MAIN CLASS
export default class ButtonDialogBoxCommandSet extends BaseListViewCommandSet<IButtonDialogBoxCommandSetProperties> {

  public onInit(): Promise<void> {
    const command: Command = this.tryGetCommand('SHOW_VERSIONS');
    command.visible = false;
    this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);
    return Promise.resolve();
  }

  private _onListViewStateChanged(): void {
    const command: Command = this.tryGetCommand('SHOW_VERSIONS');
    command.visible = !!this.context.listView.selectedRows?.length;
    this.raiseOnChange();
  }

  public onExecute(e: IListViewCommandSetExecuteEventParameters): void {
    if (e.itemId !== 'SHOW_VERSIONS') return;

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

    const listId = list.id.toString();
    const itemId = parseInt(sel.getValueByName('ID'));
    this._showVersionHistory(listId, itemId);
  }

  private async _showVersionHistory(listId: string, itemId: number): Promise<void> {
    const webUrl = this.context.pageContext.web.absoluteUrl;

    // ─── SELECT “ID” (the ordinal) so that v.ID gives the correct version number ───
    const endpoint =
      `${webUrl}/_api/web/lists(guid'${listId}')/items(${itemId})/versions` +
      `?$select=ID,VersionLabel,Modified,*`;

    try {
      const r = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=verbose',
            'odata-version': ''
          }
        }
      );

      const versions = (await r.json()).d?.results;
      if (!versions?.length) {
        Dialog.alert(`No version history found for item ID ${itemId}.`);
        return;
      }

      versions.sort((a: any, b: any) =>
        parseFloat(b.VersionLabel) - parseFloat(a.VersionLabel)
      );

      new VersionHistoryDialog(
        versions,
        itemId,
        listId,
        this.context.spHttpClient,
        webUrl
      ).show();
    } catch (err) {
      Dialog.alert(`Error fetching version history: ${(err as Error).message}`);
    }
  }
}

// ---------------- DIALOG CLASS ----------------
class VersionHistoryDialog extends BaseDialog {
constructor(
  private _versions: any[],
  private _itemId: number,
  private _listId: string,
  private _spHttpClient: SPHttpClient,
  private _web: string
) { super(); }

  public render(): void {
    this._renderDialogContent(this._versions);
  }

  private _renderDialogContent(filtered: any[], fromVal?: string, toVal?: string): void {
    const excluded = [
      'ItemChildCount','FolderChildCount','NoExecute','ContentVersion','VersionLabel','Editor','__metadata',
      'ContentTypeId','GUID','Attachments','FileRef','FileDirRef','FileLeafRef','Created','Author',
      'OData__UIVersionString','OData__ModerationStatus','FileSystemObjectType','ID','ScopeId','UniqueId',
      'ParentUniqueId','FSObjType','Order','WorkflowVersion','owshiddenversion','VersionId','IsCurrentVersion',
      'SMLastModifiedDate','SMTotalFileStreamSize','Last_x005f_x0020_x005f_Modified'
    ];

    let html = `
      <div style="width:800px;padding:20px;max-height:600px;overflow-y:auto;font-family:Segoe UI;font-size:14px;">
        <h2>📄 Version History for Item ID ${this._itemId}</h2>
        <div style="margin-bottom:15px;display:flex;gap:10px;">
          <label>From: <input type="date" id="fromDate" value="${fromVal || ''}" style="padding:5px;"></label>
          <label>To:   <input type="date" id="toDate"   value="${toVal   || ''}" style="padding:5px;"></label>
          <button id="filterButton" style="padding:5px 12px;">Submit</button>
          <button id="resetButton"  style="padding:5px 12px;">Reset</button>
        </div>

        <table style="width:100%;border-collapse:collapse;border:1px solid #ddd;">
          <thead>
            <tr style="background:#f4f4f4;">
              <th style="padding:8px;border:1px solid #ddd;">No.</th>
              <th style="padding:8px;border:1px solid #ddd;">Modified</th>
              <th style="padding:8px;border:1px solid #ddd;">Modified By</th>
            </tr>
          </thead>
          <tbody>`;

    filtered.forEach((v: any) => {
      const label     = v.VersionLabel;
      const safeLabel = label.replace(/\./g, '-');   // e.g. "2.0" → "2-0"
      const mod       = new Date(v.Modified).toLocaleString();

      html += `
        <tr>
          <td style="padding:8px;border:1px solid #ddd;">${label}</td>
          <td style="padding:8px;border:1px solid #ddd;">${mod}</td>
          <td id="modifiedBy-${safeLabel}" style="padding:8px;border:1px solid #ddd;">Loading...</td>
        </tr>
        <tr><td colspan="3" style="padding:8px;border:1px solid #ddd;">`;

      Object.keys(v).forEach(key => {
        if (excluded.includes(key) || key.startsWith('OData_') || key.startsWith('ows')) {
          return;
        }
        const val = v[key];
        if (val !== null && val !== '' && typeof val !== 'object') {
          const displayKey = (key === 'Created_x005f_x0020_x005f_Date') ? 'Created' : key;
          html += `• <strong>${displayKey}:</strong> ${val}<br>`;
        }
      });

      html += '</td></tr>';
    });

    html += '</tbody></table></div>';
    this.domElement.innerHTML = html;

    // ─── ② PASS the ordinal v.ID (not VersionId) here ───
    filtered.forEach((v: any) => {
      const safeLabel = v.VersionLabel.replace(/\./g, '-');
      this._loadFieldValuesAsText(v.ID, safeLabel);
    });

    // Submit (filter) handler
    this.domElement.querySelector('#filterButton')
      ?.addEventListener('click', () => {
        const f    = (this.domElement.querySelector('#fromDate') as HTMLInputElement).value;
        const t    = (this.domElement.querySelector('#toDate')   as HTMLInputElement).value;
        const from = f ? new Date(f) : null;
        const to   = t ? new Date(t) : null;

        const subset = this._versions.filter(item => {
          const d = new Date(item.Modified);
          return (!from || d >= from) && (!to || d <= new Date(+new Date(to) + 86399999));
        });

        this._renderDialogContent(subset, f, t);
      });

    // Reset handler
    this.domElement.querySelector('#resetButton')
      ?.addEventListener('click', () => this._renderDialogContent(this._versions));
  }

private async _loadFieldValuesAsText(versionId: number, label: string): Promise<void> {
  const url = `${this._web}/_api/web/lists(guid'${this._listId}')/items(${this._itemId})/versions(${versionId})/FieldValuesAsText`;

  try {
    const r = await this._spHttpClient.get(url, SPHttpClient.configurations.v1, {
      headers: {
        'Accept': 'application/json;odata=verbose',
        'odata-version': ''
      }
    });

    const d = (await r.json()).d;

    let editorName = 'Unknown';
    if (d?.Editor?.LookupValue) {
      editorName = d.Editor.LookupValue;
    } else if (d?.EditorId) {
      editorName = `UserId ${d.EditorId}`;
    }

    const cell = this.domElement.querySelector(`#modifiedBy-${label}`);
    if (cell) {
      cell.textContent = editorName;
    }

    console.log(`✅ Version ${versionId} Editor: ${editorName}`);

  } catch (e) {
    console.log(`❌ FieldValuesAsText call for version ${versionId} failed:`, e);
  }
}
}