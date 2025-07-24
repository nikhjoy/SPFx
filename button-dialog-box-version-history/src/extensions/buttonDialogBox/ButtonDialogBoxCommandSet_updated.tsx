import {
  BaseListViewCommandSet,
  type Command,
  type IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog, BaseDialog } from '@microsoft/sp-dialog';
import { SPHttpClient } from '@microsoft/sp-http';

/* -------------------------------------------------------------------------- */
/*  Component properties (from manifest)                                      */
/* -------------------------------------------------------------------------- */
export interface IButtonDialogBoxCommandSetProperties {
  sampleTextOne: string;
  sampleTextTwo: string;
}

/* -------------------------------------------------------------------------- */
/*  Main command-set class                                                    */
/* -------------------------------------------------------------------------- */
export default class ButtonDialogBoxCommandSet
  extends BaseListViewCommandSet<IButtonDialogBoxCommandSetProperties> {

  /* Hide button until selection exists */
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

  /* ---------------------------------------------------------------------- */
  /*  Execute                                                               */
  /* ---------------------------------------------------------------------- */
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

  /* ---------------------------------------------------------------------- */
  /*  Fetch versions – single REST call                                     */
  /* ---------------------------------------------------------------------- */
  private async _showVersionHistory(listId: string, itemId: number): Promise<void> {
    const webUrl = this.context.pageContext.web.absoluteUrl;

    /* Select the columns we need and expand the Editor lookup */
    const endpoint = `${webUrl}/_api/web/lists(guid'${listId}')/items(${itemId})/versions` +
                     `?$select=VersionId,VersionLabel,Modified,Editor/LookupValue,*` +
                     `&$expand=Editor`;

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

      /* Sort newest → oldest using VersionId (numeric) */
      versions.sort((a, b) => b.VersionId - a.VersionId);

      /* Open dialog */
      new VersionHistoryDialog(versions, itemId).show();

    } catch (err) {
      Dialog.alert(`Error fetching version history: ${(err as Error).message}`);
    }
  }
}

/* -------------------------------------------------------------------------- */
/*  Dialog class                                                              */
/* -------------------------------------------------------------------------- */
class VersionHistoryDialog extends BaseDialog {

  constructor(private _versions: any[], private _itemId: number) { super(); }

  public render(): void {
    this._renderDialogContent(this._versions);
  }

  /* Re-renders the dialog (initial + every filter change) */
  private _renderDialogContent(filtered: any[], fromVal?: string, toVal?: string): void {
    /* System/internal columns to skip */
    const excluded: string[] = [
      'ItemChildCount','FolderChildCount','NoExecute','ContentVersion','VersionLabel','Editor','__metadata',
      'ContentTypeId','GUID','Attachments','FileRef','FileDirRef','FileLeafRef','Created','Author',
      'OData__UIVersionString','OData__ModerationStatus','FileSystemObjectType','ID','ScopeId','UniqueId',
      'ParentUniqueId','FSObjType','Order','WorkflowVersion','owshiddenversion','VersionId','IsCurrentVersion',
      'SMLastModifiedDate','SMTotalFileStreamSize','Last_x005f_x0020_x005f_Modified'
    ];

    /* ---------- Build HTML ---------- */
    let html = `
      <div style="width:800px;padding:20px;max-height:600px;overflow-y:auto;font-family:Segoe UI;font-size:14px;">
        <h2>📄 Version History for Item ID ${this._itemId}</h2>

        <div style="margin-bottom:15px;display:flex;gap:10px;">
          <label>From:
            <input type="date" id="fromDate" value="${fromVal || ''}" style="padding:5px;">
          </label>
          <label>To:
            <input type="date" id="toDate" value="${toVal || ''}" style="padding:5px;">
          </label>
          <button id="resetButton" style="padding:5px 12px;">Reset</button>
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

    /* Version rows */
    filtered.forEach(v => {
      const label     = v.VersionLabel;
      const modified  = new Date(v.Modified).toLocaleString();
      const editor    = v.Editor?.LookupValue ?? '—';

      html += `
        <tr>
          <td style="padding:8px;border:1px solid #ddd;">${label}</td>
          <td style="padding:8px;border:1px solid #ddd;">${modified}</td>
          <td style="padding:8px;border:1px solid #ddd;">${editor}</td>
        </tr>
        <tr>
          <td colspan="3" style="padding:8px;border:1px solid #ddd;">`;

      /* Property list (filtered) */
      Object.keys(v).forEach(key => {
        if (excluded.includes(key) || key.startsWith('OData_') || key.startsWith('ows')) { return; }

        const val = v[key];
        if (val !== null && val !== '' && typeof val !== 'object') {
          const displayKey = (key === 'Created_x005f_x0020_x005f_Date') ? 'Created' : key;
          html += `• <strong>${displayKey}:</strong> ${val}<br>`;
        }
      });

      html += `</td></tr>`;
    });

    html += '</tbody></table></div>';
    this.domElement.innerHTML = html;

    /* ---------- Wire up date filter events ---------- */
    this.domElement.querySelector('#fromDate')
      ?.addEventListener('change', this._handleDateChange.bind(this));

    this.domElement.querySelector('#toDate')
      ?.addEventListener('change', this._handleDateChange.bind(this));

    this.domElement.querySelector('#resetButton')
      ?.addEventListener('click', () => this._renderDialogContent(this._versions));
  }

  /* Date-range filter */
  private _handleDateChange(): void {
    const f = (this.domElement.querySelector('#fromDate') as HTMLInputElement).value;
    const t = (this.domElement.querySelector('#toDate')   as HTMLInputElement).value;
    const from = f ? new Date(f) : null;
    const to   = t ? new Date(t) : null;

    const subset = this._versions.filter(v => {
      const d = new Date(v.Modified);
      return (!from || d >= from) &&
             (!to   || d <= new Date(+new Date(to) + 86_399_999)); // include entire “to” day
    });

    this._renderDialogContent(subset, f, t);
  }
}