import * as React from 'react';
import { IFetchDataProps } from './IFetchDataProps';
import styles from './FetchData.module.scss';
import {
  TextField,
  Dropdown,
  IDropdownOption,
  DatePicker,
  PrimaryButton,
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  IconButton,
  Panel,
  PanelType
} from '@fluentui/react';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
import { DefaultButton } from '@fluentui/react/lib/Button';
import { sp } from '@pnp/sp/presets/all';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

export interface IEmployee {
  id: number;
  title: string;
  name: string;
  age: number;
  departmentId: number;
  department: string;
  doj: Date | null;
  email: string;
  mobileNo: string;
  managerIds: number[];
  managerNames: string;
  jsonValue?: string;
}

type Mode = 'new' | 'edit' | 'view' | null;

interface IState {
  items: IEmployee[];
  departments: IDropdownOption[];
  users: Array<{ Id: number; Title: string; Email: string }>;
  currentItem: Partial<IEmployee>;
  currentManagerEmails: string[];
  showModal: boolean;
  modalMode: Mode;
  isLoading: boolean;
  showConfirmDialog: boolean;
  deleteItemId: number | null;
  showDetailDialog: boolean;
}

export default class EmployeeDashboard extends React.Component<IFetchDataProps, IState> {
  private _columns: IColumn[] = [];

  constructor(props: IFetchDataProps) {
    super(props);
    this.state = {
      items: [],
      departments: [],
      users: [],
      currentItem: {},
      currentManagerEmails: [],
      showModal: false,
      modalMode: null,
      isLoading: true,
      showConfirmDialog: false,
      deleteItemId: null,
      showDetailDialog: false
    };
  }

  public async componentDidMount(): Promise<void> {
    this.setState({ isLoading: true });
    await this.loadDepartments();
    await this.loadUsers();
    await this.loadAllItems();
    this.setState({ isLoading: false });
  }

  private loadDepartments = async (): Promise<void> => {
    const depts = await sp.web.lists.getByTitle('Department').items.select('ID','Title').get();
    this.setState({
      departments: depts.map(d => ({ key: d.ID, text: d.Title }))
    });
  };

  private loadUsers = async (): Promise<void> => {
    const users = await sp.web.siteUsers.select('Id','Title','Email').get();
    this.setState({ users });
  };

  private loadAllItems = async (): Promise<void> => {
    const data: any[] = await sp.web.lists.getByTitle('Employee').items
      .select(
        'ID','Title','EmpName','Age',
        'Department/ID','Department/Title',
        'DOJ','Email','MobileNo','JSONValue',
        'ManagerId','Manager/Id','Manager/Title'
      )
      .expand('Department','Manager')
      .get();

    const items = data.map(r => {
      const managerIds: number[] = Array.isArray(r.ManagerId) ? r.ManagerId : [r.ManagerId];
      const rawMgr = r.Manager;
      let managerNamesArr: string[] = [];
      if (Array.isArray(rawMgr)) {
        managerNamesArr = rawMgr.map((m: any) => m.Title);
      } else if (rawMgr && (rawMgr as any).Title) {
        managerNamesArr = [(rawMgr as any).Title];
      }

      return {
        id: r.ID,
        title: r.Title,
        name: r.EmpName,
        age: r.Age,
        departmentId: r.Department?.ID,
        department: r.Department?.Title || '',
        doj: r.DOJ ? new Date(r.DOJ) : null,
        email: r.Email,
        mobileNo: r.MobileNo,
        managerIds,
        managerNames: managerNamesArr.join(', '),
        jsonValue: r.JSONValue
      };
    });

    this.setState({ items }, this.buildColumns);
  };

  private buildColumns = () => {
    this._columns = [
      { key:'col1', name:'ID', fieldName:'id', minWidth:50 },
      { key:'col2', name:'Title', fieldName:'title', minWidth:100 },
      { key:'col3', name:'Name', fieldName:'name', minWidth:100 },
      { key:'col4', name:'Age', fieldName:'age', minWidth:50 },
      { key:'col5', name:'Department', fieldName:'department', minWidth:100 },
      { key:'col6', name:'DOJ', fieldName:'doj', minWidth:100, onRender: i => i.doj?.toLocaleDateString() },
      { key:'col7', name:'Email', fieldName:'email', minWidth:150 },
      { key:'col8', name:'Mobile', fieldName:'mobileNo', minWidth:100 },
      { key:'col9', name:'Managers', fieldName:'managerNames', minWidth:150 },
      {
        key:'col10', name:'Actions', fieldName:'actions', minWidth:200,
        onRender: item => (
          <div>
            <IconButton iconProps={{ iconName:'View' }} onClick={() => this.openModal('view', item)} />
            <IconButton iconProps={{ iconName:'Edit' }} onClick={() => this.openModal('edit', item)} />
            <IconButton iconProps={{ iconName:'Delete' }} onClick={() => this.setState({ showConfirmDialog:true, deleteItemId:item.id })} />
            <IconButton
            iconProps={{ iconName: 'AddMedium' }}
            onClick={() => this.setState({ currentItem: item, showDetailDialog: true })}
            styles={{ root: { color: '#5c2d91', marginLeft: 4 } }}
            title="View Full Details"
            />
          </div>
        )
      }
    ];
  };

  private openModal = async (mode: Mode, item?: IEmployee) => {
    let managerEmails: string[] = [];
    if ((mode === 'edit' || mode === 'view') && item?.managerIds?.length) {
      if (this.state.users.length === 0) {
        await this.loadUsers();
      }
      managerEmails = this.state.users
        .filter(u => item.managerIds!.includes(u.Id))
        .map(u => u.Email);
    }

    this.setState({
      modalMode: mode,
      currentItem: mode !== 'new' && item ? { ...item } : {},
      currentManagerEmails: managerEmails,
      showModal: true
    });
  };

  private closeModal = () => {
    this.setState({
      showModal: false,
      modalMode: null,
      currentItem: {},
      currentManagerEmails: []
    });
  };

  private saveItem = async () => {
    const { currentItem, modalMode, currentManagerEmails } = this.state;
    if (!currentItem.departmentId || !currentItem.doj || currentManagerEmails.length === 0) {
      return;
    }

    const ensured = await Promise.all(currentManagerEmails.map(email => sp.web.ensureUser(email)));
    const managerIds = ensured.map(r => r.data.Id);

    const payload = {
      Title: currentItem.title,
      EmpName: currentItem.name,
      Age: currentItem.age,
      DepartmentId: currentItem.departmentId,
      DOJ: currentItem.doj.toISOString(),
      Email: currentItem.email,
      MobileNo: currentItem.mobileNo,
      ManagerId: { results: managerIds }
    };

    if (modalMode === 'edit') {
      await sp.web.lists.getByTitle('Employee').items.getById(currentItem.id!).update(payload);
    } else {
      await sp.web.lists.getByTitle('Employee').items.add(payload);
    }

    this.closeModal();
    this.loadAllItems();
  };

  public render(): React.ReactElement {
    const {
      items, departments, showModal, modalMode, currentItem,
      currentManagerEmails, showConfirmDialog, deleteItemId    } = this.state;
    const isView = modalMode === 'view';

    return (
      <div>
        <div style={{ display:'flex', justifyContent:'flex-end', marginBottom: 16 }}>
          <PrimaryButton text="Add New" iconProps={{ iconName:'Add' }} onClick={() => this.openModal('new')} />
        </div>
        <DetailsList
          items={items}
          columns={this._columns}
          layoutMode={DetailsListLayoutMode.justified}
          onItemInvoked={item => this.openModal('view', item)}
        />

        <Panel isOpen={showModal} onDismiss={this.closeModal} type={PanelType.medium} headerText={modalMode === 'edit' ? 'Edit Employee' : modalMode === 'view' ? 'View Employee' : 'New Employee'} isBlocking>
          <TextField label="Title" value={currentItem.title || ''} readOnly={isView} onChange={(_, v) => this.setState({ currentItem:{ ...currentItem, title:v } })} />
          <TextField label="Name" value={currentItem.name || ''} readOnly={isView} onChange={(_, v) => this.setState({ currentItem:{ ...currentItem, name:v } })} />
          <TextField label="Age" value={currentItem.age?.toString() || ''} readOnly={isView} onChange={(_, v) => this.setState({ currentItem:{ ...currentItem, age: parseInt(v||'0') } })} />
          <Dropdown label="Department" options={departments} selectedKey={currentItem.departmentId} disabled={isView} onChange={(_, o) => this.setState({ currentItem:{ ...currentItem, departmentId: o?.key as number } })} />
          <DatePicker label="Date of Joining" value={currentItem.doj || undefined} disabled={isView} onSelectDate={d => this.setState({ currentItem:{ ...currentItem, doj:d||null } })} />
          <TextField label="Email" value={currentItem.email || ''} readOnly={isView} onChange={(_, v) => this.setState({ currentItem:{ ...currentItem, email:v } })} />
          <TextField label="Mobile No" value={currentItem.mobileNo || ''} readOnly={isView} onChange={(_, v) => this.setState({ currentItem:{ ...currentItem, mobileNo:v } })} />

          <label style={{ fontWeight:600, margin:'12px 0 4px' }}>Managers</label>
          <PeoplePicker
            context={{
              spHttpClient: this.props.context.spHttpClient,
              msGraphClientFactory: this.props.context.msGraphClientFactory,
              absoluteUrl: this.props.context.pageContext.web.absoluteUrl
            }}
            personSelectionLimit={5}
            showtooltip
            required
            principalTypes={[PrincipalType.User]}
            defaultSelectedUsers={currentManagerEmails}
            onChange={items => {
              const emails = items.map(i => i.secondaryText).filter((e): e is string => !!e);
              this.setState({ currentManagerEmails: emails });
            }}
            resolveDelay={200}
            disabled={isView}
          />

          {!isView && <PrimaryButton text="Save" onClick={this.saveItem} style={{ marginTop:12 }} />}
        </Panel>

        <Dialog hidden={!showConfirmDialog} onDismiss={() => this.setState({ showConfirmDialog:false, deleteItemId:null })} dialogContentProps={{ type: DialogType.normal, title: 'Confirm Deletion', subText: 'Are you sure you want to delete this employee record?' }}>
          <DialogFooter>
            <PrimaryButton text="Delete" onClick={async () => {
              if (deleteItemId !== null) {
                await sp.web.lists.getByTitle('Employee').items.getById(deleteItemId).delete();
                this.setState({ showConfirmDialog:false, deleteItemId:null });
                this.loadAllItems();
              }
            }} />
            <DefaultButton text="Cancel" onClick={() => this.setState({ showConfirmDialog:false, deleteItemId:null })} />
          </DialogFooter>
        </Dialog>

<Dialog
  hidden={!this.state.showDetailDialog}
  onDismiss={() => this.setState({ showDetailDialog: false })}
  dialogContentProps={{
    type: DialogType.normal,
    title: 'Details from JSONValue'
  }}
  modalProps={{
    isBlocking: true,
    containerClassName: styles.customDialogContainer
  }}
>
  <div style={{ padding: '10px 20px', wordBreak: 'break-word' }}>
    {this.state.currentItem.jsonValue ? (() => {
      try {
        const parsed = JSON.parse(this.state.currentItem.jsonValue);
        if (Array.isArray(parsed) && parsed.length > 0) {
          const keys = Object.keys(parsed[0]);
          return (
            <table style={{ width: '100%', borderCollapse: 'collapse', marginTop: '10px' }}>
              <thead>
                <tr>
                  {keys.map((key) => (
                    <th
                      key={key}
                      style={{
                        border: '1px solid #ccc',
                        padding: '8px',
                        backgroundColor: '#f2f2f2',
                        textAlign: 'left'
                      }}
                    >
                      {key}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {parsed.map((item: any, index: number) => (
                  <tr key={index}>
                    {keys.map((key) => (
                      <td
                        key={key}
                        style={{ border: '1px solid #ccc', padding: '8px' }}
                      >
                        {item[key]}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          );
        } else {
          return <p>No data available</p>;
        }
      } catch (e) {
        return <p>Invalid JSON format</p>;
      }
    })() : <p>No data available</p>}
  </div>

  <DialogFooter>
    <PrimaryButton text="Close" onClick={() => this.setState({ showDetailDialog: false })} />
  </DialogFooter>
</Dialog>


      </div>
    );
  }
}
