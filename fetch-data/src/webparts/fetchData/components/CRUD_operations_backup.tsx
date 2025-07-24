import * as React from 'react';
import { IFetchDataProps } from './IFetchDataProps';
import {
  TextField,
  Dropdown,
  IDropdownOption,
  DatePicker,
  PrimaryButton,
  DetailsList,
  DetailsListLayoutMode,
  IColumn
} from '@fluentui/react';
import { IPersonaProps } from '@fluentui/react';
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
  managerId?: number;
}

type Mode = 'all' | 'view' | 'update' | 'new' | null;

interface IState {
  inputId: string;
  mode: Mode;
  items: IEmployee[];
  item: IEmployee | null;
  departments: IDropdownOption[];
  users: Array<{ Id: number; Title: string; Email: string }>;
  selectedManager: IPersonaProps[];
  errorMessage: string;
  newItem: Partial<IEmployee>;
  newManagerName: string | null;
  newManagerEmail: string | null;
}

export default class EmployeeCrudWebPart extends React.Component<IFetchDataProps, IState> {
  private _columns: IColumn[] = [
    { key: 'col1', name: 'ID', fieldName: 'id', minWidth: 50 },
    { key: 'col2', name: 'Title', fieldName: 'title', minWidth: 100 },
    { key: 'col3', name: 'Name', fieldName: 'name', minWidth: 100 },
    { key: 'col4', name: 'Age', fieldName: 'age', minWidth: 50 },
    { key: 'col5', name: 'Department', fieldName: 'department', minWidth: 100 },
    { key: 'col6', name: 'DOJ', fieldName: 'doj', minWidth: 100, onRender: item => item.doj?.toLocaleDateString() },
    { key: 'col7', name: 'Email', fieldName: 'email', minWidth: 150 },
    { key: 'col8', name: 'Mobile', fieldName: 'mobileNo', minWidth: 100 },
    { key: 'col9', name: 'Manager', fieldName: 'managerId', minWidth: 100 }
  ];

  constructor(props: IFetchDataProps) {
    super(props);
    this.state = {
      inputId: '',
      mode: null,
      items: [],
      item: null,
      departments: [],
      users: [],
      selectedManager: [],
      errorMessage: '',
      newItem: {},
      newManagerEmail: null,
      newManagerName: null
    };
  }

  public componentDidMount(): void {
    this.loadDepartments();
    this.loadUsers();
  }

  private loadDepartments = async (): Promise<void> => {
    const depts = await sp.web.lists.getByTitle('Department').items.select('ID','Title').get();
    this.setState({ departments: depts.map(d => ({ key: d.ID, text: d.Title })) });
  };

  private loadUsers = async (): Promise<void> => {
    const users = await sp.web.siteUsers.select('Id','Title','Email').get();
    this.setState({ users });
  };

  private viewAll = async () => {
    const data: any[] = await sp.web.lists.getByTitle('Employee').items
      .select('ID','Title','EmpName','Age','Department/ID','Department/Title','DOJ','Email','MobileNo','Manager/Id')
      .expand('Department','Manager')
      .get();
    const items = data.map(r => ({
      id: r.ID,
      title: r.Title,
      name: r.EmpName,
      age: r.Age,
      departmentId: r.Department?.ID,
      department: r.Department?.Title || '',
      doj: r.DOJ ? new Date(r.DOJ) : null,
      email: r.Email,
      mobileNo: r.MobileNo,
      managerId: r.Manager?.Id
    }));
    this.setState({ items, mode: 'all', item: null });
  };

  private viewById = async () => {
    const id = parseInt(this.state.inputId);
    if (isNaN(id)) return this.setState({ errorMessage: 'Enter a valid ID.' });
    const r: any = await sp.web.lists.getByTitle('Employee').items.getById(id)
      .select('ID','Title','EmpName','Age','Department/ID','Department/Title','DOJ','Email','MobileNo','Manager/Id','Manager/Title')
      .expand('Department','Manager')
      .get();
    const item: IEmployee = {
      id: r.ID,
      title: r.Title,
      name: r.EmpName,
      age: r.Age,
      departmentId: r.Department?.ID,
      department: r.Department?.Title || '',
      doj: r.DOJ ? new Date(r.DOJ) : null,
      email: r.Email,
      mobileNo: r.MobileNo,
      managerId: r.Manager?.Id
    };
    const mgr = this.state.users.find(u => u.Id === item.managerId);
    this.setState({
      item,
      selectedManager: mgr ? [{ key: mgr.Id.toString(), text: mgr.Title, secondaryText: mgr.Email }] : [],
      mode: 'view'
    });
  };

  private deleteById = async () => {
    const id = parseInt(this.state.inputId);
    if (isNaN(id)) return this.setState({ errorMessage: 'Enter a valid ID.' });
    await sp.web.lists.getByTitle('Employee').items.getById(id).delete();
    alert(`Deleted item ${id}`);
    this.setState({ mode: null });
  };

  private fetchForUpdate = async () => {
    await this.viewById();
    this.setState({ mode: 'update' });
  };

  private updateField = (field: keyof IEmployee, value: any) => {
    if (!this.state.item) return;
    this.setState({ item: {...this.state.item, [field]: value} });
  };

  private saveUpdate = async () => {
    const itm = this.state.item!;
    await sp.web.lists.getByTitle('Employee').items.getById(itm.id).update({
      Title: itm.title,
      EmpName: itm.name,
      Age: itm.age,
      DepartmentId: itm.departmentId,
      DOJ: itm.doj?.toISOString() ?? null,
      Email: itm.email,
      MobileNo: itm.mobileNo,
      ManagerId: itm.managerId
    });
    alert('Item updated successfully');
    this.setState({ mode: null });
  };

  private insertNewItem = async () => {
    const { newItem, newManagerEmail } = this.state;
    if (!newItem.departmentId || !newItem.doj || !newManagerEmail) {
      alert("Please fill all required fields.");
      return;
    }
    try {
      const ensured = await sp.web.ensureUser(newManagerEmail);
      const added = await sp.web.lists.getByTitle('Employee').items.add({
        Title: newItem.title,
        EmpName: newItem.name,
        Age: newItem.age,
        DepartmentId: newItem.departmentId,
        DOJ: newItem.doj.toISOString(),
        Email: newItem.email,
        MobileNo: newItem.mobileNo,
        ManagerId: ensured.data.Id
      });
      alert(`Item inserted successfully (ID: ${added.data.ID})`);
      this.setState({ mode: null });
    } catch (err) {
      console.error("Insert failed", err);
    }
  };

  public render(): React.ReactElement {
    const { inputId, mode, items, item, departments, selectedManager, errorMessage, newItem } = this.state;
    return (
      <div>
        <TextField label="ID" value={inputId} onChange={(e, val) => this.setState({ inputId: val || '' })} />
        <div style={{ margin: '10px 0' }}>
          <PrimaryButton text="View All" onClick={this.viewAll} style={{ marginRight: 8 }} />
          <PrimaryButton text="View" onClick={this.viewById} style={{ marginRight: 8 }} />
          <PrimaryButton text="Delete" onClick={this.deleteById} style={{ marginRight: 8 }} />
          <PrimaryButton text="Update" onClick={this.fetchForUpdate} style={{ marginRight: 8 }} />
          <PrimaryButton text="New" onClick={() => this.setState({ mode: 'new', newItem: {} })} />
        </div>
        {errorMessage && <div style={{ color: 'red' }}>{errorMessage}</div>}
        {mode === 'all' && (
          <DetailsList items={items} columns={this._columns} layoutMode={DetailsListLayoutMode.justified} />
        )}
        {mode === 'view' || mode === 'update' ? item && (
          <div>
            <TextField label="Title" value={item.title} disabled={mode === 'view'} onChange={(e, v) => this.updateField('title', v)} />
            <TextField label="Name" value={item.name} disabled={mode === 'view'} onChange={(e, v) => this.updateField('name', v)} />
            <TextField label="Age" value={item.age.toString()} disabled={mode === 'view'} onChange={(e, v) => this.updateField('age', parseInt(v||'0'))} />
            <Dropdown label="Department" options={departments} selectedKey={item.departmentId} disabled={mode === 'view'} onChange={(e, opt) => this.updateField('departmentId', opt?.key)} />
            <DatePicker label="Date of Joining" value={item.doj || undefined} disabled={mode === 'view'} onSelectDate={date => this.updateField('doj', date)} />
            <TextField label="Email" value={item.email} disabled={mode === 'view'} onChange={(e, v) => this.updateField('email', v)} />
            <TextField label="Mobile No" value={item.mobileNo} disabled={mode === 'view'} onChange={(e, v) => this.updateField('mobileNo', v)} />
            <label style={{ fontWeight: 600, display: 'block', margin: '10px 0 4px' }}>Manager</label>
            <div style={{ marginTop: 10, marginBottom: 10 }}>

<PeoplePicker
  context={{
    ...this.props.context,
    absoluteUrl: this.props.context.pageContext.web.absoluteUrl,
    msGraphClientFactory: this.props.context.msGraphClientFactory,
    spHttpClient: this.props.context.spHttpClient
  }}
  personSelectionLimit={1}
  showtooltip={true}
  required={false}
  disabled={mode === 'view'}
  principalTypes={[PrincipalType.User, PrincipalType.SharePointGroup]}
  defaultSelectedUsers={
    selectedManager.length > 0 ? [selectedManager[0].text ?? ''] : []
  }
  onChange={(items) => {
    if (items[0]) {
      const matched = this.state.users.find(
        u => u.Title === items[0].text || u.Email === items[0].secondaryText
      );
      this.updateField('managerId', matched?.Id);
    } else {
      this.updateField('managerId', undefined);
    }
  }}
  resolveDelay={200}
/>

            </div>
            {mode === 'update' && <PrimaryButton text="Save" onClick={this.saveUpdate} style={{ marginTop: 12 }} />}
          </div>
        ) : null}
        {mode === 'new' && (
          <div style={{ marginTop: 20 }}>
            <TextField label="Title" onChange={(e, v) => this.setState({ newItem: { ...newItem, title: v } })} />
            <TextField label="Name" onChange={(e, v) => this.setState({ newItem: { ...newItem, name: v } })} />
            <TextField label="Age" onChange={(e, v) => this.setState({ newItem: { ...newItem, age: parseInt(v || '0') } })} />
            <Dropdown label="Department" options={departments} onChange={(e, o) => this.setState({ newItem: { ...newItem, departmentId: o?.key as number } })} />
            <DatePicker label="Date of Joining" onSelectDate={date => this.setState({ newItem: { ...newItem, doj: date || null } })} />
            <TextField label="Email" onChange={(e, v) => this.setState({ newItem: { ...newItem, email: v } })} />
            <TextField label="Mobile No" onChange={(e, v) => this.setState({ newItem: { ...newItem, mobileNo: v } })} />
            <div style={{ marginTop: 10 }}>
              <label><strong>Manager</strong></label>

<PeoplePicker
  context={{
    ...this.props.context,
    msGraphClientFactory: this.props.context.msGraphClientFactory,
    spHttpClient: this.props.context.spHttpClient,
    absoluteUrl: this.props.context.pageContext.web.absoluteUrl
  }}
  personSelectionLimit={1}
  showtooltip={true}
  required={true}
  disabled={false}
  principalTypes={[PrincipalType.User, PrincipalType.SharePointGroup]}
  defaultSelectedUsers={[]}
  onChange={(items) => {
    if (items[0]) {
      this.setState({
        newManagerName: items[0].text ?? null,
        newManagerEmail: items[0].secondaryText ?? null
      });
    } else {
      this.setState({
        newManagerName: null,
        newManagerEmail: null
      });
    }
  }}
  resolveDelay={200}
/>


            </div>
            <PrimaryButton text="Insert" onClick={this.insertNewItem} style={{ marginTop: 12 }} />
          </div>
        )}
      </div>
    );
  }
}
