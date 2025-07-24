import * as React from 'react';
import { IFetchDataProps } from './IFetchDataProps';
import { TextField, Dropdown, IDropdownOption, DatePicker, PrimaryButton } from '@fluentui/react';
import { BasePeoplePicker, IPersonaProps } from '@fluentui/react';
import { sp } from '@pnp/sp/presets/all';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

export interface IDetailListItem {
  id: number;
  title: string;
  name: string;
  age: number;
  department: string;
  departmentId: number;
  doj: Date | null;
  email: string;
  mobileNo: string;
  managerId?: number;
}

export interface IDetailListState {
  item: IDetailListItem | null;
  inputId: string;
  errorMessage: string;
  departments: IDropdownOption[];
  users: Array<{ Id: number; Title: string; Email: string }>;
  selectedManager: IPersonaProps[];
}

export default class SpfxEditableForm extends React.Component<IFetchDataProps, IDetailListState> {
  constructor(props: IFetchDataProps) {
    super(props);
    this.state = {
      item: null,
      inputId: '',
      errorMessage: '',
      departments: [],
      users: [],
      selectedManager: []
    };
  }

  public componentDidMount(): void {
    this.loadDepartments();
    this.loadTenantUsers();
  }

  private loadTenantUsers = async (): Promise<void> => {
    try {
      const users = await sp.web.siteUsers.select("Id", "Title", "Email").get();
      this.setState({ users });
    } catch (error) {
      console.error("Error loading users:", error);
    }
  };

  private loadDepartments = async (): Promise<void> => {
    const deptItems = await sp.web.lists.getByTitle("Department").items.select("ID", "Title").get();
    const deptOptions = deptItems.map(item => ({
      key: item.ID,
      text: item.Title
    }));
    this.setState({ departments: deptOptions });
  };

  public render(): React.ReactElement<IFetchDataProps> {
    const { item, departments, inputId, errorMessage } = this.state;

    return (
      <div>
        <TextField
          label="Enter ID to fetch"
          value={inputId}
          onChange={(e, newVal) => this.setState({ inputId: newVal || '', errorMessage: '' })}
        />
        <PrimaryButton
          text="Submit"
          onClick={this.getItemById}
          styles={{ root: { marginTop: '10px', marginBottom: '20px' } }}
        />
        {errorMessage && <p style={{ color: 'red' }}>{errorMessage}</p>}

        {item && (
          <div>
            <TextField label="Title" value={item.title} onChange={(e, val) => this.updateField("title", val)} />
            <TextField label="Name" value={item.name} onChange={(e, val) => this.updateField("name", val)} />
            <TextField label="Age" value={item.age.toString()} onChange={(e, val) => this.updateField("age", parseInt(val || "0"))} />
            <Dropdown
              label="Department"
              selectedKey={item.departmentId}
              options={departments}
              onChange={(e, option) => this.updateField("departmentId", option?.key)}
            />
            <DatePicker
              label="Date of Joining"
              value={item.doj ?? undefined}
              onSelectDate={(date) => this.updateField("doj", date)}
            />
            <TextField label="Email" value={item.email} onChange={(e, val) => this.updateField("email", val)} />
            <TextField label="Mobile No" value={item.mobileNo} onChange={(e, val) => this.updateField("mobileNo", val)} />
           
           <label style={{ fontWeight: '600', display: 'block', marginTop: '10px' }}>Manager</label>

<BasePeoplePicker
  key={this.state.inputId}
  selectedItems={this.state.selectedManager}
onChange={(items: IPersonaProps[] | undefined) => {
  console.log("Items in onChange:", items);  // Debugging the selection
  if (items && items.length > 0) {
    const selected = items[0];
    const match = this.state.users.find(
      u => u.Title === selected.text || u.Email === selected.secondaryText
    );
    console.log("Selected Manager:", match);  // Log the match
    this.updateField("managerId", match ? match.Id : undefined);
    this.setState({ selectedManager: items }); // Update the selectedManager state
  } else {
    this.setState({ selectedManager: [] }); // Reset the selectedManager if nothing is selected
  }
}}
  itemLimit={1}
  onRenderItem={(props) => (
    <div>
      {props.item.text} <span style={{ color: '#888' }}>({props.item.secondaryText})</span>
    </div>
  )}
  onRenderSuggestionsItem={(props: IPersonaProps) => <div>{props?.text || 'Unknown'}</div>}
  onResolveSuggestions={(filterText: string) => {
    if (filterText) {
      return this.state.users
        .filter(user => user.Title.toLowerCase().includes(filterText.toLowerCase()))
        .map(user => ({
          key: user.Id.toString(),
          text: user.Title,
          secondaryText: user.Email
        }));
    }
    return [];
  }}
/>

            <PrimaryButton
              text="Update"
              onClick={this.updateItem}
              styles={{ root: { marginTop: '20px' } }}
            />
          </div>
        )}
      </div>
    );
  }

  private updateField = (field: keyof IDetailListItem, value: any) => {
    if (!this.state.item) return;
    const updatedItem = { ...this.state.item, [field]: value };
    this.setState({ item: updatedItem });
  };

private getItemById = async (): Promise<void> => {
  const id = parseInt(this.state.inputId);
  if (isNaN(id)) {
    this.setState({ item: null, errorMessage: 'Please enter a valid numeric ID.' });
    return;
  }

  try {
    const result = await sp.web.lists.getByTitle("Employee").items.getById(id)
      .select("ID", "Title", "EmpName", "Age", "Department/ID", "Department/Title", "DOJ", "Email", "MobileNo", "Manager/Id", "Manager/Title")
      .expand("Department", "Manager")
      .get();

    console.log("Fetched item:", result);

    const matchedItem: IDetailListItem = {
      id: result.ID,
      title: result.Title || '',
      name: result.EmpName || '',
      age: result.Age || 0,
      department: result.Department?.Title || '',
      departmentId: result.Department?.ID || 0,
      doj: result.DOJ ? new Date(result.DOJ) : null,
      email: result.Email || '',
      mobileNo: result.MobileNo || '',
      managerId: result.Manager?.Id || undefined
    };

    // Prepare manager data for picker
    const managerUser = this.state.users.find(u => u.Id === matchedItem.managerId);
    const managerPersona = managerUser
      ? [{
          key: managerUser.Id.toString(),
          text: managerUser.Title,
          secondaryText: managerUser.Email
        }]
      : [];

    // Ensure picker data is being correctly passed
    this.setState({
      item: matchedItem,
      errorMessage: '',
      selectedManager: managerPersona
    });

  } catch (error) {
    console.error("Error fetching item:", error);
    this.setState({ item: null, errorMessage: `No item found with ID ${id}.` });
  }
};


  private updateItem = async (): Promise<void> => {
    const { item } = this.state;
    if (!item) return;

    if (!item.title || !item.name || !item.departmentId) {
    alert("Please fill all required fields.");
    return;
  }

    try {
      await sp.web.lists.getByTitle("Employee").items.getById(item.id).update({
        Title: item.title,
        EmpName: item.name,
        Age: item.age,
        DepartmentId: item.departmentId,
        DOJ: item.doj ? item.doj.toISOString() : null,
        Email: item.email,
        MobileNo: item.mobileNo,
        ManagerId: item.managerId
      });

      alert("Item updated successfully!");
    } catch (err) {
      alert("Update failed. " + err.message);
    }
  };
}
