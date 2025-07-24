import * as React from 'react';
import type { IInsertDataProps } from './IInsertDataProps';
import {
  TextField,
  PrimaryButton,
  mergeStyles,
  ITextFieldStyles,
  Dropdown,
  IDropdownOption,
  DatePicker
} from '@fluentui/react';
import { sp } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";

const childTextBoxClass = mergeStyles({
  display: 'block',
  marginBottom: '10px'
});

const TextFieldStyles: Partial<ITextFieldStyles> = { root: { maxWidth: '300px' } };

interface IInsertDataState {
  departmentOptions: { key: number; text: string }[];
  selectedDepartment: number | null;
  doj: Date | null;
  managerEmail: string | null;
  managerName: string | null; // 👈 Added field for display name
}

export default class InsertData extends React.Component<IInsertDataProps, IInsertDataState> {
  constructor(props: IInsertDataProps) {
    super(props);
    this.state = {
      departmentOptions: [],
      selectedDepartment: null,
      doj: null,
      managerEmail: null,
      managerName: null
    };
  }

  public componentDidMount(): void {
    this.loadDepartments();
  }

  private loadDepartments = async () => {
    try {
      const items = await sp.web.lists.getByTitle("Department").items.select("ID", "Title").get();
      const options = items.map(item => ({
        key: item.ID,
        text: item.Title
      }));
      this.setState({ departmentOptions: options });
    } catch (error) {
      console.error("Error loading departments:", error);
    }
  };

  private onDepartmentChange = (
    event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void => {
    if (option) {
      this.setState({ selectedDepartment: option.key as number });
    }
  };

  private onDateChange = (date: Date | null | undefined): void => {
    this.setState({ doj: date ?? null });
  };

  private onManagerChange = (items: any[]) => {
    if (items.length > 0) {
      this.setState({
        managerEmail: items[0].secondaryText, // Email
        managerName: items[0].text            // 👈 Display name
      });
    } else {
      this.setState({
        managerEmail: null,
        managerName: null
      });
    }
  };

  public render(): React.ReactElement<IInsertDataProps> {
    return (
      <div>
        <TextField
          className={childTextBoxClass}
          styles={TextFieldStyles}
          id="title"
          label='Title:' />

        <TextField
          className={childTextBoxClass}
          styles={TextFieldStyles}
          id="empname"
          label='EmpName:' />

        <Dropdown
          placeholder="Select a department"
          label="Department"
          options={this.state.departmentOptions}
          selectedKey={this.state.selectedDepartment ?? undefined}
          onChange={this.onDepartmentChange}
          styles={{ dropdown: { width: 300, marginBottom: 10 } }}
        />

        <TextField
          className={childTextBoxClass}
          styles={TextFieldStyles}
          id="age"
          label='Age:' />

        <DatePicker
          label="Date of Joining"
          value={this.state.doj ?? undefined}
          onSelectDate={this.onDateChange}
          placeholder="Select a date..."
          ariaLabel="Select a date"
          styles={{ root: { marginBottom: 10, maxWidth: 300 } }}
        />

        <div style={{ marginBottom: 10, maxWidth: 300 }}>
          <label>Manager Name</label>
          <PeoplePicker
            context={this.props.context as any}
            personSelectionLimit={1}
            showtooltip={true}
            required={true}
            disabled={false}
            principalTypes={[PrincipalType.User]}
            onChange={this.onManagerChange}
            resolveDelay={200}
          />
          {this.state.managerName && (
            <div style={{ marginTop: 5, fontStyle: 'italic', color: '#333' }}>
              Selected Manager: {this.state.managerName}
            </div>
          )}
        </div>

        <PrimaryButton text='Insert' onClick={this.createItem} />
      </div>
    );
  }

  private createItem = async () => {
    const title = (document.getElementById("title") as HTMLInputElement).value;
    const age = (document.getElementById("age") as HTMLInputElement).value;
    const empname = (document.getElementById("empname") as HTMLInputElement).value;
    const departmentId = this.state.selectedDepartment;
    const doj = this.state.doj;
    const managerEmail = this.state.managerEmail;

    if (!departmentId) {
      alert("Please select a department.");
      return;
    }

    if (!doj) {
      alert("Please select Date of Joining.");
      return;
    }

    if (!managerEmail) {
      alert("Please select a manager.");
      return;
    }

    try {
      const ensureUser = await sp.web.ensureUser(managerEmail);

      const addItem = await sp.web.lists.getByTitle("Employee").items.add({
        'Title': title,
        'EmpName': empname,
        'Age': age,
        'DepartmentId': departmentId,
        'DOJ': doj.toISOString(),
        'Manager': ensureUser.data.Id
      });

      console.log(addItem);
      alert(`Item created successfully with ID: ${addItem.data.ID}`);
    } catch (error) {
      console.error("Error creating item:", error);
      alert("An error occurred while creating the item.");
    }
  }
}
