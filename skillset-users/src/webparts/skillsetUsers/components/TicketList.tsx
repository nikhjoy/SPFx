import * as React from 'react';
import { useEffect, useState } from 'react';
import { sp } from '@pnp/sp/presets/all';
import {
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  Stack,
  Text,
  IconButton,
  PrimaryButton,
  Dialog,
  DialogType,
  DialogFooter,
  DefaultButton,
  TextField,
  Checkbox,
  Dropdown,
  IDropdownOption
} from '@fluentui/react';

interface ITicketListProps {
  welcomeName: string;
  onEditClick: () => void;
  onTestClick: () => void;
  onLogout: () => void;
}

const TicketList: React.FC<ITicketListProps> = ({ welcomeName, onEditClick, onTestClick, onLogout }) => {
  const [tickets, setTickets] = useState<any[]>([]);
  const [selectedTicket, setSelectedTicket] = useState<any | null>(null);
  const [dialogType, setDialogType] = useState<'view' | 'edit' | 'add' | null>(null);
  const [isDialogOpen, setIsDialogOpen] = useState(false);
  const [formData, setFormData] = useState<any>({ Title: '', Description: '', SkillsSet: '', Status: '', IsManagerApproved: false });

  const statusOptions: IDropdownOption[] = [
    { key: 'Open', text: 'Open' },
    { key: 'In Progress', text: 'In Progress' },
    { key: 'Completed', text: 'Completed' }
  ];

  const fetchTickets = async () => {
    const items = await sp.web.lists.getByTitle("Tickets").items
      .select("Id", "Title", "Description", "SkillsSet", "Requestor/Title", "IsManagerApproved", "AssignedTo/Title", "AssignedOn", "Status", "Skillset/Title")
      .expand("Requestor", "AssignedTo", "Skillset")
      .get();
    setTickets(items);
  };

  useEffect(() => {
    fetchTickets();
  }, []);

  const openDialog = (type: 'view' | 'edit' | 'add', ticket?: any) => {
    setDialogType(type);
    setSelectedTicket(ticket || null);
    setFormData(ticket || { Title: '', Description: '', SkillsSet: '', Status: '', IsManagerApproved: false });
    setIsDialogOpen(true);
  };

  const closeDialog = () => {
    setIsDialogOpen(false);
    setDialogType(null);
    setSelectedTicket(null);
  };

  const deleteTicket = async (id: number) => {
    if (confirm("Are you sure you want to delete this ticket?")) {
      await sp.web.lists.getByTitle("Tickets").items.getById(id).delete();
      fetchTickets();
    }
  };

  const handleInputChange = (ev: any, newValue?: string | boolean): void => {
    const { name } = ev.target;
    setFormData({ ...formData, [name]: newValue });
  };

  const handleDropdownChange = (event: any, option?: IDropdownOption): void => {
    if (option) setFormData({ ...formData, Status: option.key });
  };

  const handleSubmit = async () => {
    if (dialogType === 'add') {
      await sp.web.lists.getByTitle("Tickets").items.add(formData);
    } else if (dialogType === 'edit' && selectedTicket) {
      await sp.web.lists.getByTitle("Tickets").items.getById(selectedTicket.Id).update(formData);
    }
    fetchTickets();
    closeDialog();
  };

  const columns: IColumn[] = [
    { key: 'col1', name: 'Title', fieldName: 'Title', minWidth: 100 },
    { key: 'col2', name: 'Description', fieldName: 'Description', minWidth: 150 },
    { key: 'col3', name: 'Requestor', fieldName: 'Requestor', minWidth: 120, onRender: item => item.Requestor?.Title || '—' },
    { key: 'col4', name: 'Status', fieldName: 'Status', minWidth: 100 },
    { key: 'col5', name: 'Actions', fieldName: 'actions', minWidth: 120, onRender: (item) => (
      <Stack horizontal tokens={{ childrenGap: 8 }}>
        <IconButton iconProps={{ iconName: 'View' }} title="View" onClick={() => openDialog('view', item)} />
        <IconButton iconProps={{ iconName: 'Edit' }} title="Edit" onClick={() => openDialog('edit', item)} />
        <IconButton iconProps={{ iconName: 'Delete' }} title="Delete" onClick={() => deleteTicket(item.Id)} />
      </Stack>
    ) }
  ];

return (
  <>
    <Stack horizontal horizontalAlign="space-between" verticalAlign="center" style={{ marginBottom: 10 }}>
      <Text variant="xLarge">🎟️ Tickets</Text>
      <PrimaryButton text="+ Add Ticket" onClick={() => openDialog('add')} />
    </Stack>

    <DetailsList
      items={tickets}
      columns={columns}
      layoutMode={DetailsListLayoutMode.fixedColumns}
      selectionPreservedOnEmptyClick={true}
    />

    {isDialogOpen && (
      <Dialog
        hidden={!isDialogOpen}
        onDismiss={closeDialog}
        dialogContentProps={{
          type: DialogType.largeHeader,
          title:
            dialogType === 'view' ? 'View Ticket' :
            dialogType === 'edit' ? 'Edit Ticket' :
            'Add Ticket'
        }}>

        <Stack tokens={{ childrenGap: 10 }}>
          <TextField
            label="Title"
            name="Title"
            value={formData.Title}
            onChange={handleInputChange}
            disabled={dialogType === 'view'}
          />
          <TextField
            label="Description"
            name="Description"
            multiline
            value={formData.Description}
            onChange={handleInputChange}
            disabled={dialogType === 'view'}
          />
          <TextField
            label="Skills Set"
            name="SkillsSet"
            value={formData.SkillsSet}
            onChange={handleInputChange}
            disabled={dialogType === 'view'}
          />
          <Dropdown
            label="Status"
            selectedKey={formData.Status}
            options={statusOptions}
            onChange={handleDropdownChange}
            disabled={dialogType === 'view'}
          />
          <Checkbox
            label="Manager Approved"
            name="IsManagerApproved"
            checked={formData.IsManagerApproved}
            onChange={(e, checked) => setFormData({ ...formData, IsManagerApproved: !!checked })}
            disabled={dialogType === 'view'}
          />
        </Stack>

        <DialogFooter>
          <DefaultButton onClick={closeDialog} text="Close" />
          {dialogType !== 'view' && <PrimaryButton onClick={handleSubmit} text="Save" />}
        </DialogFooter>
      </Dialog>
    )}
  </>
);

};

export default TicketList;
