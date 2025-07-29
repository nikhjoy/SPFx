import * as React from 'react';
import { useEffect, useState } from 'react';
import { sp } from '@pnp/sp/presets/all';
import {
  DetailsList, DetailsListLayoutMode, IColumn,
  Stack, Text, IconButton, PrimaryButton,
  Dialog, DialogType, DialogFooter,
  DefaultButton, TextField, 
  Dropdown, IDropdownOption, Label, DatePicker
} from '@fluentui/react';
import { BasePeoplePicker } from '@fluentui/react/lib/Pickers';
import { IPersonaProps } from '@fluentui/react/lib/Persona';

interface ITicketListProps {
  welcomeName: string;
  onEditClick: () => void;
  onTestClick: () => void;
  onLogout: () => void;
}

const ensureUserPersona = async (email: string): Promise<IPersonaProps | null> => {
  console.log("🔍 ensureUserPersona called with:", email);
  try {
    const result = await sp.web.ensureUser(email);
    const user = result.data;
    const persona = {
      text: user.Title,
      secondaryText: user.Email,
      id: user.Id.toString(),
      imageInitials: user.Title?.[0] || '?'
    };
    console.log("✅ ensureUserPersona returned:", persona);
    return persona;
  } catch (error) {
    console.error("❌ ensureUserPersona error:", error);
    return null;
  }
};


const TicketList: React.FC<ITicketListProps> = ({ welcomeName, onEditClick, onTestClick, onLogout }) => {
  const [tickets, setTickets] = useState<any[]>([]);
  const [selectedTicket, setSelectedTicket] = useState<any | null>(null);
  const [isDialogOpen, setIsDialogOpen] = useState(false);
  const [isViewDialogOpen, setIsViewDialogOpen] = useState(false);
  const [viewItem, setViewItem] = useState<any | null>(null);

  const [ticketTitle, setTicketTitle] = useState('');
  const [ticketDescription, setTicketDescription] = useState('');
  const [selectedSkillset, setSelectedSkillset] = useState<number[]>([]);
  const [selectedStatus, setSelectedStatus] = useState<string | undefined>(undefined);
  const [assignedOn, setAssignedOn] = useState<Date | undefined>(undefined);

  const [requestor, setRequestor] = useState<IPersonaProps[]>([]);
  const [assignedTo, setAssignedTo] = useState<IPersonaProps[]>([]);
  const [skillsetOptions, setSkillsetOptions] = useState<IDropdownOption[]>([]);
  const [statusOptions, setStatusOptions] = useState<IDropdownOption[]>([]);

  useEffect(() => {
    fetchTickets();
    fetchSkillsets();
    fetchStatusChoices();
  }, []);

const resolvePeopleFromTenant = async (filter: string): Promise<IPersonaProps[]> => {
  try {
    const results = await sp.utility.searchPrincipals(filter, 1, 15, "", 5);

    return results
      .filter((person: any) => person.DisplayName && person.Email)
      .map((person: any) => ({
        text: person.DisplayName,
        secondaryText: person.Email,
        id: person.EntityData?.SPUserID?.toString() || ""
      }));
  } catch (error) {
    console.error("Error resolving people:", error);
    return [];
  }
};

const fetchTickets = async () => {
  try {
    // Get all tickets with SkillsetId and related user fields
    const ticketItems = await sp.web.lists.getByTitle('Tickets').items.select(
      'Id',
      'Title',
      'Description',
      'Status',
      'AssignedOn',
      'SkillsetId',
'Requestor/Title',
'Requestor/EMail',
'AssignedTo/Title',
'AssignedTo/EMail'
    ).expand('Requestor', 'AssignedTo').get();

    // Get all skillsets once
    const skillsets = await sp.web.lists.getByTitle('Skillset_Master').items.select('Id', 'Title')();
    const skillsetMap = new Map(skillsets.map(s => [s.Id, s.Title]));

const enrichedItems = ticketItems.map(item => ({
  Id: item.Id,
  Title: item.Title,
  Description: item.Description,
  Status: item.Status,
  AssignedOn: item.AssignedOn,
  SkillsetId: item.SkillsetId || [],
  Skillset: Array.isArray(item.SkillsetId)
    ? item.SkillsetId.map((id: number) => skillsetMap.get(id)).filter(Boolean).join(', ')
    : skillsetMap.get(item.SkillsetId) || 'Not Assigned',
Requestor: item.Requestor?.Title || '',
RequestorEmail: item.Requestor?.EMail || '',
AssignedTo: item.AssignedTo?.Title || '',
AssignedToEmail: item.AssignedTo?.EMail || ''
}));

     console.log("✔ Skillset mapping:", enrichedItems);

    setTickets(enrichedItems);
  } catch (error) {
   
  }
};


  const fetchSkillsets = async () => {
    const items = await sp.web.lists.getByTitle('Skillset_Master').items.select('Id', 'Title')();
    setSkillsetOptions(items.map(item => ({ key: item.Id, text: item.Title })));
  };

  const fetchStatusChoices = async () => {
    const field: any = await sp.web.lists.getByTitle('Tickets').fields.getByInternalNameOrTitle('Status')();
    const choices: string[] = field?.Choices || [];
    setStatusOptions(choices.map(choice => ({ key: choice, text: choice })));
  };

  const openAddDialog = () => {
    resetForm();
    setSelectedTicket(null);
    setIsDialogOpen(true);
  };

const openEditDialog = (ticket: any) => {
  setSelectedTicket(ticket);
  setTicketTitle(ticket.Title);
  setTicketDescription(ticket.Description);
  setSelectedSkillset(ticket.SkillsetId || []);
  setSelectedStatus(ticket.Status);
  setAssignedOn(ticket.AssignedOn ? new Date(ticket.AssignedOn) : undefined);

const loadPeoplePickersAndOpen = async () => {
  const r = ticket.RequestorEmail ? await ensureUserPersona(ticket.RequestorEmail) : null;
  const a = ticket.AssignedToEmail ? await ensureUserPersona(ticket.AssignedToEmail) : null;

  console.log("📌 Setting requestor to:", r);
  console.log("📌 Setting assignedTo to:", a);

  setRequestor(r ? [r] : []);
  setAssignedTo(a ? [a] : []);
  setIsDialogOpen(true);
};

  loadPeoplePickersAndOpen();
};


  const resetForm = () => {
    setTicketTitle('');
    setTicketDescription('');
    setSelectedSkillset([]);
    setSelectedStatus(undefined);
    setAssignedOn(undefined);
    setRequestor([]);
    setAssignedTo([]);
  };


const handleSave = async () => {
  try {


const requestorId = requestor[0]?.id ? parseInt(requestor[0].id) : null;
const assignedToId = assignedTo[0]?.id ? parseInt(assignedTo[0].id) : null;


    const ticketData: any = {
      Title: ticketTitle,
      Description: ticketDescription,
      Status: selectedStatus,
      AssignedOn: assignedOn ? assignedOn.toISOString() : null,
      SkillsetId: { results: selectedSkillset },
      RequestorId: requestorId,
      AssignedToId: assignedToId
    };

    if (selectedTicket) {
      await sp.web.lists.getByTitle('Tickets').items.getById(selectedTicket.Id).update(ticketData);
    } else {
      await sp.web.lists.getByTitle('Tickets').items.add(ticketData);
    }

    closeDialog();
    fetchTickets();
  } catch (error) {
    console.error("❌ Failed to save ticket:", error);
  }
};


  const handleDelete = async (ticketId: number) => {
    if (confirm("Are you sure you want to delete this ticket?")) {
      await sp.web.lists.getByTitle('Tickets').items.getById(ticketId).recycle();
      fetchTickets();
    }
  };

  const closeDialog = () => {
    setIsDialogOpen(false);
    resetForm();
  };

  const openViewDialog = (item: any) => {
    setViewItem(item);
    setIsViewDialogOpen(true);
  };

  const closeViewDialog = () => {
    setIsViewDialogOpen(false);
    setViewItem(null);
  };

const columns: IColumn[] = [
  { key: 'col1', name: 'Title', fieldName: 'Title', minWidth: 100 },
  { key: 'col2', name: 'Description', fieldName: 'Description', minWidth: 150 },
  { key: 'col3', name: 'Requestor', fieldName: 'Requestor', minWidth: 120 },
  { key: 'colAssignedTo', name: 'Assigned To', fieldName: 'AssignedTo', minWidth: 120 },
  { key: 'col4', name: 'Skillset', fieldName: 'Skillset', minWidth: 100 },
  { key: 'col6', name: 'Assigned On', fieldName: 'AssignedOn', minWidth: 120 },
  { key: 'col7', name: 'Status', fieldName: 'Status', minWidth: 100 },
  {
    key: 'col8',
    name: 'Actions',
    minWidth: 120,
    onRender: item => (
      <Stack horizontal tokens={{ childrenGap: 8 }}>
        <IconButton iconProps={{ iconName: 'View' }} title="View" onClick={() => openViewDialog(item)} />
        <IconButton iconProps={{ iconName: 'Edit' }} title="Edit" onClick={() => openEditDialog(item)} />
        <IconButton iconProps={{ iconName: 'Delete' }} title="Delete" onClick={() => handleDelete(item.Id)} />
      </Stack>
    )
  }
];

console.log("🎯 Final requestor state in render:", requestor);
console.log("🎯 Final assignedTo state in render:", assignedTo);

  return (
    <Stack tokens={{ childrenGap: 15 }}>
      <Stack horizontal horizontalAlign="space-between">
        <Text variant="xLarge">🎟️ Tickets</Text>
        <PrimaryButton text="+ Add Ticket" onClick={openAddDialog} />
      </Stack>

      <DetailsList items={tickets} columns={columns} layoutMode={DetailsListLayoutMode.fixedColumns} />

      <Dialog hidden={!isDialogOpen} onDismiss={closeDialog} dialogContentProps={{
        type: DialogType.largeHeader,
        title: selectedTicket ? "Edit Ticket" : "Add Ticket"
      }}>
        <Stack tokens={{ childrenGap: 10 }}>
          <TextField label="Title" value={ticketTitle} onChange={(e, v) => setTicketTitle(v || '')} />
          <TextField label="Description" multiline rows={3} value={ticketDescription} onChange={(e, v) => setTicketDescription(v || '')} />
          <Dropdown
  label="Skillset"
  options={skillsetOptions}
  selectedKeys={selectedSkillset}
  multiSelect
  onChange={(e, option, index) => {
    if (option) {
      const newSelection = option.selected
        ? [...selectedSkillset, option.key as number]
        : selectedSkillset.filter(k => k !== option.key);
      setSelectedSkillset(newSelection);
    }
  }}
/>

          <Dropdown label="Status" options={statusOptions} selectedKey={selectedStatus} onChange={(e, o) => setSelectedStatus(o?.key as string)} />

          <Label>Requestor</Label>
<BasePeoplePicker
  selectedItems={requestor} // fully controlled
  onResolveSuggestions={resolvePeopleFromTenant}
  onChange={(items: IPersonaProps[]) => setRequestor(items)}
  itemLimit={1}
  onRenderItem={(props) => <span>{props.item.text}</span>}
  onRenderSuggestionsItem={(item) => <span>{item.text}</span>}
/>

          <Label>Assigned To</Label>
<BasePeoplePicker
 selectedItems={assignedTo}
 onChange={(items: IPersonaProps[]) => setAssignedTo(items)}
  onResolveSuggestions={resolvePeopleFromTenant}
  itemLimit={1}
  onRenderItem={(props) => <span>{props.item.text}</span>}
  onRenderSuggestionsItem={(item) => <span>{item.text}</span>}
/>


          <DatePicker label="Assigned On" value={assignedOn} onSelectDate={(date) => setAssignedOn(date ?? undefined)} />
        </Stack>

        <DialogFooter>
          <PrimaryButton text="Save" onClick={handleSave} />
          <DefaultButton text="Cancel" onClick={closeDialog} />
        </DialogFooter>
      </Dialog>

      <Dialog hidden={!isViewDialogOpen} onDismiss={closeViewDialog} dialogContentProps={{
        type: DialogType.normal,
        title: 'Ticket Details'
      }}>
        <Stack tokens={{ childrenGap: 8 }}>
          {viewItem && (
            <>
              <Text><strong>Title:</strong> {viewItem.Title}</Text>
              <Text><strong>Description:</strong> {viewItem.Description}</Text>
              <Text><strong>Status:</strong> {viewItem.Status}</Text>
              <Text><strong>Skillset:</strong> {viewItem.Skillset}</Text>
              <Text><strong>Requestor:</strong> {viewItem.Requestor}</Text>
              <Text><strong>Assigned To:</strong> {viewItem.AssignedTo}</Text>
              <Text><strong>Assigned On:</strong> {viewItem.AssignedOn}</Text>
            </>
          )}
        </Stack>
        <DialogFooter>
          <PrimaryButton text="Close" onClick={closeViewDialog} />
        </DialogFooter>
      </Dialog>
    </Stack>
  );
};

export default TicketList;
