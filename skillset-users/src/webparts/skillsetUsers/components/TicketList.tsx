import * as React from 'react';
import { useEffect, useState } from 'react';
import { sp } from '@pnp/sp/presets/all';
import {
  DetailsList, DetailsListLayoutMode, IColumn,
  Stack, Text, IconButton, PrimaryButton,
  Dialog, DialogType, DialogFooter, TextField,
  Dropdown, Label, DatePicker, DefaultButton, IDropdownOption
} from '@fluentui/react';
import { IGroup } from '@fluentui/react';
import { PeoplePicker, IPeoplePickerUserItem } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import { SPHttpClient, MSGraphClientFactory } from '@microsoft/sp-http';

interface ITicketListProps {
  welcomeName: string;
  selectedRole: string;
  loginEmail: string;
  context: {
    spHttpClient: SPHttpClient;
    msGraphClientFactory: MSGraphClientFactory;
    absoluteUrl: string;
  };
  onEditClick: () => void;
  onTestClick: () => void;
  onLogout: () => void;
}

const TicketList: React.FC<ITicketListProps> = ({ welcomeName, selectedRole, loginEmail, context, onEditClick, onTestClick, onLogout }) => {
  const [tickets, setTickets] = useState<any[]>([]);
  const [selectedTicket, setSelectedTicket] = useState<any | null>(null);
  const [isDialogOpen, setIsDialogOpen] = useState(false);
  const [isViewDialogOpen, setIsViewDialogOpen] = useState(false);
  const [viewItem, setViewItem] = useState<any | null>(null);
  const [ticketTitle, setTicketTitle] = useState('');
  const [ticketDescription, setTicketDescription] = useState('');
  const [selectedSkillset, setSelectedSkillset] = useState<number[]>([]);
  const [assignedOn, setAssignedOn] = useState<Date | undefined>(undefined);
  const [requestor, setRequestor] = useState<string[]>([]);
  const [assignedTo, setAssignedTo] = useState<string[]>([]);
  const [skillsetOptions, setSkillsetOptions] = useState<IDropdownOption[]>([]);
  const [deleteConfirmOpen, setDeleteConfirmOpen] = useState(false);
  const [ticketToDelete, setTicketToDelete] = useState<any | null>(null);


  const closeDialog = () => {
    setIsDialogOpen(false);
    setTicketTitle('');
    setTicketDescription('');
    setSelectedSkillset([]);
    setAssignedOn(undefined);
    setRequestor([]);
    setAssignedTo([]);
    setSelectedTicket(null);
  };


  useEffect(() => {
  fetchTickets();
  const fetchSkillsets = async () => {
    const skills = await sp.web.lists.getByTitle('Skillset_Master').items.select('Id', 'Title')();
    setSkillsetOptions(skills.map(s => ({ key: s.Id, text: s.Title })));
  };
  fetchSkillsets();
}, []);


  const fetchTickets = async () => {
  try {
    const items = await sp.web.lists.getByTitle('Tickets').items.select(
      'Id', 'Title', 'Description', 'Status', 'AssignedOn', 'SkillsetId',
      'Requestor/Title', 'Requestor/EMail',
      'AssignedTo/Title', 'AssignedTo/EMail',
      'Manager/Title', 'Manager/EMail'
    ).expand('Requestor', 'AssignedTo', 'Manager').get();

    const skills = await sp.web.lists.getByTitle('Skillset_Master').items.select('Id', 'Title')();
    const skillMap = new Map(skills.map(s => [s.Id, s.Title]));

const enriched = items.map(item => {
  console.log('📌 Requestor:', item.Requestor);
  console.log('📌 AssignedTo:', item.AssignedTo);

  return {
    Id: item.Id,
    Title: item.Title,
    Description: item.Description,
    Status: item.Status,
    AssignedOn: item.AssignedOn,
    SkillsetId: item.SkillsetId || [],
    Skillset: Array.isArray(item.SkillsetId)
      ? item.SkillsetId.map((id: number) => skillMap.get(id)).filter(Boolean).join(', ')
      : skillMap.get(item.SkillsetId) || 'Not Assigned',
    Requestor: item.Requestor?.Title || '',
    RequestorEmail: item.Requestor?.EMail || item.Requestor?.UserPrincipalName || '',
    AssignedTo: item.AssignedTo?.Title || '',
    AssignedToEmail: item.AssignedTo?.EMail || item.AssignedTo?.UserPrincipalName || '',
    Manager: item.Manager?.Title || ''
  };
});

console.log('🔍 Raw SharePoint items:', items);


    setTickets(enriched);
  } catch (error) {
    console.error('❌ fetchTickets error:', error);
  }
};


const handleSave = async () => {
  try {
    const requestorEmail = requestor[0];
    const assignedToEmail = assignedTo[0];
    let requestorId = null;
    let assignedToId = null;
    let managerId = null;

    if (requestorEmail) {
      const ensuredRequestor = await sp.web.ensureUser(requestorEmail);
      requestorId = ensuredRequestor.data.Id;

const managerItems = await sp.web.lists.getByTitle('Manager_Map')
  .items
  .select('ID', 'User_Name/Id', 'User_Name/EMail', 'Manager_Name/Id', 'Manager_Name/EMail')
  .expand('User_Name', 'Manager_Name')
  .filter(`User_Name/EMail eq '${requestorEmail}'`)
  .top(1)
  .get();

  console.log("📁 Raw Manager_Map items:", managerItems);

const managerEmail = managerItems?.[0]?.Manager_Name?.EMail;
if (managerEmail) {
  const ensuredManager = await sp.web.ensureUser(managerEmail);
  managerId = ensuredManager.data.Id;
  console.log("👨‍💼 Manager ID to set:", managerId);
  console.log("🧾 Will assign to field 'ManagerId'");
}
    }

    if (assignedToEmail) {
      const ensuredAssignedTo = await sp.web.ensureUser(assignedToEmail);
      assignedToId = ensuredAssignedTo.data.Id;
    }

    const data: any = {
      Title: ticketTitle,
      Description: ticketDescription,
      AssignedOn: assignedOn ? assignedOn.toISOString() : null,
      SkillsetId: { results: selectedSkillset },
      Status: selectedTicket ? selectedTicket.Status : 'Submitted'
    };

    if (requestorId) data.RequestorId = requestorId;
    if (assignedToId) data.AssignedToId = assignedToId;

if (managerId !== null && managerId !== undefined) {
  console.log("👨‍💼 Manager ID to set:", managerId);
  console.log("🧾 Will assign to field 'ManagerId'");
  data['ManagerId'] = managerId;
}

    if (selectedTicket) {
      await sp.web.lists.getByTitle('Tickets').items.getById(selectedTicket.Id).update(data);
    } else {
      console.log("🧾 Payload being sent to SharePoint:", data);
      await sp.web.lists.getByTitle('Tickets').items.add(data);
    }

    closeDialog();
    fetchTickets();
  } catch (error) {
    console.error('❌ handleSave error:', error);
  }
};



    const openEditDialog = async (ticket: any) => {
    setSelectedTicket(ticket);
    setTicketTitle(ticket.Title);
    setTicketDescription(ticket.Description);
    setSelectedSkillset(ticket.SkillsetId || []);
    setAssignedOn(ticket.AssignedOn ? new Date(ticket.AssignedOn) : undefined);
    setRequestor([ticket.RequestorEmail]);
    setAssignedTo([ticket.AssignedToEmail]);
    setIsDialogOpen(true);
  };

  const openViewDialog = (item: any) => {
  setViewItem(item);
  setIsViewDialogOpen(true);
};

const confirmDelete = async () => {
  if (ticketToDelete) {
    await sp.web.lists.getByTitle('Tickets').items.getById(ticketToDelete.Id).recycle();
    setDeleteConfirmOpen(false);
    setTicketToDelete(null);
    await fetchTickets();
  }
};

  const filteredTickets = React.useMemo(() => {
    if (selectedRole === 'Support_Manager') return tickets.filter(t => t.Status === 'Submitted');
    if (selectedRole === 'Support_Seeker') return tickets.filter(t => t.RequestorEmail === loginEmail);
    return tickets;
  }, [tickets, selectedRole, loginEmail]);

  const managerGroups = React.useMemo<IGroup[] | undefined>(() => {
    if (selectedRole !== 'Support_Manager') return undefined;
    return [
      {
        key: 'group1',
        name: 'Submitted Requests',
        startIndex: 0,
        count: filteredTickets.length,
        level: 0,
        isCollapsed: false
      }
    ];
  }, [selectedRole, filteredTickets]);

  const columns: IColumn[] = [
    { key: 'col1', name: 'Title', fieldName: 'Title', minWidth: 100 },
    { key: 'col2', name: 'Description', fieldName: 'Description', minWidth: 150 },
    { key: 'col3', name: 'Requestor', fieldName: 'Requestor', minWidth: 120 },
    { key: 'col4', name: 'Assigned To', fieldName: 'AssignedTo', minWidth: 120 },
    { key: 'col5', name: 'Skillset', fieldName: 'Skillset', minWidth: 100 },
    { key: 'col6', name: 'Manager', fieldName: 'Manager', minWidth: 120 },
    { key: 'col7', name: 'Assigned On', fieldName: 'AssignedOn', minWidth: 120 },
    { key: 'col8', name: 'Status', fieldName: 'Status', minWidth: 100 },
    {
      key: 'col9',
      name: 'Actions',
      minWidth: 160,
      onRender: item => (
        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <IconButton iconProps={{ iconName: 'View' }} onClick={() => openViewDialog(item)} />
          <IconButton iconProps={{ iconName: 'Edit' }} onClick={() => openEditDialog(item)} />
          <IconButton iconProps={{ iconName: 'Delete' }} onClick={() => {
            setTicketToDelete(item);
            setDeleteConfirmOpen(true);
          }}
          />
          <Dialog
  hidden={!deleteConfirmOpen}
  onDismiss={() => setDeleteConfirmOpen(false)}
  dialogContentProps={{
    type: DialogType.normal,
    title: 'Confirm Delete',
    subText: `Are you sure you want to delete the ticket "${ticketToDelete?.Title}"?`,
  }}
>
  <DialogFooter>
    <PrimaryButton text="Yes, Delete" onClick={confirmDelete} />
    <DefaultButton text="Cancel" onClick={() => setDeleteConfirmOpen(false)} />
  </DialogFooter>
</Dialog>

        </Stack>
      )
    }
  ];

  //3

  return (
    <Stack tokens={{ childrenGap: 15 }}>
      <Stack horizontal horizontalAlign="space-between">
        <Text variant="xLarge">🎟️ Tickets</Text>
        <PrimaryButton text="+ Add Ticket" onClick={() => setIsDialogOpen(true)} />
      </Stack>

      <DetailsList
        items={filteredTickets}
        columns={columns}
        groups={managerGroups}
        layoutMode={DetailsListLayoutMode.fixedColumns}
      />

      <Dialog hidden={!isDialogOpen} onDismiss={closeDialog} dialogContentProps={{
        type: DialogType.largeHeader,
        title: selectedTicket ? 'Edit Ticket' : 'Add Ticket'
      }}>
        <Stack tokens={{ childrenGap: 10 }}>
          <TextField label="Title" value={ticketTitle} onChange={(e, v) => setTicketTitle(v || '')} required />
          <TextField label="Description" multiline rows={3} value={ticketDescription} onChange={(e, v) => setTicketDescription(v || '')} />
          <Dropdown
            label="Skillset"
            options={skillsetOptions}
            selectedKeys={selectedSkillset}
            multiSelect
            onChange={(e, option) => {
              if (option) {
                const updated = option.selected
                  ? [...selectedSkillset, option.key as number]
                  : selectedSkillset.filter(id => id !== option.key);
                setSelectedSkillset(updated);
              }
            }}
          />
<Label>Requestor</Label>
<PeoplePicker
  context={context}
  defaultSelectedUsers={requestor}
  onChange={(items: IPeoplePickerUserItem[]) =>
    setRequestor(items.map(i => i.secondaryText || ''))
  }
  personSelectionLimit={1}
  showHiddenInUI={false}
  principalTypes={[1]}
  resolveDelay={250}
/>

<Label>Assigned To</Label>
<PeoplePicker
  context={context}
  defaultSelectedUsers={assignedTo}
  onChange={(items: IPeoplePickerUserItem[]) =>
    setAssignedTo(items.map(i => i.secondaryText || ''))
  }
  personSelectionLimit={1}
  showHiddenInUI={false}
  principalTypes={[1]}
  resolveDelay={250}
/>

          <DatePicker label="Assigned On" value={assignedOn} onSelectDate={(date) => setAssignedOn(date || undefined)} />
        </Stack>
        <DialogFooter>
          <PrimaryButton text="Save" onClick={handleSave} />
          <DefaultButton text="Cancel" onClick={closeDialog} />
        </DialogFooter>
      </Dialog>

      <Dialog hidden={!isViewDialogOpen} onDismiss={() => setIsViewDialogOpen(false)} dialogContentProps={{
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
          <PrimaryButton text="Close" onClick={() => setIsViewDialogOpen(false)} />
        </DialogFooter>
      </Dialog>
    </Stack>
  );
};

export default TicketList;