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
  selectedRole: string;
  onEditClick: () => void;
  onTestClick: () => void;
  onLogout: () => void;
}

const ensureUserPersona = async (email: string): Promise<IPersonaProps | null> => {
  try {
    const result = await sp.web.ensureUser(email);
    const user = result.data;
    return {
      text: user.Title,
      secondaryText: user.Email,
      id: user.Id.toString(),
      imageInitials: user.Title?.[0] || '?'
    };
  } catch {
    return null;
  }
};

const TicketList: React.FC<ITicketListProps> = ({ welcomeName, selectedRole, onEditClick, onTestClick, onLogout }) => {
  const [tickets, setTickets] = useState<any[]>([]);
  const [submittedTickets, setSubmittedTickets] = useState<any[]>([]);
  const [showSubmitted, setShowSubmitted] = useState(true);
  const [selectedTicket, setSelectedTicket] = useState<any | null>(null);
  const [isDialogOpen, setIsDialogOpen] = useState(false);
  const [isViewDialogOpen, setIsViewDialogOpen] = useState(false);
  const [viewItem, setViewItem] = useState<any | null>(null);
  const [ticketTitle, setTicketTitle] = useState('');
  const [ticketDescription, setTicketDescription] = useState('');
  const [selectedSkillset, setSelectedSkillset] = useState<number[]>([]);
  const [assignedOn, setAssignedOn] = useState<Date | undefined>(undefined);
  const [requestor, setRequestor] = useState<IPersonaProps[]>([]);
  const [assignedTo, setAssignedTo] = useState<IPersonaProps[]>([]);
  const [skillsetOptions, setSkillsetOptions] = useState<IDropdownOption[]>([]);

  useEffect(() => {
    fetchTickets();
    fetchSkillsets();
  }, []);

  const resolvePeopleFromTenant = async (filter: string): Promise<IPersonaProps[]> => {
    try {
      const results = await sp.utility.searchPrincipals(filter, 1, 15, '', 5);
      return results
        .filter((p: any) => p.DisplayName && p.Email)
        .map((p: any) => ({
          text: p.DisplayName,
          secondaryText: p.Email,
          id: p.EntityData?.SPUserID?.toString() || ''
        }));
    } catch (error) {
      console.error('❌ resolvePeopleFromTenant error:', error);
      return [];
    }
  };

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

      const enriched = items.map(item => ({
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
        RequestorEmail: item.Requestor?.EMail || '',
        AssignedTo: item.AssignedTo?.Title || '',
        AssignedToEmail: item.AssignedTo?.EMail || '',
        Manager: item.Manager?.Title || '',
      }));

      setTickets(enriched);
      setSubmittedTickets(enriched.filter(t => t.Status === 'Submitted'));
    } catch (error) {
      console.error('❌ fetchTickets error:', error);
    }
  };

  const fetchSkillsets = async () => {
    const skills = await sp.web.lists.getByTitle('Skillset_Master').items.select('Id', 'Title')();
    setSkillsetOptions(skills.map(s => ({ key: s.Id, text: s.Title })));
  };

const handleSave = async () => {
  try {
    // 🔍 Log values from People Pickers
    console.log("🎯 Requestor displayName:", requestor[0]?.text);
    console.log("🎯 Requestor email:", requestor[0]?.secondaryText);
    console.log("🎯 AssignedTo displayName:", assignedTo[0]?.text);
    console.log("🎯 AssignedTo email:", assignedTo[0]?.secondaryText);

    // ✅ STEP 1: Resolve People Picker values via ensureUser
    const requestorEmail = requestor[0]?.secondaryText || '';
    const assignedToEmail = assignedTo[0]?.secondaryText || '';

    let requestorId = null;
    let assignedToId = null;
    let managerId = null;

    if (requestorEmail) {
      const ensuredRequestor = await sp.web.ensureUser(requestorEmail);
      requestorId = ensuredRequestor.data.Id;
    }

    if (assignedToEmail) {
      const ensuredAssignedTo = await sp.web.ensureUser(assignedToEmail);
      assignedToId = ensuredAssignedTo.data.Id;
    }

const requestorDisplayName = requestor[0]?.text || '';


if (requestorDisplayName) {
  const managerItems = await sp.web.lists.getByTitle("Manager_Map")
    .items
    .filter(`User_Name/Title eq '${requestorDisplayName}'`)
    .select("Manager_Name/Title", "User_Name/Title")
    .expand("User_Name", "Manager_Name")
    .top(1)
    .get();

  console.log("📦 Manager_Map items returned:", managerItems);

  if (managerItems.length > 0) {
    const ensuredManager = await sp.web.ensureUser(managerItems[0].Manager_Name?.Title);
    managerId = ensuredManager.data.Id;
  }
}

    // ✅ STEP 3: Build the ticket insert object
    const data: any = {
      Title: ticketTitle,
      Description: ticketDescription,
      Status: selectedTicket ? selectedTicket.Status : 'Submitted',
      AssignedOn: assignedOn ? assignedOn.toISOString() : null,
      SkillsetId: { results: selectedSkillset }
    };

    if (managerId) {
  data.ManagerId = managerId;
}
    if (requestorId) data.RequestorId = requestorId;
    if (assignedToId) data.AssignedToId = assignedToId;
    if (managerId) data.ManagerId = managerId;

    console.log("🧠 requestorId:", requestorId);
console.log("🧠 assignedToId:", assignedToId);
console.log("🧠 managerId:", managerId);

    console.log("📝 Inserting ticket with data:", data);

    // ✅ STEP 4: Insert or update
    if (selectedTicket) {
      await sp.web.lists.getByTitle('Tickets').items.getById(selectedTicket.Id).update(data);
    } else {
      await sp.web.lists.getByTitle('Tickets').items.add(data);
    }

    console.log("✅ Ticket saved successfully");

    closeDialog();
    fetchTickets();
  } catch (error) {
    console.error('❌ handleSave error:', error);
  }
};



  const handleStatusChange = async (id: number, status: string) => {
    await sp.web.lists.getByTitle('Tickets').items.getById(id).update({ Status: status });
    fetchTickets();
  };

  const openEditDialog = async (ticket: any) => {
    try {
      setSelectedTicket(ticket);
      setTicketTitle(ticket.Title);
      setTicketDescription(ticket.Description);
      setSelectedSkillset(ticket.SkillsetId || []);
      setAssignedOn(ticket.AssignedOn ? new Date(ticket.AssignedOn) : undefined);
      const requestorPersona = ticket.RequestorEmail ? await ensureUserPersona(ticket.RequestorEmail) : null;
      const assignedToPersona = ticket.AssignedToEmail ? await ensureUserPersona(ticket.AssignedToEmail) : null;
      setRequestor(requestorPersona ? [requestorPersona] : []);
      setAssignedTo(assignedToPersona ? [assignedToPersona] : []);
      setIsDialogOpen(true);
    } catch (err) {
      console.error('❌ openEditDialog error:', err);
    }
  };

  const openViewDialog = (item: any) => {
    setViewItem(item);
    setIsViewDialogOpen(true);
  };

  const closeViewDialog = () => {
    setIsViewDialogOpen(false);
    setViewItem(null);
  };

  const resetForm = () => {
    setTicketTitle('');
    setTicketDescription('');
    setSelectedSkillset([]);
    setAssignedOn(undefined);
    setRequestor([]);
    setAssignedTo([]);
    setSelectedTicket(null);
  };

  const closeDialog = () => {
    setIsDialogOpen(false);
    resetForm();
  };

  const handleDelete = async (id: number) => {
    if (confirm('Are you sure?')) {
      await sp.web.lists.getByTitle('Tickets').items.getById(id).recycle();
      fetchTickets();
    }
  };

  const columns: IColumn[] = [
    { key: 'col1', name: 'Title', fieldName: 'Title', minWidth: 100 },
    { key: 'col2', name: 'Description', fieldName: 'Description', minWidth: 150 },
    { key: 'col3', name: 'Requestor', fieldName: 'Requestor', minWidth: 120 },
    { key: 'colAssignedTo', name: 'Assigned To', fieldName: 'AssignedTo', minWidth: 120 },
    { key: 'col4', name: 'Skillset', fieldName: 'Skillset', minWidth: 100 },
    { key: 'colManager', name: 'Manager', fieldName: 'Manager', minWidth: 120 },
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

  return (
    <Stack tokens={{ childrenGap: 15 }}>
      {selectedRole === 'Support_Manager' && (
        <Stack tokens={{ childrenGap: 8 }}>
          <PrimaryButton text={showSubmitted ? 'Hide Submitted Requests' : 'Show Submitted Requests'} onClick={() => setShowSubmitted(!showSubmitted)} />
          {showSubmitted && (
            <DetailsList
              items={submittedTickets}
              columns={[
                { key: 'title', name: 'Title', fieldName: 'Title', minWidth: 100 },
                { key: 'desc', name: 'Description', fieldName: 'Description', minWidth: 150 },
                {
                  key: 'actions', name: 'Actions', minWidth: 200, onRender: item => (
                    <Stack horizontal tokens={{ childrenGap: 8 }}>
                      <PrimaryButton text="Accept" onClick={() => handleStatusChange(item.Id, 'Approved')} />
                      <DefaultButton text="Reject" onClick={() => handleStatusChange(item.Id, 'Rejected')} />
                    </Stack>
                  )
                }
              ]}
              layoutMode={DetailsListLayoutMode.fixedColumns}
            />
          )}
        </Stack>
      )}

      <Stack horizontal horizontalAlign="space-between">
        <Text variant="xLarge">🎟️ Tickets</Text>
        <PrimaryButton text="+ Add Ticket" onClick={() => setIsDialogOpen(true)} />
      </Stack>

      <DetailsList items={tickets} columns={columns} layoutMode={DetailsListLayoutMode.fixedColumns} />

      <Dialog hidden={!isDialogOpen} onDismiss={closeDialog} dialogContentProps={{ type: DialogType.largeHeader, title: selectedTicket ? 'Edit Ticket' : 'Add Ticket' }}>
        <Stack tokens={{ childrenGap: 10 }}>
          <TextField label="Title" value={ticketTitle} onChange={(e, v) => setTicketTitle(v || '')} />
          <TextField label="Description" multiline rows={3} value={ticketDescription} onChange={(e, v) => setTicketDescription(v || '')} />
          <Dropdown label="Skillset" options={skillsetOptions} selectedKeys={selectedSkillset} multiSelect onChange={(e, option) => {
            if (option) {
              const updated = option.selected
                ? [...selectedSkillset, option.key as number]
                : selectedSkillset.filter(id => id !== option.key);
              setSelectedSkillset(updated);
            }
          }} />
          <Label>Requestor</Label>
          <BasePeoplePicker
  selectedItems={requestor}
  onResolveSuggestions={resolvePeopleFromTenant}
  onChange={(items: IPersonaProps[] = []) => setRequestor(items)}
  itemLimit={1}
  onRenderItem={(props) => <span>{props.item.text}</span>}
  onRenderSuggestionsItem={(item) => <span>{item.text}</span>}
/>

          <Label>Assigned To</Label>
          <BasePeoplePicker
  selectedItems={assignedTo}
  onResolveSuggestions={resolvePeopleFromTenant}
  onChange={(items: IPersonaProps[] = []) => setAssignedTo(items)}
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
