import * as React from 'react';
import { useEffect, useState } from 'react';
import { sp } from '@pnp/sp/presets/all';
import {
  DetailsList, DetailsListLayoutMode, IColumn,
  Stack, Text, IconButton, PrimaryButton,
  Dialog, DialogType, DialogFooter, TextField,
  Dropdown, Label, DatePicker, DefaultButton, IDropdownOption, Callout, Checkbox
} from '@fluentui/react';
import { IGroup } from '@fluentui/react';
import { PeoplePicker, IPeoplePickerUserItem } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import { SPHttpClient, MSGraphClientFactory } from '@microsoft/sp-http';
import { IDetailsColumnProps } from '@fluentui/react/lib/DetailsList'
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
  const [actionMessage, setActionMessage] = useState<string | null>(null);
  const [requestorFilterOptions, setRequestorFilterOptions] = useState<IDropdownOption[]>([]);

  const [] = useState<string | null>(null);
  const [] = useState<string | null>(null);
  const [isRequestorFilterCalloutVisible, setIsRequestorFilterCalloutVisible] = useState(false);
  const [requestorFilterAnchor, setRequestorFilterAnchor] = useState<HTMLElement | null>(null);
  const [selectedRequestors, setSelectedRequestors] = useState<string[]>([]);
  const [selectedAssignedTo, setSelectedAssignedTo] = useState<string[]>([]);
  const [assignedToFilterOptions, setAssignedToFilterOptions] = useState<IDropdownOption[]>([]);
  const [isAssignedToCalloutVisible, setIsAssignedToCalloutVisible] = useState(false);
  const [assignedToFilterAnchor, setAssignedToFilterAnchor] = useState<HTMLElement | null>(null);

  const [selectedManagers, setSelectedManagers] = useState<string[]>([]);
  const [managerFilterOptions, setManagerFilterOptions] = useState<IDropdownOption[]>([]);
  const [isManagerCalloutVisible, setIsManagerCalloutVisible] = useState(false);
  const [managerFilterAnchor, setManagerFilterAnchor] = useState<HTMLElement | null>(null);

  const [selectedStatuses, setSelectedStatuses] = useState<string[]>([]);
  const [statusFilterOptions, setStatusFilterOptions] = useState<IDropdownOption[]>([]);
  const [isStatusCalloutVisible, setIsStatusCalloutVisible] = useState(false);
  const [statusFilterAnchor, setStatusFilterAnchor] = useState<HTMLElement | null>(null);
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

  const onRenderFilterIcon = (
    anchorSetter: (el: HTMLElement | null) => void,
    visibilitySetter: (v: boolean) => void
  ) => (
    <IconButton
      iconProps={{ iconName: 'Filter' }}
      onClick={e => {
        anchorSetter(e.currentTarget as HTMLElement);
        visibilitySetter(true);
      }}
      title="Filter"
    />
  );


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
      const uniqueRequestors = Array.from(new Set(enriched.map(t => t.Requestor).filter(Boolean)));
      const requestorOptions = uniqueRequestors.map(name => ({ key: name, text: name }));

      setRequestorFilterOptions(requestorOptions);
      setSelectedRequestors(requestorOptions.map(opt => opt.key as string)); // ✅ All selected by default

      const uniqueAssignedTo = Array.from(new Set(enriched.map(t => t.AssignedTo).filter(Boolean)));
      const assignedToOptions = uniqueAssignedTo.map(name => ({ key: name, text: name }));
      setAssignedToFilterOptions(assignedToOptions);
      setSelectedAssignedTo(assignedToOptions.map(opt => opt.key as string));

      const uniqueManagers = Array.from(new Set(enriched.map(t => t.Manager).filter(Boolean)));
      const managerOptions = uniqueManagers.map(name => ({ key: name, text: name }));
      setManagerFilterOptions(managerOptions);
      setSelectedManagers(managerOptions.map(opt => opt.key as string));

      const uniqueStatuses = Array.from(new Set(enriched.map(t => t.Status).filter(Boolean)));
      const statusOptions = uniqueStatuses.map(name => ({ key: name, text: name }));
      setStatusFilterOptions(statusOptions);
      setSelectedStatuses(statusOptions.map(opt => opt.key as string));

    } catch (error) {
      console.error('❌ fetchTickets error:', error);
    }
  };

  //end part 1

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
    let result = tickets;

    if (selectedRole === 'Support_Seeker') {
      result = result.filter(t => t.RequestorEmail === loginEmail && t.Status !== 'Submitted');
    } else if (selectedRole === 'Support_Manager') {
      const statusOrder: Record<string, number> = {
        Submitted: 1,
        Approved: 2,
        Rejected: 3
      };

      result = [...result].sort(
        (a, b) =>
          (statusOrder[a.Status as keyof typeof statusOrder] ?? 99) -
          (statusOrder[b.Status as keyof typeof statusOrder] ?? 99)
      );
    } else if (selectedRole === 'Support_Provider') {
      result = result.filter(t => t.Status !== 'Submitted');
    }

    // New multi-select requestor filter
    if (
      selectedRequestors.length > 0 &&
      !selectedRequestors.includes('all')
    ) {
      result = result.filter(t =>
        selectedRequestors.includes(t.Requestor)
      );
    }

    if (selectedAssignedTo.length > 0) {
      result = result.filter(t => selectedAssignedTo.includes(t.AssignedTo));
    }
    if (selectedManagers.length > 0) {
      result = result.filter(t => selectedManagers.includes(t.Manager));
    }
    if (selectedStatuses.length > 0) {
      result = result.filter(t => selectedStatuses.includes(t.Status));
    }


    return result;
  }, [tickets, selectedRole, loginEmail, selectedRequestors, selectedAssignedTo, selectedManagers, selectedStatuses]);

  const managerGroups = React.useMemo<IGroup[] | undefined>(() => {
    if (selectedRole !== 'Support_Manager') return undefined;

    const grouped: { [key: string]: any[] } = {
      Submitted: [],
      Approved: [],
      Rejected: []
    };

    filteredTickets.forEach(ticket => {
      if (grouped[ticket.Status]) grouped[ticket.Status].push(ticket);
    });

    const groupList: IGroup[] = [];
    let startIndex = 0;

    for (const status of ['Submitted', 'Approved', 'Rejected']) {
      const groupItems = grouped[status];
      if (groupItems.length > 0) {
        groupList.push({
          key: status,
          name: `${status} Requests`,
          startIndex,
          count: groupItems.length,
          level: 0,
          isCollapsed: false
        });
        startIndex += groupItems.length;
      }
    }

    return groupList;
  }, [selectedRole, filteredTickets]);

  const onRenderHeader = (
    props?: IDetailsColumnProps,
    defaultRender?: (props?: IDetailsColumnProps) => JSX.Element
  ): JSX.Element | null => {
    if (!props || !defaultRender) return null;

    const colKey = props.column.key;

    const filterIcon = (() => {
      if (colKey === 'col3') return onRenderFilterIcon(setRequestorFilterAnchor, setIsRequestorFilterCalloutVisible);
      if (colKey === 'col4') return onRenderFilterIcon(setAssignedToFilterAnchor, setIsAssignedToCalloutVisible);
      if (colKey === 'col6') return onRenderFilterIcon(setManagerFilterAnchor, setIsManagerCalloutVisible);
      if (colKey === 'col8') return onRenderFilterIcon(setStatusFilterAnchor, setIsStatusCalloutVisible);
      return null;
    })();

    return (
      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 4 }}>
        <Text>{props.column.name}</Text>
        {filterIcon}
      </Stack>
    );
  };


  const renderManager = (item: any): JSX.Element => {
    return <span>{item.Manager || '—'}</span>;
  };

  const renderStatus = (item: any): JSX.Element => {
    return <span>{item.Status || '—'}</span>;
  };

  const columns: IColumn[] = [
    { key: 'col1', name: 'Title', fieldName: 'Title', minWidth: 100 },
    { key: 'col2', name: 'Description', fieldName: 'Description', minWidth: 150 },
    {
      key: 'col3',
      name: 'Requestor',
      fieldName: 'Requestor',
      minWidth: 120,
      onRenderHeader: onRenderHeader,
    },
    {
      key: 'col4',
      name: 'Assigned To',
      fieldName: 'AssignedTo',
      minWidth: 120,
      onRenderHeader: onRenderHeader
    },

    {
      key: 'col6',
      name: 'Manager',
      fieldName: 'Manager',
      minWidth: 120,
      onRender: renderManager,
      onRenderHeader: onRenderHeader
    },
    {
      key: 'col8',
      name: 'Status',
      fieldName: 'Status',
      minWidth: 100,
      onRender: renderStatus,
      onRenderHeader: onRenderHeader
    },

    {
      key: 'col9',
      name: 'Actions',
      minWidth: 300,
      onRender: (item: any) => (
        <Stack horizontal wrap tokens={{ childrenGap: 6 }}>
          {/* Shared buttons */}
          <IconButton iconProps={{ iconName: 'View' }} onClick={() => openViewDialog(item)} />
          {selectedRole !== 'Support_Manager' && (
            <IconButton iconProps={{ iconName: 'Edit' }} onClick={() => openEditDialog(item)} />
          )}

          <IconButton iconProps={{ iconName: 'Delete' }} onClick={() => {
            setTicketToDelete(item);
            setDeleteConfirmOpen(true);
          }} />

          {/* Approve/Reject visible only for Support_Manager and Submitted status */}
          {selectedRole === 'Support_Manager' && item.Status === 'Submitted' && (
            <>
              {console.log("📌 Approve/Reject rendering for ticket ID", item.Id, " | role =", selectedRole, " | status =", item.Status)}

              <PrimaryButton
                text="Approve"
                onClick={() => handleApproval(item.Id, 'Approved')}
                styles={{
                  root: {
                    backgroundColor: 'green',
                    color: 'white',
                    padding: '0 8px',
                    minWidth: 80
                  }
                }}
              />
              <DefaultButton
                text="Reject"
                onClick={() => handleApproval(item.Id, 'Rejected')}
                styles={{
                  root: {
                    backgroundColor: 'red',
                    color: 'white',
                    padding: '0 8px',
                    minWidth: 80
                  }
                }}
              />
            </>
          )}

          {/* Delete confirmation dialog embedded here */}
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

  //end part 3

  const handleApproval = async (ticketId: number, status: 'Approved' | 'Rejected') => {
    try {
      await sp.web.lists.getByTitle('Tickets').items.getById(ticketId).update({ Status: status });
      setActionMessage(`Ticket ${status} successfully.`);
      await fetchTickets();

      setTimeout(() => setActionMessage(null), 3000); // Auto-clear message
    } catch (error) {
      console.error(`Error updating status to ${status}:`, error);
      setActionMessage(`Failed to update status.`);
      setTimeout(() => setActionMessage(null), 3000);
    }
  };

  console.log("📌 Callout visible:", isRequestorFilterCalloutVisible);
  console.log("📌 Anchor element:", requestorFilterAnchor);


  return (
    <Stack tokens={{ childrenGap: 15 }}>

    <Stack horizontal tokens={{ childrenGap: 10 }}>
      <PrimaryButton
        text="Clear Filters"
        onClick={() => {
          setSelectedRequestors(requestorFilterOptions.map(opt => opt.key as string));
          setSelectedAssignedTo(assignedToFilterOptions.map(opt => opt.key as string));
          setSelectedManagers(managerFilterOptions.map(opt => opt.key as string));
          setSelectedStatuses(statusFilterOptions.map(opt => opt.key as string));
        }}
        styles={{ root: { backgroundColor: '#d83b01', color: 'white', padding: '0 12px', minWidth: 100 } }}
      />
      <PrimaryButton
        text="+ Add Ticket"
        onClick={() => setIsDialogOpen(true)}
      />
    </Stack>

      {/* ✅ Success message appears here */}
      {actionMessage && (
        <Text variant="medium" styles={{ root: { color: 'green', fontWeight: 600 } }}>
          {actionMessage}
        </Text>
      )}

      <DetailsList
        items={filteredTickets}
        columns={columns}
        groups={managerGroups}
        layoutMode={DetailsListLayoutMode.fixedColumns}
      />

      {isRequestorFilterCalloutVisible && requestorFilterAnchor && (
        <Callout
          target={requestorFilterAnchor}
          onDismiss={() => setIsRequestorFilterCalloutVisible(false)}
          setInitialFocus
          directionalHint={8}
          isBeakVisible={true}
          gapSpace={5}
        >
          <Stack tokens={{ childrenGap: 4 }} styles={{ root: { padding: 10 } }}>
            <Text variant="mediumPlus">Filter by Requestor</Text>
            <Checkbox
              label="All"
              checked={selectedRequestors.length === requestorFilterOptions.length}
              onChange={() => {
                const all = requestorFilterOptions.map(opt => opt.key as string);
                const isAllSelected = selectedRequestors.length === all.length;
                setSelectedRequestors(isAllSelected ? [] : all);
              }}
            />
            {requestorFilterOptions.map(opt => (
              <Checkbox
                key={opt.key}
                label={opt.text}
                checked={selectedRequestors.includes(opt.key as string)}
                onChange={() => {
                  const key = opt.key as string;
                  setSelectedRequestors(prev =>
                    prev.includes(key) ? prev.filter(k => k !== key) : [...prev, key]
                  );
                }}
              />
            ))}
          </Stack>
        </Callout>
      )}

      {isAssignedToCalloutVisible && assignedToFilterAnchor && (
        <Callout
          target={assignedToFilterAnchor}
          onDismiss={() => setIsAssignedToCalloutVisible(false)}
          setInitialFocus
          directionalHint={8}
          isBeakVisible={true}
          gapSpace={5}
        >
          <Stack tokens={{ childrenGap: 4 }} styles={{ root: { padding: 10 } }}>
            <Text variant="mediumPlus">Filter by Assigned To</Text>
            <Checkbox
              label="All"
              checked={selectedAssignedTo.length === assignedToFilterOptions.length}
              onChange={() => {
                const all = assignedToFilterOptions.map(opt => opt.key as string);
                const isAllSelected = selectedAssignedTo.length === all.length;
                setSelectedAssignedTo(isAllSelected ? [] : all);
              }}
            />
            {assignedToFilterOptions.map(opt => (
              <Checkbox
                key={opt.key}
                label={opt.text}
                checked={selectedAssignedTo.includes(opt.key as string)}
                onChange={() => {
                  const key = opt.key as string;
                  setSelectedAssignedTo(prev =>
                    prev.includes(key) ? prev.filter(k => k !== key) : [...prev, key]
                  );
                }}
              />
            ))}
          </Stack>
        </Callout>
      )}

      {isManagerCalloutVisible && managerFilterAnchor && (
        <Callout
          target={managerFilterAnchor}
          onDismiss={() => setIsManagerCalloutVisible(false)}
          setInitialFocus
          directionalHint={8}
          isBeakVisible={true}
          gapSpace={5}
        >
          <Stack tokens={{ childrenGap: 4 }} styles={{ root: { padding: 10 } }}>
            <Text variant="mediumPlus">Filter by Manager</Text>
            <Checkbox
              label="All"
              checked={selectedManagers.length === managerFilterOptions.length}
              onChange={() => {
                const all = managerFilterOptions.map(opt => opt.key as string);
                const isAllSelected = selectedManagers.length === all.length;
                setSelectedManagers(isAllSelected ? [] : all);
              }}
            />
            {managerFilterOptions.map(opt => (
              <Checkbox
                key={opt.key}
                label={opt.text}
                checked={selectedManagers.includes(opt.key as string)}
                onChange={() => {
                  const key = opt.key as string;
                  setSelectedManagers(prev =>
                    prev.includes(key) ? prev.filter(k => k !== key) : [...prev, key]
                  );
                }}
              />
            ))}
          </Stack>
        </Callout>
      )}

      {isStatusCalloutVisible && statusFilterAnchor && (
        <Callout
          target={statusFilterAnchor}
          onDismiss={() => setIsStatusCalloutVisible(false)}
          setInitialFocus
          directionalHint={8}
          isBeakVisible={true}
          gapSpace={5}
        >
          <Stack tokens={{ childrenGap: 4 }} styles={{ root: { padding: 10 } }}>
            <Text variant="mediumPlus">Filter by Status</Text>
            <Checkbox
              label="All"
              checked={selectedStatuses.length === statusFilterOptions.length}
              onChange={() => {
                const all = statusFilterOptions.map(opt => opt.key as string);
                const isAllSelected = selectedStatuses.length === all.length;
                setSelectedStatuses(isAllSelected ? [] : all);
              }}
            />
            {statusFilterOptions.map(opt => (
              <Checkbox
                key={opt.key}
                label={opt.text}
                checked={selectedStatuses.includes(opt.key as string)}
                onChange={() => {
                  const key = opt.key as string;
                  setSelectedStatuses(prev =>
                    prev.includes(key) ? prev.filter(k => k !== key) : [...prev, key]
                  );
                }}
              />
            ))}
          </Stack>
        </Callout>
      )}

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

