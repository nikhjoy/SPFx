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

const STATUS = {
  Submitted: 'Submitted',
  ManagerApproved: 'Manager Approved',
  ManagerRejected: 'Manager Rejected',
  ProviderAccepted: 'Provider Accepted',
  ProviderRejected: 'Provider Rejected',
} as const;


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
  type ManagerTab =
    | 'Pending'             // Seeker/Manager old tab
    | 'Approved'            // Seeker/Manager old tab
    | 'Rejected'            // Seeker/Manager old tab
    | 'AllAccepted'         // Provider new tab
    | 'Matching'            // Provider new tab
    | 'ApprovedByYou'       // Provider new tab
    | 'RejectedByYou'       // Provider new tab
    | null;

  const [managerTab, setManagerTab] = useState<ManagerTab>(null);

  const [providerSkillIds, setProviderSkillIds] = useState<number[]>([]);

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

  useEffect(() => {
    const loadProviderSkills = async () => {
      try {
        // All_Users has an Email field and Skillset multi-lookup to Skillset_Master
        const userItems = await sp.web.lists.getByTitle('All_Users')
          .items
          .select('Id', 'Email', 'Skillset/Id')
          .expand('Skillset')
          .filter(`Email eq '${loginEmail}'`)
          .top(1)
          .get();

        const skillIds =
          userItems?.[0]?.Skillset?.map((s: { Id: number }) => s.Id) ?? [];
        setProviderSkillIds(skillIds);
      } catch (e) {
        console.error('❌ loadProviderSkills error:', e);
        setProviderSkillIds([]); // be safe
      }
    };

    if (loginEmail) loadProviderSkills();
  }, [loginEmail]);


  // Auto-select default tab based on role
  useEffect(() => {
    if (selectedRole === 'Support_Provider') {
      setManagerTab('AllAccepted'); // ✅ new provider default
    } else if (
      selectedRole === 'Support_Seeker' ||
      selectedRole === 'Support_Manager'
    ) {
      setManagerTab('Pending');
    } else {
      setManagerTab(null);
    }
  }, [selectedRole]);


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

  const hasSkillOverlap = (t: any, providerSkillIds: number[]) =>
    Array.isArray(t.SkillsetId) && t.SkillsetId.some((id: number) => providerSkillIds.includes(id));



  const fetchTickets = async () => {
    try {
      const items = await sp.web.lists
        .getByTitle('Tickets')
        .items.select(
          'Id', 'Title', 'Description', 'Status', 'AssignedOn', 'SkillsetId',
          'Requestor/Title', 'Requestor/EMail',
          'AssignedTo/Title', 'AssignedTo/EMail',
          'Manager/Title', 'Manager/EMail'
        )
        .expand('Requestor', 'AssignedTo', 'Manager')
        .get();

      const skills = await sp.web.lists.getByTitle('Skillset_Master').items.select('Id', 'Title')();
      const skillMap = new Map(skills.map(s => [s.Id, s.Title]));

      const enriched = items.map(item => {
        // 🔁 Normalize legacy statuses so tabs & filters work consistently
        const rawStatus = String(item.Status || '').trim();
        const normalizedStatus =
          rawStatus.toLowerCase() === 'approved' ? STATUS.ManagerApproved :
            rawStatus.toLowerCase() === 'rejected' ? STATUS.ManagerRejected :
              rawStatus.toLowerCase() === 'submitted' ? STATUS.Submitted :
                rawStatus; // keep other values as-is (e.g., Provider Accepted/Rejected)

        console.log('📌 Requestor:', item.Requestor);
        console.log('📌 AssignedTo:', item.AssignedTo);

        return {
          Id: item.Id,
          Title: item.Title,
          Description: item.Description,
          Status: normalizedStatus, // 👈 use normalized value
          AssignedOn: item.AssignedOn,
          SkillsetId: item.SkillsetId || [],
          Skillset: Array.isArray(item.SkillsetId)
            ? item.SkillsetId.map((id: number) => skillMap.get(id)).filter(Boolean).join(', ')
            : (skillMap.get(item.SkillsetId) || 'Not Assigned'),
          Requestor: item.Requestor?.Title || '',
          RequestorEmail: item.Requestor?.EMail || item.Requestor?.UserPrincipalName || '',
          AssignedTo: item.AssignedTo?.Title || '',
          AssignedToEmail: item.AssignedTo?.EMail || item.AssignedTo?.UserPrincipalName || '',
          Manager: item.Manager?.Title || '',
          ManagerEmail: item.Manager?.EMail || item.Manager?.UserPrincipalName || ''
        };
      });

      console.log('🔍 Raw SharePoint items:', items);
      console.groupCollapsed('FETCH ▶ tickets');
      console.log('login:', (loginEmail || '').toLowerCase(), 'role:', selectedRole);
      console.log('total fetched:', enriched.length);

      console.table(
        enriched.slice(0, 10).map(t => ({
          Id: t.Id,
          Status: t.Status,
          Manager: t.Manager,
          ManagerEmail: (t.ManagerEmail || '').toLowerCase(),
          RequestorEmail: (t.RequestorEmail || '').toLowerCase(),
          AssignedToEmail: (t.AssignedToEmail || '').toLowerCase()
        }))
      );
      console.groupEnd();

      setTickets(enriched);

      const uniqueRequestors = Array.from(new Set(enriched.map(t => t.Requestor).filter(Boolean)));
      const requestorOptions = uniqueRequestors.map(name => ({ key: name, text: name }));
      setRequestorFilterOptions(requestorOptions);
      setSelectedRequestors(requestorOptions.map(opt => opt.key as string)); // all selected

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

  //end of part 1

  // Manager can edit only Approved tickets that they manage
  const canManagerEditApproved = React.useMemo(() => {
    if (!selectedTicket) return false;
    const me = (loginEmail || '').trim().toLowerCase();
    const mgr = (selectedTicket.ManagerEmail || '').trim().toLowerCase();
    return (
      selectedRole === 'Support_Manager' &&
      selectedTicket.Status === STATUS.ManagerApproved &&
      mgr === me
    );
  }, [selectedRole, selectedTicket, loginEmail]);


  const handleSave = async () => {
    try {
      const requestorEmail = requestor[0];
      const assignedToEmail = assignedTo[0];
      let requestorId = null;
      let assignedToId = null;
      let managerId = null;

      // Resolve IDs only if needed
      if (!canManagerEditApproved && requestorEmail) {
        const ensuredRequestor = await sp.web.ensureUser(requestorEmail);
        requestorId = ensuredRequestor.data.Id;

        const managerItems = await sp.web.lists.getByTitle('Manager_Map')
          .items
          .select('ID', 'User_Name/Id', 'User_Name/EMail', 'Manager_Name/Id', 'Manager_Name/EMail')
          .expand('User_Name', 'Manager_Name')
          .filter(`User_Name/EMail eq '${requestorEmail}'`)
          .top(1)
          .get();

        const managerEmail = managerItems?.[0]?.Manager_Name?.EMail;
        if (managerEmail) {
          const ensuredManager = await sp.web.ensureUser(managerEmail);
          managerId = ensuredManager.data.Id;
        }
      }

      if (assignedToEmail) {
        const ensuredAssignedTo = await sp.web.ensureUser(assignedToEmail);
        assignedToId = ensuredAssignedTo.data.Id;
      }

      let data: any;

      if (canManagerEditApproved && selectedTicket) {
        // ✅ Manager editing an Approved ticket: update ONLY AssignedTo + AssignedOn
        data = {
          AssignedOn: assignedOn ? assignedOn.toISOString() : null,
          ...(assignedToId ? { AssignedToId: assignedToId } : {})
        };
      } else {
        // Normal add/edit flow
        data = {
          Title: ticketTitle,
          Description: ticketDescription,
          AssignedOn: assignedOn ? assignedOn.toISOString() : null,
          SkillsetId: { results: selectedSkillset },
          Status: selectedTicket ? selectedTicket.Status : STATUS.Submitted
        };

        if (requestorId) data.RequestorId = requestorId;
        if (assignedToId) data.AssignedToId = assignedToId;
        if (managerId !== null && managerId !== undefined) data.ManagerId = managerId;
      }

      if (selectedTicket) {
        await sp.web.lists.getByTitle('Tickets').items.getById(selectedTicket.Id).update(data);
      } else {
        await sp.web.lists.getByTitle('Tickets').items.add(data);
      }

      closeDialog();
      fetchTickets();
    } catch (error) {
      console.error('❌ handleSave error:', error);
    }
  };


  const openAddDialog = () => {
    setSelectedTicket(null);
    setTicketTitle('');
    setTicketDescription('');
    setSelectedSkillset([]);
    setAssignedOn(undefined);
    // ✅ Prefill Requestor with the web-part login email
    setRequestor(loginEmail ? [loginEmail] : []);
    setAssignedTo([]);
    setIsDialogOpen(true);
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

    const normalizedLogin = (loginEmail || '').trim().toLowerCase();

    // --- 1) Apply TAB FILTERS FIRST ---
    if (managerTab && normalizedLogin) {
      if (selectedRole === 'Support_Manager') {
        if (managerTab === 'Pending') {
          const pending = tickets.filter(
            t =>
              String(t.Status || '') === STATUS.Submitted &&
              String(t.ManagerEmail || '').trim().toLowerCase() === normalizedLogin
          );
          return pending;
        }

        if (managerTab === 'Approved') {
          const mine = tickets.filter(
            t =>
              (
                String(t.Status || '') === STATUS.ManagerApproved ||
                String(t.Status || '') === STATUS.ProviderAccepted
              ) &&
              String(t.ManagerEmail || '').trim().toLowerCase() === normalizedLogin
          );
          return mine;
        }

        if (managerTab === 'Rejected') {
          const mine = tickets.filter(
            t =>
              (
                String(t.Status || '') === STATUS.ManagerRejected ||
                String(t.Status || '') === STATUS.ProviderRejected
              ) &&
              String(t.ManagerEmail || '').trim().toLowerCase() === normalizedLogin
          );
          return mine;
        }

        return result;
      }

      else if (selectedRole === 'Support_Seeker') {
        if (managerTab === 'Pending') {
          result = tickets.filter(
            t =>
              t.Status === STATUS.Submitted &&
              (t.RequestorEmail || '').toLowerCase() === normalizedLogin
          );
        }
        if (managerTab === 'Approved') {
          result = tickets.filter(
            t =>
              (t.Status === STATUS.ManagerApproved || t.Status === STATUS.ProviderAccepted) &&
              (t.RequestorEmail || '').toLowerCase() === normalizedLogin
          );
        }
        if (managerTab === 'Rejected') {
          result = tickets.filter(
            t =>
              (t.Status === STATUS.ManagerRejected || t.Status === STATUS.ProviderRejected) &&
              (t.RequestorEmail || '').toLowerCase() === normalizedLogin
          );
        }
        return result;
      }

      else if (selectedRole === 'Support_Provider') {
        if (managerTab === 'AllAccepted') {
          return (tickets || []).filter(t =>
            String(t.Status ?? '').trim() === STATUS.ManagerApproved
          );
        }

        if (managerTab === 'Matching') {
          return (tickets || []).filter(t =>
            String(t.Status ?? '').trim() === STATUS.ManagerApproved &&
            hasSkillOverlap(t, providerSkillIds) &&
            !String(t.AssignedToEmail ?? '').trim()
          );
        }

        const me = (loginEmail || '').trim().toLowerCase();
        switch (managerTab) {
          case 'ApprovedByYou':
            return tickets.filter(
              t =>
                String(t.Status ?? '').trim() === STATUS.ProviderAccepted &&
                (t.AssignedToEmail || '').trim().toLowerCase() === me
            );
          case 'RejectedByYou':
            // Without a "ProviderActor" field we cannot filter "by you" reliably.
            // Option A: show all Provider Rejected
            return tickets.filter(
              t => String(t.Status ?? '').trim() === STATUS.ProviderRejected
            );
          default:
            return tickets;
        }
      }


      console.log('selectedRole:', selectedRole);
      console.log('managerTab:', managerTab);
    }
    else {
      // --- 2) DEFAULT ROLE VIEWS (no tab selected) ---
      if (selectedRole === 'Support_Seeker') {
        result = result.filter(
          t => (t.RequestorEmail || '').toLowerCase() === normalizedLogin && t.Status !== 'Submitted'
        );
      } else if (selectedRole === 'Support_Manager') {
        const statusOrder: Record<string, number> = {
          [STATUS.Submitted]: 1,
          [STATUS.ManagerApproved]: 2,
          [STATUS.ProviderAccepted]: 3,
          [STATUS.ManagerRejected]: 4,
          [STATUS.ProviderRejected]: 5
        };
        result = [...result].sort((a, b) =>
          (statusOrder[a.Status] ?? 99) - (statusOrder[b.Status] ?? 99)
        );
      } else if (selectedRole === 'Support_Provider') {
        result = result.filter(t => t.Status === STATUS.ManagerApproved);
      }

    }

    // --- 3) Column header filters (unchanged) ---
    // --- 3) Column header filters (with safe skips for Manager tabs) ---
    console.groupCollapsed('FILTER ▶ before column filters');
    console.log('counts:', {
      before: result.length,
      reqSel: selectedRequestors.length,
      asgSel: selectedAssignedTo.length,
      mgrSel: selectedManagers.length,
      stSel: selectedStatuses.length,
      managerTab,
      role: selectedRole
    });
    console.groupEnd();

    // Requestor
    if (selectedRequestors.length > 0) {
      result = result.filter(t => !t.Requestor || selectedRequestors.includes(t.Requestor));
    }

    // Assigned To
    if (selectedAssignedTo.length > 0) {
      result = result.filter(t => !t.AssignedTo || selectedAssignedTo.includes(t.AssignedTo));
    }

    // Manager — SKIP on Manager tabs (Pending/Approved/Rejected) so blank Manager names don't hide rows
    // --- 3) Column header filters ---
    /* ...existing Requestor / AssignedTo filters... */

    // Manager — skip on Manager tabs
    const skipManagerFilter =
      selectedRole === 'Support_Manager' &&
      (managerTab === 'Pending' || managerTab === 'Approved' || managerTab === 'Rejected');

    if (!skipManagerFilter && selectedManagers.length > 0) {
      result = result.filter(t => !t.Manager || selectedManagers.includes(t.Manager));
    }

    // Status — skip on Manager tabs so ProviderAccepted/Rejected aren't hidden
    const skipStatusFilter =
      selectedRole === 'Support_Manager' &&
      (managerTab === 'Pending' || managerTab === 'Approved' || managerTab === 'Rejected');

    if (!skipStatusFilter && selectedStatuses.length > 0) {
      result = result.filter(t => selectedStatuses.includes(String(t.Status || '')));
    }

    // Status
    if (selectedStatuses.length > 0) {
      result = result.filter(t => selectedStatuses.includes(String(t.Status || '')));
    }

    console.groupCollapsed('FILTER ▶ after column filters');
    console.log('after count:', result.length, 'firstIds:', result.slice(0, 5).map(t => t.Id));
    console.groupEnd();

    return result;
  }, [
    tickets,
    selectedRole,
    loginEmail,
    providerSkillIds, // ✅ include
    selectedRequestors,
    selectedAssignedTo,
    selectedManagers,
    selectedStatuses,
    managerTab
  ]);


  const managerGroups = React.useMemo<IGroup[] | undefined>(() => {
    if (selectedRole !== 'Support_Manager') return undefined;

    const buckets: Record<string, any[]> = {
      [STATUS.Submitted]: [],
      [STATUS.ManagerApproved]: [],
      [STATUS.ManagerRejected]: []
    };

    filteredTickets.forEach(t => {
      if (buckets[t.Status]) buckets[t.Status].push(t);
    });

    const order = [STATUS.Submitted, STATUS.ManagerApproved, STATUS.ManagerRejected];
    const groups: IGroup[] = [];
    let startIndex = 0;

    for (const status of order) {
      const arr = buckets[status];
      if (arr.length > 0) {
        groups.push({
          key: status,
          name: `${status} Requests`,
          startIndex,
          count: arr.length,
          level: 0,
          isCollapsed: false
        });
        startIndex += arr.length;
      }
    }

    // ⚠️ Important: if no groups, return undefined so DetailsList shows items normally
    return groups.length ? groups : undefined;
  }, [selectedRole, filteredTickets]);

  //end of part 2


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
      key: 'col5',
      name: 'Assigned On',
      fieldName: 'AssignedOn',
      minWidth: 120,
      onRender: (item: any) => (
        <span>{item.AssignedOn ? new Date(item.AssignedOn).toLocaleDateString() : '—'}</span>
      ),
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
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
          {/* Always: View */}
          <IconButton iconProps={{ iconName: 'View' }} onClick={() => openViewDialog(item)} />

          {/* Support_Manager: show Approve/Reject only on Submitted */}
          {selectedRole === 'Support_Manager' && item.Status === STATUS.Submitted && (
            <>
              <PrimaryButton
                text="Approve"
                onClick={() => handleApproval(item.Id, 'Approved')}
                styles={{ root: { backgroundColor: 'green', color: 'white', padding: '0 8px', minWidth: 90 } }}
              />
              <DefaultButton
                text="Reject"
                onClick={() => handleApproval(item.Id, 'Rejected')}
                styles={{ root: { backgroundColor: 'red', color: 'white', padding: '0 8px', minWidth: 90 } }}
              />
            </>
          )}

          {/* Support_Provider: ONLY Accept/Reject as Provider (NO edit/delete) */}
          {selectedRole === 'Support_Provider' &&
            item.Status === STATUS.ManagerApproved &&
            (!item.AssignedToEmail || item.AssignedToEmail.trim() === '') &&
            Array.isArray(item.SkillsetId) &&
            item.SkillsetId.some((id: number) => providerSkillIds.includes(id)) && (
              <>
                <PrimaryButton
                  text="Provider Accept"
                  onClick={() => acceptRequest(item.Id)}
                  styles={{ root: { padding: '0 8px', minWidth: 140 } }}
                />
                <DefaultButton
                  text="Provider Reject"
                  onClick={() => rejectAsProvider(item.Id)}
                  styles={{ root: { padding: '0 8px', minWidth: 140 } }}
                />
              </>
            )
          }

          {/* Everyone else (NOT Support_Provider): Edit/Delete as before */}
          {selectedRole !== 'Support_Provider' && (
            <>
              {(
                selectedRole !== 'Support_Manager' ||
                (selectedRole === 'Support_Manager' &&
                  item.Status === STATUS.ManagerApproved &&
                  ((item.ManagerEmail || '').toLowerCase() === (loginEmail || '').toLowerCase()))
              ) && (
                  <IconButton iconProps={{ iconName: 'Edit' }} onClick={() => openEditDialog(item)} />
                )}

              <IconButton
                iconProps={{ iconName: 'Delete' }}
                onClick={() => { setTicketToDelete(item); setDeleteConfirmOpen(true); }}
              />
            </>
          )}

          {/* Delete confirmation dialog stays here */}
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

  const acceptRequest = async (ticketId: number) => {
    try {
      const ensured = await sp.web.ensureUser(loginEmail);
      await sp.web.lists.getByTitle('Tickets').items.getById(ticketId).update({
        AssignedToId: ensured.data.Id,
        AssignedOn: new Date().toISOString(),
        Status: STATUS.ProviderAccepted
      });
      setActionMessage('Ticket accepted and assigned to you.');
      await fetchTickets();
      setTimeout(() => setActionMessage(null), 3000);
    } catch (e) {
      console.error('❌ acceptRequest error:', e);
      setActionMessage('Failed to accept ticket.');
      setTimeout(() => setActionMessage(null), 3000);
    }
  };

  const rejectAsProvider = async (ticketId: number) => {
    try {
      await sp.web.lists.getByTitle('Tickets').items.getById(ticketId).update({
        Status: STATUS.ProviderRejected
      });
      setActionMessage('Ticket rejected.');
      await fetchTickets();
      setTimeout(() => setActionMessage(null), 3000);
    } catch (e) {
      console.error('❌ rejectAsProvider error:', e);
      setActionMessage('Failed to reject ticket.');
      setTimeout(() => setActionMessage(null), 3000);
    }
  };


  const handleApproval = async (ticketId: number, decision: 'Approved' | 'Rejected') => {
    try {
      const newStatus =
        decision === 'Approved' ? STATUS.ManagerApproved : STATUS.ManagerRejected;
      const ensuredMgr = await sp.web.ensureUser(loginEmail);

      await sp.web.lists.getByTitle('Tickets').items.getById(ticketId).update({ Status: newStatus, ManagerId: ensuredMgr.data.Id });
      setActionMessage(`Ticket ${decision} successfully.`);
      await fetchTickets();
      setTimeout(() => setActionMessage(null), 3000);
    } catch (error) {
      console.error(`Error updating status to ${decision}:`, error);
      setActionMessage(`Failed to update status.`);
      setTimeout(() => setActionMessage(null), 3000);
    }
  };


  console.log("📌 Callout visible:", isRequestorFilterCalloutVisible);
  console.log("📌 Anchor element:", requestorFilterAnchor);



  const itemsForList = React.useMemo(() => {
    if (!managerGroups) return filteredTickets;

    // Keep item order aligned with groups (Submitted → Manager Approved → Manager Rejected)
    const ordered = [
      ...filteredTickets.filter(t => t.Status === STATUS.Submitted),
      ...filteredTickets.filter(t => t.Status === STATUS.ManagerApproved),
      ...filteredTickets.filter(t => t.Status === STATUS.ManagerRejected),
    ];

    console.groupCollapsed('GROUP ▶ ordered items');
    console.log({
      submitted: ordered.filter(t => t.Status === STATUS.Submitted).length,
      managerApproved: ordered.filter(t => t.Status === STATUS.ManagerApproved).length,
      managerRejected: ordered.filter(t => t.Status === STATUS.ManagerRejected).length
    });
    console.groupEnd();

    return ordered;
  }, [filteredTickets, managerGroups]);


  return (
    <Stack tokens={{ childrenGap: 15 }}>

<Stack horizontal verticalAlign="center" tokens={{ childrenGap: 12 }} wrap={false}>
  {/* LEFT: Role-based Tabs */}
  {selectedRole === 'Support_Provider' ? (
    <Stack horizontal tokens={{ childrenGap: 0 }} wrap={false} styles={{ root: { borderBottom: '2px solid #ddd' } }}>
      {[
        { key: 'AllAccepted', text: 'All Admin Accepted' },
        { key: 'Matching', text: 'Your Matching' },
        { key: 'ApprovedByYou', text: 'Approved by You' },
        { key: 'RejectedByYou', text: 'Rejected by You' },
      ].map(tab => (
        <DefaultButton
          key={tab.key}
          text={tab.text}
          onClick={() => setManagerTab(tab.key as any)}
          styles={{
            root: {
              borderRadius: 0,
              border: 'none',
              borderBottom: managerTab === tab.key ? '3px solid #0078d4' : '3px solid transparent',
              background: 'transparent',
              fontWeight: managerTab === tab.key ? 600 : 400,
              color: managerTab === tab.key ? '#0078d4' : '#333',
              padding: '8px 16px',
              marginRight: 12,
            },
          }}
        />
      ))}
    </Stack>
  ) : (
    <Stack horizontal tokens={{ childrenGap: 0 }} wrap={false} styles={{ root: { borderBottom: '2px solid #ddd' } }}>
      {[
        {
          key: 'Pending',
          text: selectedRole === 'Support_Manager' ? 'Pending Requests' : 'Submitted Tickets',
        },
        { key: 'Approved', text: 'Approved' },
        { key: 'Rejected', text: 'Rejected' },
      ].map(tab => (
        <DefaultButton
          key={tab.key}
          text={tab.text}
          onClick={() => setManagerTab(tab.key as any)}
          styles={{
            root: {
              borderRadius: 0,
              border: 'none',
              borderBottom: managerTab === tab.key ? '3px solid #0078d4' : '3px solid transparent',
              background: 'transparent',
              fontWeight: managerTab === tab.key ? 600 : 400,
              color: managerTab === tab.key ? '#0078d4' : '#333',
              padding: '8px 16px',
              marginRight: 12,
            },
          }}
        />
      ))}
    </Stack>
  )}
        {/* RIGHT: Clear + Add (Add only for Seeker) */}
        <Stack horizontal tokens={{ childrenGap: 10 }} wrap={false}>
          <PrimaryButton
            text="Clear Filters"
            onClick={() => {
              setSelectedRequestors(requestorFilterOptions.map(opt => opt.key as string));
              setSelectedAssignedTo(assignedToFilterOptions.map(opt => opt.key as string));
              setSelectedManagers(managerFilterOptions.map(opt => opt.key as string));
              setSelectedStatuses(statusFilterOptions.map(opt => opt.key as string));
              setManagerTab(selectedRole === 'Support_Provider' ? 'AllAccepted' : 'Pending');
              setIsRequestorFilterCalloutVisible(false);
              setIsAssignedToCalloutVisible(false);
              setIsManagerCalloutVisible(false);
              setIsStatusCalloutVisible(false);
            }}
            styles={{ root: { backgroundColor: '#d83b01', color: 'white', padding: '0 12px', minWidth: 120 } }}
          />

          {selectedRole === 'Support_Seeker' && (
            <PrimaryButton text="+ Add Ticket" onClick={openAddDialog} />
          )}
        </Stack>
      </Stack>

      {/* ✅ Success message appears here */}
      {actionMessage && (
        <Text variant="medium" styles={{ root: { color: 'green', fontWeight: 600 } }}>
          {actionMessage}
        </Text>
      )}

{itemsForList.length === 0 ? (
  <Text variant="medium" styles={{ root: { marginTop: 8, fontStyle: 'italic' } }}>
    Currently there are no tickets.
  </Text>
) : (
  <DetailsList
    items={itemsForList}
    columns={columns}
    groups={managerGroups}
    layoutMode={DetailsListLayoutMode.fixedColumns}
  />
)}

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
          <TextField
            label="Title"
            value={ticketTitle}
            onChange={(e, v) => setTicketTitle(v || '')}
            required
            disabled={canManagerEditApproved}
          />

          <TextField
            label="Description"
            multiline
            rows={3}
            value={ticketDescription}
            onChange={(e, v) => setTicketDescription(v || '')}
            disabled={canManagerEditApproved}
          />

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
            disabled={canManagerEditApproved}
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
            disabled={canManagerEditApproved}
          />


          {selectedRole !== 'Support_Seeker' && (
            <>
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
              /* no disabled -> editable */
              />

              <DatePicker
                label="Assigned On"
                value={assignedOn}
                onSelectDate={(date) => setAssignedOn(date || undefined)}
              /* no disabled -> editable */
              />
            </>
          )}


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

