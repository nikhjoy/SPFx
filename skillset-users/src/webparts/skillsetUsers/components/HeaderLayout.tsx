import * as React from 'react';
import { Text, IconButton, IContextualMenuProps, Stack, Label, Dropdown, IDropdownOption } from '@fluentui/react';

export interface IHeaderLayoutProps {
  welcomeName: string;
  onEditClick: () => void;
  onTestClick: () => void;
  onTicketsClick: () => void;
  onLogout: () => void;
  onRateUsersClick?: (ticket?: any | null) => void; // allow optional ticket param
  userRole?: string[];
  selectedRole?: string | number; // accept string or numeric key
  onRoleChange?: (role: string) => void;
  children: React.ReactNode;
}

const HeaderLayout: React.FC<IHeaderLayoutProps> = ({
  welcomeName,
  onEditClick,
  onTestClick,
  onTicketsClick,
  onLogout,
  onRateUsersClick,
  userRole,
  selectedRole,
  onRoleChange,
  children
}) => {
  const settingsMenuProps: IContextualMenuProps = {
    items: (() => {
      const base = [
        {
          key: 'userDetails',
          text: 'User Details',
          iconProps: { iconName: 'Contact' },
          onClick: onEditClick
        },
        {
          key: 'skillsetDashboard',
          text: 'Skillset Dashboard',
          iconProps: { iconName: 'TestBeaker' },
          onClick: onTestClick
        },
        {
          key: 'tickets',
          text: 'Tickets',
          iconProps: { iconName: 'ReportDocument' },
          onClick: onTicketsClick
        }
      ];

      if (String(selectedRole) === 'Support_Manager') {
        base.push({
          key: 'rateUsers',
          text: 'Rate Users',
          iconProps: { iconName: 'FavoriteStar' },
          onClick: () => {
            onRateUsersClick?.(null);
          }
        });
      }

      base.push({
        key: 'logout',
        text: 'Logout',
        iconProps: { iconName: 'SignOut' },
        onClick: onLogout
      });

      return base;
    })()
  };

  const renderRole = () => {
    if (!userRole || userRole.length === 0) return null;
    if (userRole.length === 1) {
      return (
        <Label styles={{ root: { color: 'white', fontWeight: 500, fontSize: 14, paddingRight: 10 } }}>{userRole[0]}</Label>
      );
    }
    const roleOptions: IDropdownOption[] = userRole.map(role => ({ key: role, text: role }));
    return (
      <Dropdown
        label=""
        options={roleOptions}
        selectedKey={selectedRole}
        onChange={(e, option) => onRoleChange && onRoleChange(String(option?.key))}
        styles={{
          dropdown: { minWidth: 150 },
          title: { background: 'white', color: '#333' }
        }}
      />
    );
  };

  return (
    <div style={{ padding: 20 }}>
      <div style={{
        background: 'linear-gradient(to right, #0066cc, #3399ff)',
        padding: 20,
        borderRadius: 8,
        color: 'white',
        marginBottom: 20,
        display: 'flex',
        justifyContent: 'space-between',
        alignItems: 'center'
      }}>
        <Text variant="xLarge" styles={{ root: { fontWeight: 600 } }}>
          Welcome to your Dashboard, {welcomeName}
        </Text>

        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 12 }} horizontalAlign="end" style={{ flexShrink: 0 }}>
          {renderRole()}
          <IconButton
            iconProps={{ iconName: 'Settings' }}
            title="Settings"
            ariaLabel="Settings"
            menuProps={settingsMenuProps}
            styles={{ root: { color: 'white' } }}
            menuIconProps={{ style: { display: 'none' } }}
          />
        </Stack>
      </div>

      <div style={{
        padding: 20,
        backgroundColor: 'white',
        color: '#333',
        borderRadius: 8,
        boxShadow: '0 4px 10px rgba(0,0,0,0.1)'
      }}>
        {children}
      </div>
    </div>
  );
};

export default HeaderLayout;
