import * as React from 'react';
import { Text, IconButton, IContextualMenuProps } from '@fluentui/react';

export interface IHeaderLayoutProps {
  welcomeName: string;
  onEditClick: () => void;
  onTestClick: () => void;
  onLogout: () => void;
  children: React.ReactNode;
}

const HeaderLayout: React.FC<IHeaderLayoutProps> = ({
  welcomeName,
  onEditClick,
  onTestClick,
  onLogout,
  children
}) => {
  const settingsMenuProps: IContextualMenuProps = {
    items: [
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
        key: 'logout',
        text: 'Logout',
        iconProps: { iconName: 'SignOut' },
        onClick: onLogout
      }
    ]
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
        <IconButton
          iconProps={{ iconName: 'Settings' }}
          title="Settings"
          ariaLabel="Settings"
          menuProps={settingsMenuProps}
          styles={{ root: { color: 'white' } }}
        />
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
