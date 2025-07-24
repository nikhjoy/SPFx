import * as React from 'react';
import { Stack, PrimaryButton } from '@fluentui/react';

export interface IDashboardViewProps {
  welcomeName: string;
  onEditClick: () => void;
  onTestClick: () => void;
  onLogout: () => void;
}

const DashboardView: React.FC<IDashboardViewProps> = ({
  welcomeName,
  onEditClick,
  onTestClick,
  onLogout
}) => {
  return (
    <div style={{
      background: 'linear-gradient(to right, #0066cc, #3399ff)',
      padding: '20px',
      borderRadius: '8px',
      color: 'white',
      marginBottom: '20px'
    }}>
      <Stack horizontal horizontalAlign="space-between">
        <span style={{ fontSize: 24, fontWeight: 600 }}>
          Welcome to your Dashboard, {welcomeName}
        </span>
        <PrimaryButton
          text="Logout"
          onClick={onLogout}
          styles={{
            root: {
              backgroundColor: '#ffffff',
              color: '#0066cc',
              fontWeight: 'bold',
              border: 'none'
            }
          }}
        />
      </Stack>

      <div style={{
        marginTop: 30,
        padding: 20,
        backgroundColor: 'white',
        color: '#333',
        borderRadius: 8,
        boxShadow: '0 4px 10px rgba(0,0,0,0.1)'
      }}>
        <h2 style={{ marginTop: 0 }}>Quick Actions</h2>
        <Stack horizontal tokens={{ childrenGap: 20 }}>
          <PrimaryButton
            text="Edit Profile"
            onClick={onEditClick}
            styles={{ root: { height: 40, padding: '0 30px' } }}
          />
          <PrimaryButton
            text="Take Test"
            onClick={onTestClick}
            styles={{ root: { height: 40, padding: '0 30px' } }}
          />
        </Stack>
      </div>
    </div>
  );
};

export default DashboardView;
