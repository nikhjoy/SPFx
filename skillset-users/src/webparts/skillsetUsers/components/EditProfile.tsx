import * as React from 'react';
import {
  TextField,
  Dropdown,
  IDropdownOption,
  PrimaryButton,
  DialogFooter,
  Stack
} from '@fluentui/react';

export interface IEditProfileProps {
  loginForm: {
    fullName: string;
    email: string;
    age: string;
    skillsets: number[];
  };
  loading: boolean;
  skillsetOptions: IDropdownOption[];
  onInputChange: (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => void;
  onSkillsetChange: (event: React.FormEvent<HTMLDivElement>, option: IDropdownOption) => void;
  onSave: () => void;
  onBack: () => void;
  onLogout: () => void; // ✅ Add this line
}

const EditProfile: React.FC<IEditProfileProps> = ({
  loginForm,
  loading,
  skillsetOptions,
  onInputChange,
  onSkillsetChange,
  onSave,
  onBack
}) => {
  return (
    <div style={{ fontFamily: 'Segoe UI' }}>
      {/* Gradient Header */}
      <div style={{
        background: 'linear-gradient(to right, #0066cc, #3399ff)',
        padding: '20px',
        borderRadius: '8px',
        color: 'white',
        marginBottom: '20px'
      }}>
        <Stack horizontal horizontalAlign="space-between">
          <span style={{ fontSize: 24, fontWeight: 600 }}>
            Edit Your Profile
          </span>
        </Stack>
      </div>

      {/* White Card Section */}
      <div style={{
        backgroundColor: 'white',
        padding: 20,
        borderRadius: 8,
        boxShadow: '0 4px 10px rgba(0,0,0,0.1)'
      }}>
        <h2 style={{ marginTop: 0 }}>Update Your Information</h2>
        <TextField label="Full Name" name="fullName" value={loginForm.fullName} onChange={onInputChange} required />
        <TextField label="Email" value={loginForm.email} readOnly />
        <TextField label="Age" name="age" value={loginForm.age} onChange={onInputChange} type="number" required />
        <Dropdown
          label="Skillsets"
          placeholder="Select skillsets"
          multiSelect
          options={skillsetOptions}
          selectedKeys={loginForm.skillsets}
          onChange={onSkillsetChange}
        />
        <DialogFooter>
          <PrimaryButton text="Save" onClick={onSave} disabled={loading || !loginForm.fullName || !loginForm.age} />
          <PrimaryButton text="Back to Dashboard" onClick={onBack} />
        </DialogFooter>
      </div>
    </div>
  );
};

export default EditProfile;
