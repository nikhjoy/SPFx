import * as React from 'react';
import {
  TextField,
  Dropdown,
  IDropdownOption,
  PrimaryButton,
  DialogFooter
} from '@fluentui/react';

export interface IEditProfileProps {
  loginForm: {
    fullName: string;
    email: string;
    age: string;
    skillsets: number[];
  };

  userRoles: number[];
  roleOptions: IDropdownOption[];
  onRoleChange: (event: React.FormEvent<HTMLDivElement>, option: IDropdownOption) => void;

  loading: boolean;
  skillsetOptions: IDropdownOption[];
  onInputChange: (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => void;
  onSkillsetChange: (event: React.FormEvent<HTMLDivElement>, option: IDropdownOption) => void;
  onSave: () => void;
  onBack: () => void;
  onLogout: () => void;
  onEditClick?: () => void;
  onTestClick?: () => void;
  onProfileUpdate: () => void;
}

const EditProfile: React.FC<IEditProfileProps> = ({
  loginForm,
  loading,
  skillsetOptions,
  onInputChange,
  onSkillsetChange,
  onSave,
  onBack,
  onLogout,
  onEditClick,
  onTestClick,
  userRoles,
  roleOptions,
  onRoleChange,
  onProfileUpdate
}) => {
  return (
    <>
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
      <Dropdown
        label="User Roles"
        placeholder="Select roles"
        multiSelect
        options={roleOptions}
        selectedKeys={userRoles}
        onChange={onRoleChange}
      />


      <DialogFooter>
        <PrimaryButton
          text="Save"
          onClick={async () => {
            await onSave();
            onProfileUpdate();
          }}
          disabled={loading || !loginForm.fullName || !loginForm.age}
        />
        <PrimaryButton text="Back to Dashboard" onClick={onBack} />
      </DialogFooter>
    </>
  );
};

export default EditProfile;
