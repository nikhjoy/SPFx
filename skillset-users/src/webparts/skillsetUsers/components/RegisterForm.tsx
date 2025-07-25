import * as React from 'react';
import { useEffect, useState } from 'react';
import { sp } from '@pnp/sp/presets/all';
import {
  TextField,
  PrimaryButton,
  Dropdown,
  IDropdownOption,
  Stack,
  Text,
  MessageBar,
  MessageBarType
} from '@fluentui/react';

interface IRegisterFormProps {
  onBack: () => void;
  skillsetOptions: IDropdownOption[];
}

const RegisterForm: React.FC<IRegisterFormProps> = ({ onBack, skillsetOptions }) => {
  const [form, setForm] = useState({
    fullName: '', email: '', age: '', password: '', confirmPassword: '', skillsets: [] as number[], roleIds: [] as number[]
  });
  const [roleOptions, setRoleOptions] = useState<IDropdownOption[]>([]);
  const [loading, setLoading] = useState(false);
  const [success, setSuccess] = useState(false);
  const [error, setError] = useState('');

  useEffect(() => {
    const fetchRoles = async () => {
      try {
        const roles = await sp.web.lists.getByTitle('Role_Master').items.select('Id', 'Title').get();
        const filtered = roles.filter(r => r.Title !== 'App_Admin');
        setRoleOptions(filtered.map(role => ({ key: role.Id, text: role.Title })));
      } catch (err) {
        console.error('Error loading roles:', err);
      }
    };
    fetchRoles();
  }, []);

  const handleInput = (e: any, newValue?: string) => {
    const { name } = e.target;
    setForm({ ...form, [name]: newValue || '' });
  };

  const handleSkillChange = (event: any, option: IDropdownOption) => {
    const newSkills = option.selected
      ? [...form.skillsets, Number(option.key)]
      : form.skillsets.filter(id => id !== option.key);
    setForm({ ...form, skillsets: newSkills });
  };

  const handleRoleChange = (event: any, option: IDropdownOption) => {
    const newRoles = option.selected
      ? [...form.roleIds, Number(option.key)]
      : form.roleIds.filter(id => id !== option.key);
    setForm({ ...form, roleIds: newRoles });
  };

  const handleRegister = async () => {
    setError('');
    setSuccess(false);
    if (form.password !== form.confirmPassword) {
      setError("Passwords do not match");
      return;
    }
    if (!form.fullName || !form.email || !form.age || !form.password || form.roleIds.length === 0 || form.skillsets.length === 0) {
      setError("All fields are required");
      return;
    }
    try {
      setLoading(true);
      const encoder = new TextEncoder();
      const data = encoder.encode(form.password);
      const hashBuffer = await crypto.subtle.digest('SHA-256', data);
      const hashedPassword = Array.from(new Uint8Array(hashBuffer)).map(b => b.toString(16).padStart(2, '0')).join('');

      await sp.web.lists.getByTitle("All_Users").items.add({
        Title: form.fullName,
        Email: form.email,
        Age: parseInt(form.age),
        Password: hashedPassword,
        SkillsetId: { results: form.skillsets },
        User_RoleId: { results: form.roleIds }
      });
      setSuccess(true);
      setForm({ fullName: '', email: '', age: '', password: '', confirmPassword: '', skillsets: [], roleIds: [] });
    } catch (err) {
      console.error("Registration failed:", err);
      setError("Registration failed. Please try again.");
    } finally {
      setLoading(false);
    }
  };

  return (
    <Stack tokens={{ childrenGap: 15 }} styles={{ root: { maxWidth: 500, margin: 'auto' } }}>
      <Text variant="large">Register to Support</Text>

      {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}
      {success && <MessageBar messageBarType={MessageBarType.success}>Account created successfully. Please login.</MessageBar>}

      <TextField label="Full Name" name="fullName" value={form.fullName} onChange={handleInput} required />
      <TextField label="Email" name="email" value={form.email} onChange={handleInput} required />
      <TextField label="Age" name="age" value={form.age} onChange={handleInput} required />
      <TextField label="Password" name="password" type="password" value={form.password} onChange={handleInput} canRevealPassword required />
      <TextField label="Confirm Password" name="confirmPassword" type="password" value={form.confirmPassword} onChange={handleInput} canRevealPassword required />

      <Dropdown
        label="Select Skillsets"
        multiSelect
        options={skillsetOptions}
        selectedKeys={form.skillsets}
        onChange={handleSkillChange}
        required
      />

      <Dropdown
        label="Select Role(s)"
        multiSelect
        options={roleOptions}
        selectedKeys={form.roleIds}
        onChange={handleRoleChange}
        required
      />

      <Stack horizontal tokens={{ childrenGap: 10 }}>
        <PrimaryButton text="Register" onClick={handleRegister} disabled={loading} />
        <PrimaryButton text="Back to Login" onClick={onBack} disabled={loading} />
      </Stack>
    </Stack>
  );
};

export default RegisterForm;
