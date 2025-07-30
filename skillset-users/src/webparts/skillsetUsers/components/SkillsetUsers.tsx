import * as React from 'react'; 
import { useState, useEffect } from 'react';
import { sp } from '@pnp/sp/presets/all';
import { IDropdownOption, Stack, Text, PrimaryButton } from '@fluentui/react';
import LoginForm from './LoginForm';
import HeaderLayout from './HeaderLayout';
import EditProfile from './EditProfile';
import TestPage from './TestPage';
import TicketList from './TicketList';
import RegisterForm from './RegisterForm';
import { ISkillsetUsersProps } from './ISkillsetUsersProps';

const SkillsetUsers: React.FC<ISkillsetUsersProps> = (props) => {
  const [loading, setLoading] = useState(false);
  const [loginError, setLoginError] = useState('');
  const [loginItemId, setLoginItemId] = useState<number | null>(null);
  const [loginForm, setLoginForm] = useState({
    fullName: '', email: '', age: '', skillsets: [] as number[], password: ''
  });
  const [welcomeName, setWelcomeName] = useState('');
  const [userRoles, setUserRoles] = useState<number[]>([]);
  const [selectedRole, setSelectedRole] = useState('');
  const [skillsetOptions, setSkillsetOptions] = useState<IDropdownOption[]>([]);
  const [, setSelectedTestSkill] = useState<number | null>(null);
  const [, setShowTestSection] = useState(false);
  const [view, setView] = useState<'dashboard' | 'edit' | 'test' | 'login' | 'register'>('login');
  const [roleOptions, setRoleOptions] = useState<IDropdownOption[]>([]);

  useEffect(() => {
    const init = async () => {
      try {
        const skillsets = await sp.web.lists.getByTitle("Skillset_Master").items.select("Id", "Title").get();
        setSkillsetOptions(skillsets.map(item => ({ key: item.Id, text: item.Title })));

        const roles = await sp.web.lists.getByTitle("Role_Master").items.select("Id", "Title").get();
        setRoleOptions(roles.map(item => ({ key: item.Id, text: item.Title })));

        const currentUser = await sp.web.currentUser.get();
        setLoginForm(prev => ({ ...prev, email: currentUser.Email }));
      } catch (err) {
        console.error("Initialization error:", err);
      }
    };
    init();
  }, []);

  const hashPassword = async (plainText: string): Promise<string> => {
    const encoder = new TextEncoder();
    const data = encoder.encode(plainText);
    const hashBuffer = await crypto.subtle.digest('SHA-256', data);
    return Array.from(new Uint8Array(hashBuffer)).map(b => b.toString(16).padStart(2, '0')).join('');
  };

  const handleLoginSubmit = async () => {
    setLoading(true);
    setLoginError('');
    try {
      const hashedPassword = await hashPassword(loginForm.password);
      const items = await sp.web.lists.getByTitle("All_Users").items
        .filter(`Email eq '${loginForm.email}' and Password eq '${hashedPassword}'`)
        .select("Id", "Title", "Email", "Age", "Skillset/Id", "User_Role/Id", "User_Role/Title")
        .expand("Skillset", "User_Role")
        .top(1)
        .get();
      if (items.length === 0) {
        setLoginError("Invalid credentials. Please try again.");
      } else {
        const user = items[0];
        const skills = user.Skillset ? user.Skillset.map((s: any) => s.Id) : [];
        const roleIds = user.User_Role ? user.User_Role.map((r: any) => r.Id) : [];
        setLoginItemId(user.Id);
        setLoginForm({
          fullName: user.Title,
          email: user.Email,
          age: user.Age.toString(),
          password: '',
          skillsets: skills
        });
        setWelcomeName(user.Title);
        setUserRoles(roleIds);
        const roleTitles = user.User_Role ? user.User_Role.map((r: any) => r.Title) : [];
        setSelectedRole(roleTitles[0] || '');
        setView('dashboard');
      }
    } catch (err) {
      console.error("Login error:", err);
      setLoginError("Something went wrong. Please try again.");
    }
    setLoading(false);
  };

  const handleInputChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
    const { name } = event.currentTarget;
    setLoginForm({ ...loginForm, [name]: newValue || '' });
  };

  const handleSkillsetChange = (event: React.FormEvent<HTMLDivElement>, option: IDropdownOption): void => {
    setLoginForm(prev => {
      const newSkills = option.selected
        ? [...prev.skillsets, Number(option.key)]
        : prev.skillsets.filter(k => k !== option.key);
      return { ...prev, skillsets: newSkills };
    });
  };

  const handleRoleChange = (event: React.FormEvent<HTMLDivElement>, option: IDropdownOption): void => {
    if (option) {
      const updatedRoles = option.selected
        ? [...userRoles, option.key as number]
        : userRoles.filter(id => id !== option.key);
      setUserRoles(updatedRoles);
    }
  };

  const handleSave = async () => {
    if (!loginItemId) return;
    setLoading(true);
    try {
      await sp.web.lists.getByTitle("All_Users").items.getById(loginItemId).update({
        Title: loginForm.fullName,
        Age: parseInt(loginForm.age),
        SkillsetId: { results: loginForm.skillsets },
        User_RoleId: { results: userRoles }
      });
      setView('dashboard');
    } catch (error) {
      console.error("Error updating user:", error);
    }
    setLoading(false);
  };

  const handleLogout = () => {
    setLoginItemId(null);
    setLoginForm({ fullName: '', email: '', age: '', password: '', skillsets: [] });
    setShowTestSection(false);
    setSelectedTestSkill(null);
    setWelcomeName('');
    setUserRoles([]);
    setSelectedRole('');
    setView('login');
  };

  const selectedRoleTitles = roleOptions
    .filter(opt => userRoles.includes(opt.key as number))
    .map(opt => opt.text as string);

  return (
    <div style={{ marginTop: 20, fontFamily: 'Segoe UI', backgroundColor: '#f4f6f8', padding: 30, minHeight: '100vh' }}>
      {view === 'login' && (
        <>
          <LoginForm
            loginForm={loginForm}
            loginError={loginError}
            loading={loading}
            onInputChange={handleInputChange}
            onSubmit={handleLoginSubmit}
          />
          <Stack horizontalAlign="center" tokens={{ childrenGap: 6 }} styles={{ root: { marginTop: 12 } }}>
            <Text variant="small">New user?</Text>
            <PrimaryButton text="Sign up" onClick={() => setView('register')} styles={{ root: { padding: '0 12px', height: 32, fontSize: 12 } }} />
          </Stack>
        </>
      )}

      {view === 'register' && (
        <RegisterForm
          onBack={() => setView('login')}
          skillsetOptions={skillsetOptions}
        />
      )}

      {view !== 'login' && view !== 'register' && (
        <HeaderLayout
          welcomeName={welcomeName}
          userRole={selectedRoleTitles}
          selectedRole={selectedRole}
          onRoleChange={setSelectedRole}
          onEditClick={() => setView('edit')}
          onTestClick={() => setView('test')}
          onTicketsClick={() => setView('dashboard')}
          onLogout={handleLogout}
        >
          {view === 'dashboard' && (
<TicketList
  welcomeName={welcomeName}
  selectedRole={selectedRole} // ✅ Add this line
  onEditClick={() => setView('edit')}
  onTestClick={() => setView('test')}
  onLogout={handleLogout}
/>

          )}

          {view === 'edit' && (
            <EditProfile
              loginForm={loginForm}
              loading={loading}
              skillsetOptions={skillsetOptions}
              onInputChange={handleInputChange}
              onSkillsetChange={handleSkillsetChange}
              onSave={handleSave}
              onBack={() => setView('dashboard')}
              onLogout={handleLogout}
              userRoles={userRoles}
              roleOptions={roleOptions}
              onRoleChange={handleRoleChange}
            />
          )}

          {view === 'test' && (
            <TestPage
              skillsetOptions={skillsetOptions}
              selectedSkillIds={loginForm.skillsets}
              welcomeName={welcomeName}
              userEmail={loginForm.email}
              onLogout={handleLogout}
              onBack={() => setView('dashboard')}
            />
          )}
        </HeaderLayout>
      )}
    </div>
  );
};

export default SkillsetUsers;
