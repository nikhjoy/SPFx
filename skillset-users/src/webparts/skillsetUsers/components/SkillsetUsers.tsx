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
import { FormEvent } from 'react';

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
  const [reloadKey, setReloadKey] = useState(0);

const handleProfileUpdate = async () => {
  const currentUser = await sp.web.currentUser.get();
  const userEmail = currentUser.Email;

  const users = await sp.web.lists.getByTitle("All_Users").items
    .filter(`Email eq '${userEmail}'`)
    .select("Id", "User_Role/Title")
    .expand("User_Role")
    .top(1)
    .get();

  const updatedRoles: string[] = users[0]?.User_Role?.map((r: any) => r.Title) || [];

  console.log("🔄 Updated roles from SharePoint:", updatedRoles);

  // ✅ Forcefully assign fallback role — skip the includes() check
  let fallbackRole = '';

  if (updatedRoles.includes("Support_Seeker")) {
    fallbackRole = "Support_Seeker";
  } else if (updatedRoles.includes("Support_Provider")) {
    fallbackRole = "Support_Provider";
  } else if (updatedRoles.includes("Support_Manager")) {
    fallbackRole = "Support_Manager";
  }

  console.log("✅ Setting selectedRole to:", fallbackRole);
  setSelectedRole(fallbackRole);

  setReloadKey(prev => prev + 1);
  setView("dashboard");
};


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

const handleRoleChange = (
  event: FormEvent<HTMLDivElement>,
  option: IDropdownOption
): void => {
  const roleKey = option.key as number;

  const updatedRoles = option.selected
    ? [...userRoles, roleKey]
    : userRoles.filter(id => id !== roleKey);

  setUserRoles(updatedRoles);

  // 🔍 Get text of deselected role
  const deselectedRoleTitle = roleOptions.find(r => r.key === roleKey)?.text;

  // ✅ If deselected role was currently selected dropdown, pick fallback
  if (
    !option.selected && // role was deselected
    deselectedRoleTitle === selectedRole // the dropdown currently showed this
  ) {
    // Find fallback role
    const fallback = roleOptions.find(
      r => updatedRoles.includes(r.key as number) &&
           (r.text === "Support_Seeker" || r.text === "Support_Provider")
    );

    if (fallback) {
      console.log("🔁 Auto-updating dropdown to fallback:", fallback.text);
      setSelectedRole(fallback.text);
    } else {
      console.log("⚠️ No fallback role available. Clearing dropdown.");
      setSelectedRole('');
    }
  }
};


  const handleSave = async (): Promise<void> => {
    if (!loginItemId) return;
    setLoading(true);
    try {
      console.log("📝 Final roles to be saved:", userRoles);
      await sp.web.lists.getByTitle("All_Users").items.getById(loginItemId).update({
        Title: loginForm.fullName,
        Age: parseInt(loginForm.age),
        SkillsetId: { results: loginForm.skillsets },
        User_RoleId: { results: userRoles }
      });
      console.log("📝 Saving roles to SharePoint:", userRoles);
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
    onRoleChange={(role: string) => {
      console.log("🔁 Role dropdown changed to:", role);
      setSelectedRole(role);
    }}
    onEditClick={() => setView('edit')}
    onTestClick={() => setView('test')}
    onTicketsClick={() => setView('dashboard')}
    onLogout={handleLogout}
  >

          {view === 'dashboard' && (
            <>
            {console.log("🎯 TicketList is rendering with selectedRole =", selectedRole)}
<TicketList
  key={reloadKey} 
  welcomeName={welcomeName}
  selectedRole={selectedRole}
  loginEmail={loginForm.email}
  context={{
    spHttpClient: props.context.spHttpClient,
    msGraphClientFactory: props.context.msGraphClientFactory,
    absoluteUrl: props.context.pageContext.web.absoluteUrl
  }}
  onEditClick={() => setView('edit')}
  onTestClick={() => setView('test')}
  onLogout={handleLogout}
/>
</>
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
              onProfileUpdate={handleProfileUpdate}
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
