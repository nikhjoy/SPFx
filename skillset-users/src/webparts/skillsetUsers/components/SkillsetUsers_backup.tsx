import * as React from 'react';
import { useState, useEffect } from 'react';
import {
  PrimaryButton, TextField, Dropdown, IDropdownOption, Stack,
  DialogFooter, MessageBar, MessageBarType, Label} from '@fluentui/react';
import { sp } from '@pnp/sp/presets/all';
import { ISkillsetUsersProps } from './ISkillsetUsersProps';

const SkillsetUsers: React.FC<ISkillsetUsersProps> = (props) => {
  const [loading, setLoading] = useState(false);
  const [] = useState(false);
  const [loginError, setLoginError] = useState('');
  const [loginItemId, setLoginItemId] = useState<number | null>(null);
  const [loginForm, setLoginForm] = useState({
    fullName: '', email: '', age: '', skillsets: [] as number[], password: ''
  });
  const [, setEditSuccess] = useState(false);
  const [welcomeName, setWelcomeName] = useState('');
  const [skillsetOptions, setSkillsetOptions] = useState<IDropdownOption[]>([]);
  const [selectedTestSkill, setSelectedTestSkill] = useState<number | null>(null);
  const [showTestSection, setShowTestSection] = useState(false);
  const [view, setView] = useState<'dashboard' | 'edit' | 'test' | 'login'>('login');

  useEffect(() => {
    const init = async () => {
      try {
        const skillsets = await sp.web.lists.getByTitle("Skillset_Master").items.select("Id", "Title").get();
        setSkillsetOptions(skillsets.map(item => ({ key: item.Id, text: item.Title })));
        const currentUser = await sp.web.currentUser.get();
        setLoginForm(prev => ({ ...prev, email: currentUser.Email }));
      } catch (err) {
        console.error("Initialization error:", err);
      }
    };
    init();
  }, []);

const validateEmail = (value: string): string => {
  if (!value) return '';
  const emailRegex = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+$/;
  return emailRegex.test(value) ? '' : 'Invalid email address';
};


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
        .select("Id", "Title", "Email", "Age", "Skillset/Id")
        .expand("Skillset")
        .top(1)
        .get();
      if (items.length === 0) {
        setLoginError("Invalid credentials. Please try again.");
      } else {
        const user = items[0];
        const skills = user.Skillset ? user.Skillset.map((s: any) => s.Id) : [];
        setLoginItemId(user.Id);
        setLoginForm({
          fullName: user.Title,
          email: user.Email,
          age: user.Age.toString(),
          password: '',
          skillsets: skills
        });
        setWelcomeName(user.Title);
        setView('dashboard');
      }
    } catch (err) {
      console.error("Login error:", err);
      setLoginError("Something went wrong. Please try again.");
    }
    setLoading(false);
  };

  const handleLoginInputChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
    const { name } = event.currentTarget;
    setLoginForm({ ...loginForm, [name]: newValue });
  };

  const handleLoginSkillsetChange = (event: React.FormEvent<HTMLDivElement>, option: IDropdownOption): void => {
    setLoginForm(prev => {
      const newSkills = option.selected
        ? [...prev.skillsets, Number(option.key)]
        : prev.skillsets.filter(k => k !== option.key);
      return { ...prev, skillsets: newSkills };
    });
  };

  const handleLoginSave = async () => {
    if (!loginItemId) return;
    setLoading(true);
    try {
      await sp.web.lists.getByTitle("All_Users").items.getById(loginItemId).update({
        Title: loginForm.fullName,
        Age: parseInt(loginForm.age),
        SkillsetId: { results: loginForm.skillsets }
      });
      setEditSuccess(true);
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
    setView('login');
  };

  const renderLogin = () => (
    <Stack tokens={{ childrenGap: 15 }}>
      <TextField label="Email" name="email" value={loginForm.email} onChange={(e, v) => setLoginForm(prev => ({ ...prev, email: v || '' }))} required onBlur={() => setLoginError(validateEmail(loginForm.email))} errorMessage={loginError} />
      <TextField label="Password" name="password" type="password" canRevealPassword value={loginForm.password} onChange={(e, v) => setLoginForm(prev => ({ ...prev, password: v || '' }))} required />
      {loginError && (
        <MessageBar messageBarType={MessageBarType.error} isMultiline={false} onDismiss={() => setLoginError('')}>
          {loginError}
        </MessageBar>
      )}
      <PrimaryButton text="Login" onClick={handleLoginSubmit} disabled={!loginForm.email || !loginForm.password || loading} />
    </Stack>
  );

const renderDashboard = () => (
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
        onClick={handleLogout}
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
          onClick={() => setView('edit')}
          styles={{ root: { height: 40, padding: '0 30px' } }}
        />
        <PrimaryButton
          text="Take Test"
          onClick={() => setView('test')}
          styles={{ root: { height: 40, padding: '0 30px' } }}
        />
      </Stack>
    </div>
  </div>
);


const renderEditProfile = () => (
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
        <PrimaryButton
          text="Logout"
          onClick={handleLogout}
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
    </div>

    {/* White Card Section */}
    <div style={{
      backgroundColor: 'white',
      padding: 20,
      borderRadius: 8,
      boxShadow: '0 4px 10px rgba(0,0,0,0.1)'
    }}>
      <h2 style={{ marginTop: 0 }}>Update Your Information</h2>
      <TextField label="Full Name" name="fullName" value={loginForm.fullName} onChange={handleLoginInputChange} required />
      <TextField label="Email" value={loginForm.email} readOnly />
      <TextField label="Age" name="age" value={loginForm.age} onChange={handleLoginInputChange} type="number" required />
      <Dropdown
        label="Skillsets"
        placeholder="Select skillsets"
        multiSelect
        options={skillsetOptions}
        selectedKeys={loginForm.skillsets}
        onChange={handleLoginSkillsetChange}
      />
      <DialogFooter>
        <PrimaryButton text="Save" onClick={handleLoginSave} disabled={loading || !loginForm.fullName || !loginForm.age} />
        <PrimaryButton text="Back to Dashboard" onClick={() => setView('dashboard')} />
      </DialogFooter>
    </div>
  </div>
);


const renderTestPage = () => (
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
          Take a Skill Test
        </span>
        <PrimaryButton
          text="Logout"
          onClick={handleLogout}
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
    </div>

    {/* Test Content Card */}
    <div style={{
      backgroundColor: 'white',
      padding: 20,
      borderRadius: 8,
      boxShadow: '0 4px 10px rgba(0,0,0,0.1)'
    }}>
      <h2 style={{ marginTop: 0 }}>Choose a Skill</h2>
      <Label>Select a skill to take the test:</Label>
      <Stack horizontal wrap tokens={{ childrenGap: 10 }}>
        {loginForm.skillsets.map(skillId => {
          const skill = skillsetOptions.find(opt => opt.key === skillId);
          return (
            <PrimaryButton key={skillId} text={skill?.text || ''} onClick={() => {
              setSelectedTestSkill(skillId);
              setShowTestSection(true);
            }} />
          );
        })}
      </Stack>

      {showTestSection && selectedTestSkill !== null && (
        <div style={{ marginTop: '30px' }}>
          <Label>Test for: <strong>{skillsetOptions.find(opt => opt.key === selectedTestSkill)?.text}</strong></Label>
          <p><strong>Q1:</strong> What is the output of <code>console.log(typeof null)</code>?</p>
          <Dropdown
            options={[
              { key: 'a', text: 'object' },
              { key: 'b', text: 'null' },
              { key: 'c', text: 'undefined' },
              { key: 'd', text: 'function' }
            ]}
            placeholder="Select an answer"
          />
          <PrimaryButton text="Submit Test (Not Implemented)" disabled style={{ marginTop: 10 }} />
        </div>
      )}

      <PrimaryButton text="View Previous Results" onClick={() => alert("Feature coming soon...")} style={{ marginTop: 20, marginRight: 20 }} />
      <PrimaryButton text="Back to Dashboard" onClick={() => setView('dashboard')} style={{ marginTop: 10 }} />
    </div>
  </div>
);


return (
  <div style={{ marginTop: 20, fontFamily: 'Segoe UI', backgroundColor: '#f4f6f8', padding: 30, minHeight: '100vh' }}>
    {view === 'login' && renderLogin()}
    {view === 'dashboard' && renderDashboard()}
    {view === 'edit' && renderEditProfile()}
    {view === 'test' && renderTestPage()}
  </div>
);

};

export default SkillsetUsers;