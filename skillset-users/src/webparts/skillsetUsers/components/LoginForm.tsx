import * as React from 'react';
import {
  Stack,
  TextField,
  PrimaryButton,
  MessageBar,
  MessageBarType
} from '@fluentui/react';

interface ILoginFormProps {
  loginForm: {
    email: string;
    password: string;
  };
  loginError: string;
  loading: boolean;
  onInputChange: (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => void;
  onSubmit: () => void;
}

const LoginForm: React.FC<ILoginFormProps> = ({
  loginForm,
  loginError,
  loading,
  onInputChange,
  onSubmit
}) => {
  return (
    <Stack tokens={{ childrenGap: 15 }}>
      <TextField
        label="Email"
        name="email"
        value={loginForm.email}
        onChange={onInputChange}
        required
      />
      <TextField
        label="Password"
        name="password"
        type="password"
        canRevealPassword
        value={loginForm.password}
        onChange={onInputChange}
        required
      />
      {loginError && (
        <MessageBar
          messageBarType={MessageBarType.error}
          isMultiline={false}
        >
          {loginError}
        </MessageBar>
      )}
      <PrimaryButton
        text="Login"
        onClick={onSubmit}
        disabled={!loginForm.email || !loginForm.password || loading}
      />
    </Stack>
  );
};

export default LoginForm;
