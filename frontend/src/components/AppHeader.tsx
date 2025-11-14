import { Button, Persona, Text } from '@fluentui/react-components';
import { useAuth } from '../context/AuthProvider';

export const AppHeader = () => {
  const { account, isAuthenticating, login, logout } = useAuth();

  return (
    <header className="app-header">
      <Text as="h1" weight="bold">
        Copilot Chat Prototype
      </Text>
      <div className="header-actions">
        {account ? (
          <>
            <Persona
              name={account.name ?? 'Signed in user'}
              secondaryText={account.username}
              presence={{ status: 'available' }}
            />
            <Button appearance="secondary" onClick={() => void logout()}>
              Sign out
            </Button>
          </>
        ) : (
          <Button
            appearance="primary"
            onClick={() => void login()}
            disabled={isAuthenticating}
          >
            Sign in
          </Button>
        )}
      </div>
    </header>
  );
};
