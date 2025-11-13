import {
  PublicClientApplication,
  EventType,
  AccountInfo,
} from '@azure/msal-browser';
import { createContext, useContext, useEffect, useMemo, useState, ReactNode } from 'react';
import { msalConfig, loginRequest } from '../msalConfig';

interface AuthContextValue {
  account: AccountInfo | null;
  isAuthenticating: boolean;
  login: () => Promise<void>;
  logout: () => Promise<void>;
  acquireToken: () => Promise<string | null>;
}

const AuthContext = createContext<AuthContextValue | undefined>(undefined);

const pca = new PublicClientApplication(msalConfig);

const pcaReady: { promise: Promise<void> | null } = {
  promise: null,
};

const ensurePcaInitialized = () => {
  if (!pcaReady.promise) {
    pcaReady.promise = pca.initialize();
  }
  return pcaReady.promise;
};

export const AuthProvider = ({ children }: { children: ReactNode }) => {
  const [account, setAccount] = useState<AccountInfo | null>(null);
  const [isAuthenticating, setIsAuthenticating] = useState(true);

  useEffect(() => {
    let isMounted = true;
    let callbackId: string | null = null;

    ensurePcaInitialized()
      .then(() => {
        if (!isMounted) {
          return;
        }

        const accounts = pca.getAllAccounts();
        if (accounts.length > 0) {
          setAccount(accounts[0]);
          pca.setActiveAccount(accounts[0]);
        }
        setIsAuthenticating(false);

        callbackId = pca.addEventCallback((event) => {
          if (event.eventType === EventType.LOGIN_SUCCESS && event.payload?.account) {
            const nextAccount = event.payload.account as AccountInfo;
            setAccount(nextAccount);
            pca.setActiveAccount(nextAccount);
          }
          if (event.eventType === EventType.LOGOUT_SUCCESS) {
            setAccount(null);
            pca.setActiveAccount(null);
          }
        });
      })
      .catch((error) => {
        console.error('Failed to initialize MSAL', error);
        setIsAuthenticating(false);
      });

    return () => {
      isMounted = false;
      if (callbackId) {
        pca.removeEventCallback(callbackId);
      }
    };
  }, []);

  const value = useMemo<AuthContextValue>(() => ({
    account,
    isAuthenticating,
    login: async () => {
      await ensurePcaInitialized();
      setIsAuthenticating(true);
      try {
        await pca.loginPopup(loginRequest);
      } finally {
        setIsAuthenticating(false);
      }
    },
    logout: async () => {
      await ensurePcaInitialized();
      const activeAccount = pca.getActiveAccount() ?? account;
      await pca.logoutPopup({ account: activeAccount ?? undefined });
    },
    acquireToken: async () => {
      await ensurePcaInitialized();
      try {
        const result = await pca.acquireTokenSilent({
          account: pca.getActiveAccount() ?? account ?? undefined,
          scopes: loginRequest.scopes,
        });
        return result.accessToken;
      } catch (error) {
        if ((error as Error).name === 'InteractionRequiredAuthError') {
          const result = await pca.acquireTokenPopup(loginRequest);
          return result.accessToken;
        }
        throw error;
      }
    },
  }), [account, isAuthenticating]);

  return <AuthContext.Provider value={value}>{children}</AuthContext.Provider>;
};

export const useAuth = () => {
  const context = useContext(AuthContext);
  if (!context) {
    throw new Error('useAuth must be used within an AuthProvider');
  }
  return context;
};
