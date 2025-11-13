import { Configuration } from '@azure/msal-browser';

const tenantId = import.meta.env.VITE_AZURE_AD_TENANT_ID as string | undefined;
const clientId = import.meta.env.VITE_AZURE_AD_CLIENT_ID as string | undefined;
const redirectUri = import.meta.env.VITE_AZURE_AD_REDIRECT_URI as string | undefined;

if (!tenantId || !clientId) {
  // eslint-disable-next-line no-console
  console.warn('Missing Azure AD configuration. Sign-in will not work until environment variables are provided.');
}

export const msalConfig: Configuration = {
  auth: {
    clientId: clientId ?? '00000000-0000-0000-0000-000000000000',
    authority: `https://login.microsoftonline.com/${tenantId ?? 'common'}`,
    redirectUri: redirectUri ?? window.location.origin,
  },
  cache: {
    cacheLocation: 'localStorage',
  },
};

const scopesFromEnv = (import.meta.env.VITE_AZURE_AD_SCOPES as string | undefined)
  ?.split(',')
  .map((scope) => scope.trim())
  .filter((scope) => scope.length > 0);

export const loginRequest = {
  scopes:
    scopesFromEnv && scopesFromEnv.length > 0
      ? scopesFromEnv
      : ['https://graph.microsoft.com/Chat.ReadWrite'],
};
