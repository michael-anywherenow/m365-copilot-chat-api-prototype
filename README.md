# Microsoft 365 Copilot Chat Prototype

This repository contains a single-page web application that demonstrates how to call the [Microsoft 365 Copilot Chat API](https://learn.microsoft.com/en-us/microsoft-365-copilot/extensibility/api/ai-services/chat/overview) directly from the browser.

The UI is built with Vite, React, and Fluent UI. Authentication is handled with MSAL for browsers and the app issues chat requests against Microsoft Graph by using the signed-in user's delegated token.

> **Prerequisites**
>
> - An Azure AD application registered as a **single-page application** with redirect URIs that include your development origin (e.g. `http://localhost:5173`). Make sure the "Single-page application" platform is added under **Authentication** in the Azure portal; using the default "Web" platform will trigger `AADSTS9002326` errors when MSAL redeems tokens from the browser.
> - Delegated Microsoft Graph permissions for Copilot chat (for example `Chat.Read`).
> - A Microsoft 365 user in the tenant with an active Copilot license.

## Getting started

1. Install dependencies:

   ```bash
   npm install
   ```

2. Create the environment file:

   ```bash
   cp frontend/.env.example frontend/.env
   ```

3. Edit `frontend/.env` with the Azure AD application values and optional Copilot API overrides.

4. Start the development server:

   ```bash
   npm run dev
   ```

5. Open the printed URL (default is [http://localhost:5173](http://localhost:5173)), sign in with a Copilot-enabled account, and start chatting.

## Environment configuration (`frontend/.env`)

```ini
VITE_AZURE_AD_TENANT_ID=<your-tenant-id>
VITE_AZURE_AD_CLIENT_ID=<your-spa-client-id>
VITE_AZURE_AD_REDIRECT_URI=http://localhost:5173
VITE_AZURE_AD_SCOPES=https://graph.microsoft.com/Chat.Read
VITE_COPILOT_ENDPOINT=https://graph.microsoft.com/v1.0/copilot
VITE_COPILOT_SUBSCRIPTION_KEY=
```

- `VITE_AZURE_AD_SCOPES` should include the delegated Graph permissions you consented to in Azure AD. Separate multiple scopes with commas. If you override the value, MSAL automatically adds the `openid` and `profile` scopes required for sign-in.
- `VITE_COPILOT_ENDPOINT` is optional—omit or leave blank to use the Microsoft Graph default. Supply a tenant-specific endpoint if required. The app automatically targets `/conversations` for the first request and `/conversations/{conversationId}/chat` for follow-up turns.
- `VITE_COPILOT_SUBSCRIPTION_KEY` is optional and only necessary if your tenant requires the `Ocp-Apim-Subscription-Key` header for Copilot requests.

## Useful scripts

| Command | Description |
| --- | --- |
| `npm run dev` | Start the Vite development server. |
| `npm run build` | Build the frontend for production. |

## Production notes

- Consider switching to redirect-based authentication (`loginRedirect`) for better reliability on mobile devices and popup-restricted browsers.
- Review the scopes you request and limit them to the minimum needed for your scenario.
- Store long-lived configuration (e.g. subscription key) in a secure server-side service if you cannot expose it to the browser in production.

## Troubleshooting

- If chat requests fail with 401/403 errors, confirm the signed-in user has a Copilot license and that consent has been granted for the delegated Graph scopes listed in `.env`.
- If sign-in fails with `AADSTS9002326`, verify the Azure AD app registration has a **Single-page application** platform configured and that `http://localhost:5173` is listed as a redirect URI.
- The Copilot endpoint can return additional validation errors if the prompt contains blocked content. Inspect the console/network tab for the detailed response.
