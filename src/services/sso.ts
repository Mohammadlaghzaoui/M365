import { PublicClientApplication, AccountInfo } from '@azure/msal-browser';
import { getSSOSettings } from '../store/settings';

/**
 * Microsoft 365 SSO via Entra ID (MSAL, authorization code + PKCE in a popup).
 *
 * One-time setup in the customer's tenant:
 *  1. Entra admin center > App registrations > New registration
 *  2. Supported account types: single tenant
 *  3. Platform: Single-page application (SPA)
 *     Redirect URI: the exact portal URL (e.g. https://<host>/M365/)
 *  4. Copy the Directory (tenant) ID + Application (client) ID into Settings.
 */

let pca: PublicClientApplication | null = null;
let initializedFor = '';

async function getClient(): Promise<PublicClientApplication> {
  const cfg = getSSOSettings();
  if (!cfg.clientId || !cfg.tenantId) throw new Error('SSO is not configured. Set tenant ID and client ID in Settings.');
  const key = `${cfg.tenantId}:${cfg.clientId}`;
  if (!pca || initializedFor !== key) {
    pca = new PublicClientApplication({
      auth: {
        clientId: cfg.clientId,
        authority: `https://login.microsoftonline.com/${cfg.tenantId}`,
        redirectUri: window.location.origin + window.location.pathname,
      },
      cache: { cacheLocation: 'localStorage' },
    });
    await pca.initialize();
    initializedFor = key;
  }
  return pca;
}

export async function signIn(): Promise<AccountInfo> {
  const client = await getClient();
  const result = await client.loginPopup({ scopes: ['User.Read'] });
  client.setActiveAccount(result.account);
  return result.account;
}

export async function signOut(): Promise<void> {
  const client = await getClient();
  const account = client.getActiveAccount() ?? client.getAllAccounts()[0];
  if (account) await client.logoutPopup({ account });
}

export async function currentAccount(): Promise<AccountInfo | null> {
  const cfg = getSSOSettings();
  if (!cfg.enabled || !cfg.clientId || !cfg.tenantId) return null;
  try {
    const client = await getClient();
    return client.getActiveAccount() ?? client.getAllAccounts()[0] ?? null;
  } catch {
    return null;
  }
}

/** Acquire a Graph token silently (popup fallback) — ready for future Graph-powered features. */
export async function getGraphToken(scopes: string[] = ['User.Read']): Promise<string> {
  const client = await getClient();
  const account = client.getActiveAccount() ?? client.getAllAccounts()[0];
  if (!account) throw new Error('Not signed in.');
  try {
    const res = await client.acquireTokenSilent({ scopes, account });
    return res.accessToken;
  } catch {
    const res = await client.acquireTokenPopup({ scopes, account });
    return res.accessToken;
  }
}
