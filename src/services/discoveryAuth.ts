import { PublicClientApplication, AccountInfo } from '@azure/msal-browser';
import { load, save } from '../store/useLocalStorage';

/**
 * Discovery authentication — SEPARATE from the portal SSO (Kelso).
 *
 * Uses a MULTI-TENANT app with the `organizations` authority, so the engineer
 * signs in interactively to ANY customer tenant being migrated — no tenant ID
 * is configured anywhere. Each migration = a fresh interactive sign-in (popup)
 * with that customer's admin credentials, consenting to READ-ONLY scopes.
 */

export interface DiscoveryAuthSettings {
  clientId: string; // multi-tenant public-client app registration (SPA)
}

export function getDiscoveryAuth(): DiscoveryAuthSettings {
  return { clientId: '', ...load<Partial<DiscoveryAuthSettings>>('discovery-auth', {}) };
}
export function saveDiscoveryAuth(s: DiscoveryAuthSettings) {
  save('discovery-auth', s);
}
export function discoveryAuthConfigured(): boolean {
  return !!getDiscoveryAuth().clientId;
}

let pca: PublicClientApplication | null = null;
let initFor = '';

async function client(): Promise<PublicClientApplication> {
  const { clientId } = getDiscoveryAuth();
  if (!clientId) throw new Error('Discovery connection is not configured. Add the multi-tenant app Client ID in Settings → Sign-in → Migration Discovery connection.');
  if (!pca || initFor !== clientId) {
    pca = new PublicClientApplication({
      auth: {
        clientId,
        // ANY work/school tenant — the engineer picks the customer tenant at sign-in.
        authority: 'https://login.microsoftonline.com/organizations',
        redirectUri: window.location.origin + window.location.pathname,
      },
      cache: { cacheLocation: 'sessionStorage' }, // per-session; not persisted across customers
    });
    await pca.initialize();
    initFor = clientId;
  }
  return pca;
}

let activeAccount: AccountInfo | null = null;

/** Interactive popup sign-in to the customer tenant being analyzed. */
export async function connectTenant(scopes: string[]): Promise<{ tenantName: string; tenantId: string; username: string }> {
  const c = await client();
  const result = await c.loginPopup({ scopes, prompt: 'select_account' });
  activeAccount = result.account;
  c.setActiveAccount(result.account);
  return {
    tenantName: result.account?.tenantId ?? '',
    tenantId: result.account?.tenantId ?? '',
    username: result.account?.username ?? '',
  };
}

export function connectedTenant(): { username: string; tenantId: string } | null {
  if (!activeAccount) return null;
  return { username: activeAccount.username, tenantId: activeAccount.tenantId };
}

export async function disconnectTenant(): Promise<void> {
  if (!pca || !activeAccount) { activeAccount = null; return; }
  try { await pca.logoutPopup({ account: activeAccount }); } catch { /* ignore */ }
  activeAccount = null;
}

/** Token provider for read-only discovery against the connected customer tenant. */
export async function getDiscoveryToken(scopes: string[]): Promise<string> {
  const c = await client();
  const account = activeAccount ?? c.getActiveAccount() ?? c.getAllAccounts()[0];
  if (!account) throw new Error('No tenant connected. Click "Connect customer tenant" first.');
  try {
    const res = await c.acquireTokenSilent({ scopes, account });
    return res.accessToken;
  } catch {
    const res = await c.acquireTokenPopup({ scopes, account });
    return res.accessToken;
  }
}
