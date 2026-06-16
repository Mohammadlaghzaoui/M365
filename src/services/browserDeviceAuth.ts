/**
 * Browser-only OAuth 2.0 DEVICE CODE flow against Microsoft identity platform.
 *
 * Zero setup, no app registration, no agent: uses Microsoft's public client
 * "Microsoft Graph Command Line Tools" (14d82eec-...) so any customer admin can
 * sign in with a code at https://microsoft.com/devicelogin. Read-only Graph.
 *
 * Endpoints (organizations = any work/school tenant):
 *   POST /organizations/oauth2/v2.0/devicecode   -> user_code + verification_uri
 *   POST /organizations/oauth2/v2.0/token        -> poll until signed in
 */

export const PUBLIC_GRAPH_CLI_CLIENT = '14d82eec-204b-4c2f-b7e8-296a70dab67e';
const AUTHORITY = 'https://login.microsoftonline.com/organizations';

// Read-only scopes for the assessment (delegated). offline_access for refresh.
const SCOPE = [
  'User.Read.All', 'Group.Read.All', 'Directory.Read.All', 'Organization.Read.All',
  'Domain.Read.All', 'Reports.Read.All', 'Sites.Read.All', 'Application.Read.All',
  'Policy.Read.All', 'DeviceManagementConfiguration.Read.All', 'DeviceManagementManagedDevices.Read.All',
  'offline_access', 'openid', 'profile',
].join(' ');

export interface DeviceCode {
  device_code: string;
  user_code: string;
  verification_uri: string;
  expires_in: number;
  interval: number;
  message: string;
}

let accessToken = '';

/** Step 1: request a device code. */
export async function requestDeviceCode(): Promise<DeviceCode> {
  const res = await fetch(`${AUTHORITY}/oauth2/v2.0/devicecode`, {
    method: 'POST',
    headers: { 'content-type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ client_id: PUBLIC_GRAPH_CLI_CLIENT, scope: SCOPE }).toString(),
  });
  if (!res.ok) {
    throw new Error(`Could not start device login (${res.status}). If this is a CORS error, run the analysis through a hosted agent instead.`);
  }
  return res.json();
}

/** Step 2: poll the token endpoint until the admin completes sign-in. */
export async function pollForToken(dc: DeviceCode, onTick?: () => void): Promise<{ accessToken: string; tenantId: string }> {
  const deadline = Date.now() + dc.expires_in * 1000;
  let interval = (dc.interval || 5) * 1000;
  for (;;) {
    if (Date.now() > deadline) throw new Error('The sign-in code expired. Please start again.');
    await new Promise((r) => setTimeout(r, interval));
    onTick?.();
    const res = await fetch(`${AUTHORITY}/oauth2/v2.0/token`, {
      method: 'POST',
      headers: { 'content-type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
        client_id: PUBLIC_GRAPH_CLI_CLIENT,
        device_code: dc.device_code,
      }).toString(),
    });
    const data = await res.json().catch(() => ({}));
    if (res.ok && data.access_token) {
      accessToken = data.access_token;
      const tenantId = decodeTid(data.id_token) || decodeTid(data.access_token) || '';
      return { accessToken: data.access_token, tenantId };
    }
    if (data.error === 'authorization_pending') continue;
    if (data.error === 'slow_down') { interval += 5000; continue; }
    if (data.error === 'authorization_declined') throw new Error('Sign-in was declined by the admin.');
    if (data.error === 'expired_token') throw new Error('The sign-in code expired. Please start again.');
    throw new Error(data.error_description || data.error || 'Device sign-in failed.');
  }
}

/** Token provider for the read-only discovery (ignores requested scopes — one Graph token). */
export async function getBrowserDeviceToken(): Promise<string> {
  if (!accessToken) throw new Error('Not signed in. Start the code sign-in first.');
  return accessToken;
}

export function clearBrowserDeviceToken() { accessToken = ''; }

function decodeTid(jwt?: string): string {
  if (!jwt || jwt.split('.').length < 2) return '';
  try {
    const payload = JSON.parse(atob(jwt.split('.')[1].replace(/-/g, '+').replace(/_/g, '/')));
    return payload.tid || '';
  } catch { return ''; }
}
