/**
 * Device-code sign-in. The one.com PHP broker is used ONLY to request the device
 * code (Microsoft blocks /devicecode in browsers via CORS). The actual token
 * sign-in is polled DIRECTLY from the browser (Microsoft's /token endpoint allows
 * CORS), so the sign-in is attributed to the engineer's real location — not the
 * one.com server (which is in Denmark). If the direct poll is blocked, it falls
 * back to the broker. Read-only Microsoft Graph runs in the browser with the token.
 */

const BROKER = new URL('api/auth.php', document.baseURI).toString();
const LOCAL_HELPER = 'http://localhost:8799/devicecode'; // optional local helper = sign-in from YOUR PC (Belgium)
const AUTHORITY = 'https://login.microsoftonline.com/organizations';
const GRAPH_CLI_CLIENT = '14d82eec-204b-4c2f-b7e8-296a70dab67e';

export interface DeviceCode {
  device_code: string;
  user_code: string;
  verification_uri: string;
  verification_uri_complete?: string;
  expires_in: number;
  interval: number;
  message: string;
}

let accessToken = '';
let pollViaBroker = false; // flips to true only if the direct browser poll is blocked

/** Is the PHP broker reachable on this host? (cheap probe) */
export async function brokerAvailable(): Promise<boolean> {
  try {
    const res = await fetch(BROKER, { method: 'OPTIONS' });
    return res.ok || res.status === 204;
  } catch {
    return false;
  }
}

/** Step 1: get a device code — PREFER the local helper (your PC = your location), else the server broker. */
export async function requestDeviceCode(): Promise<DeviceCode> {
  // 1) Local helper running on this PC? Then the code is requested from YOUR location (e.g. Belgium).
  try {
    const ctl = new AbortController();
    const t = setTimeout(() => ctl.abort(), 1500);
    const lr = await fetch(LOCAL_HELPER, { method: 'POST', signal: ctl.signal });
    clearTimeout(t);
    if (lr.ok) {
      const ld = await lr.json().catch(() => ({}));
      if (ld.device_code) { pollViaBroker = false; return ld; }
    }
  } catch { /* helper not running — fall back to the server broker */ }

  // 2) Fall back to the one.com server broker (requested server-side; shows the server's location).
  const res = await fetch(`${BROKER}?action=start`, { method: 'POST' });
  const data = await res.json().catch(() => ({}));
  if (!res.ok || !data.device_code) {
    if (res.status === 404) {
      throw new Error('The sign-in helper (api/auth.php) is not on this server yet. Upload public/api/auth.php to one.com, then try again.');
    }
    throw new Error(data.error_description || data.error || `Could not start sign-in (${res.status}).`);
  }
  pollViaBroker = false;
  return data;
}

/** One poll attempt — prefers a DIRECT browser call so the sign-in shows the engineer's location. */
async function pollOnce(deviceCode: string): Promise<Record<string, string>> {
  const body = new URLSearchParams({
    grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
    client_id: GRAPH_CLI_CLIENT,
    device_code: deviceCode,
  }).toString();
  if (!pollViaBroker) {
    try {
      const res = await fetch(`${AUTHORITY}/oauth2/v2.0/token`, {
        method: 'POST', headers: { 'content-type': 'application/x-www-form-urlencoded' }, body,
      });
      return await res.json().catch(() => ({}));
    } catch {
      pollViaBroker = true; // CORS/network blocked the direct call — fall back to the server broker
    }
  }
  const res = await fetch(`${BROKER}?action=poll`, {
    method: 'POST', headers: { 'content-type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ device_code: deviceCode }).toString(),
  });
  return await res.json().catch(() => ({}));
}

/** Step 2: poll until the admin completes sign-in (direct from the browser when possible). */
export async function pollForToken(dc: DeviceCode, onTick?: () => void): Promise<{ accessToken: string; tenantId: string }> {
  const deadline = Date.now() + dc.expires_in * 1000;
  let interval = (dc.interval || 5) * 1000;
  for (;;) {
    if (Date.now() > deadline) throw new Error('The sign-in code expired. Please start again.');
    await new Promise((r) => setTimeout(r, interval));
    onTick?.();
    const data = await pollOnce(dc.device_code);
    if (data.access_token) {
      accessToken = data.access_token;
      const tenantId = decodeTid(data.id_token) || decodeTid(data.access_token) || '';
      return { accessToken: data.access_token, tenantId };
    }
    if (data.error === 'authorization_pending') continue;
    if (data.error === 'slow_down') { interval += 5000; continue; }
    if (data.error === 'authorization_declined') throw new Error('Sign-in was declined by the admin.');
    if (data.error === 'expired_token') throw new Error('The sign-in code expired. Please start again.');
    if (data.error === 'agent_network') throw new Error('The one.com helper could not reach Microsoft. Check that outbound HTTPS/cURL is allowed.');
    throw new Error(data.error_description || data.error || 'Device sign-in failed.');
  }
}

/** Token provider for the read-only discovery (one Graph token, scopes ignored). */
export async function getOnecomToken(): Promise<string> {
  if (!accessToken) throw new Error('Not signed in. Start the code sign-in first.');
  return accessToken;
}

export function clearOnecomToken() { accessToken = ''; }

function decodeTid(jwt?: string): string {
  if (!jwt || jwt.split('.').length < 2) return '';
  try {
    const payload = JSON.parse(atob(jwt.split('.')[1].replace(/-/g, '+').replace(/_/g, '/')));
    return payload.tid || '';
  } catch { return ''; }
}
