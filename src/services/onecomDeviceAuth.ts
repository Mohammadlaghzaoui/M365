/**
 * Device-code sign-in via a tiny PHP broker hosted on one.com.
 *
 * This is the "enter a code, approve in Microsoft" flow (like Codex), with
 * NOTHING installed locally and NO Render/agent. The browser can't call
 * Microsoft's /devicecode endpoint (no CORS), so `public/api/auth.php` — which
 * lives on your own one.com web hosting — does that one server-side step. The
 * read-only Microsoft Graph collection then runs in the browser with the token.
 */

const BROKER = new URL('api/auth.php', document.baseURI).toString();

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

/** Is the PHP broker reachable on this host? (cheap probe) */
export async function brokerAvailable(): Promise<boolean> {
  try {
    const res = await fetch(BROKER, { method: 'OPTIONS' });
    return res.ok || res.status === 204;
  } catch {
    return false;
  }
}

/** Step 1: ask the broker (server-side) for a device code. */
export async function requestDeviceCode(): Promise<DeviceCode> {
  const res = await fetch(`${BROKER}?action=start`, { method: 'POST' });
  const data = await res.json().catch(() => ({}));
  if (!res.ok || !data.device_code) {
    if (res.status === 404) {
      throw new Error('The sign-in helper (api/auth.php) is not on this server yet. Upload public/api/auth.php to one.com, then try again.');
    }
    throw new Error(data.error_description || data.error || `Could not start sign-in (${res.status}).`);
  }
  return data;
}

/** Step 2: poll the broker until the admin completes sign-in. */
export async function pollForToken(dc: DeviceCode, onTick?: () => void): Promise<{ accessToken: string; tenantId: string }> {
  const deadline = Date.now() + dc.expires_in * 1000;
  let interval = (dc.interval || 5) * 1000;
  for (;;) {
    if (Date.now() > deadline) throw new Error('The sign-in code expired. Please start again.');
    await new Promise((r) => setTimeout(r, interval));
    onTick?.();
    const res = await fetch(`${BROKER}?action=poll`, {
      method: 'POST',
      headers: { 'content-type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({ device_code: dc.device_code }).toString(),
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
