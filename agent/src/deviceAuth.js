import { PublicClientApplication } from '@azure/msal-node';
import { runReadOnlyDiscovery, DISCOVERY_SCOPES } from './discovery.js';

/**
 * Zero-setup tenant analysis via OAuth 2.0 DEVICE CODE flow.
 *
 * Uses Microsoft's own public client "Microsoft Graph Command Line Tools"
 * (14d82eec-204b-4c2f-b7e8-296a70dab67e) — so there is NO app registration to
 * create. The engineer connects ANY customer tenant by entering a code at
 * https://microsoft.com/devicelogin and signing in as that tenant's admin.
 * All Graph access is READ-ONLY; the agent runs server-side (no browser CORS).
 */

const GRAPH_CLI_CLIENT = process.env.DISCOVERY_CLIENT_ID || '14d82eec-204b-4c2f-b7e8-296a70dab67e';

const pca = new PublicClientApplication({
  auth: {
    clientId: GRAPH_CLI_CLIENT,
    authority: 'https://login.microsoftonline.com/organizations',
  },
});

// In-memory sessions (ephemeral; tokens never persisted to disk).
export const sessions = new Map();

function newSession() {
  const id = Math.random().toString(36).slice(2) + Date.now().toString(36);
  const s = {
    id, phase: 'starting', userCode: '', verificationUri: '', message: '',
    tenantName: '', tenantId: '', username: '',
    log: [], result: null, error: null, createdAt: Date.now(),
  };
  sessions.set(id, s);
  // Cap memory: drop sessions older than 30 min.
  for (const [k, v] of sessions) if (Date.now() - v.createdAt > 30 * 60_000) sessions.delete(k);
  return s;
}

const addLog = (s, text, level = 'info') => s.log.push({ t: new Date().toISOString(), level, text });

/** Start a device-code session; returns the code immediately, auth + discovery continue in background. */
export function startDeviceDiscovery() {
  const s = newSession();
  let gotCode = false;
  const codeReady = new Promise((resolve) => {
    pca.acquireTokenByDeviceCode({
      scopes: DISCOVERY_SCOPES,
      deviceCodeCallback: (resp) => {
        s.userCode = resp.userCode;
        s.verificationUri = resp.verificationUri;
        s.message = resp.message;
        s.phase = 'awaiting-auth';
        gotCode = true;
        resolve();
      },
    }).then(async (auth) => {
      s.phase = 'collecting';
      s.tenantId = auth.account?.tenantId ?? '';
      s.username = auth.account?.username ?? '';
      addLog(s, `Signed in as ${s.username} (tenant ${s.tenantId}). Starting read-only collection ...`, 'ok');
      try {
        s.result = await runReadOnlyDiscovery(auth.accessToken, (text, level) => addLog(s, text, level));
        s.tenantName = s.result?.org?.displayName ?? '';
        s.phase = 'done';
        addLog(s, 'Read-only assessment complete. No tenant changes were made.', 'ok');
      } catch (e) {
        s.phase = 'error';
        s.error = String(e?.message ?? e);
        addLog(s, `Collection failed: ${s.error}`, 'err');
      }
    }).catch((e) => {
      s.phase = 'error';
      s.error = String(e?.message ?? e);
      if (!gotCode) resolve();
    });
  });
  return { session: s, codeReady };
}

export function publicSession(s) {
  if (!s) return null;
  return {
    id: s.id, phase: s.phase, userCode: s.userCode, verificationUri: s.verificationUri, message: s.message,
    tenantName: s.tenantName, tenantId: s.tenantId, username: s.username,
    log: s.log, result: s.result, error: s.error,
  };
}
