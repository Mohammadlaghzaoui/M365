import 'isomorphic-fetch';
import { ClientSecretCredential } from '@azure/identity';
import { config } from './config.js';

let credential = null;
let cachedToken = null;

export function graphConfigured() {
  return Boolean(config.tenantId && config.clientId && config.clientSecret);
}

async function getToken() {
  if (!graphConfigured()) throw new Error('Graph app-auth not configured (TENANT_ID/CLIENT_ID/CLIENT_SECRET).');
  if (cachedToken && cachedToken.expiresOnTimestamp - 60000 > Date.now()) return cachedToken.token;
  if (!credential) credential = new ClientSecretCredential(config.tenantId, config.clientId, config.clientSecret);
  const t = await credential.getToken('https://graph.microsoft.com/.default');
  cachedToken = t;
  return t.token;
}

/** Thin Graph REST helper (app-only). */
export async function graphRequest(method, path, body) {
  const token = await getToken();
  const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    method,
    headers: {
      authorization: `Bearer ${token}`,
      'content-type': 'application/json',
      ConsistencyLevel: 'eventual',
    },
    body: body ? JSON.stringify(body) : undefined,
  });
  const text = await res.text();
  const data = text ? JSON.parse(text) : null;
  if (!res.ok) throw new Error(`Graph ${res.status}: ${data?.error?.message ?? text}`);
  return data;
}
