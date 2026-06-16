import { getIntegrations } from '../store/settings';
import { DiscoveryResult, LogLevel } from './graphDiscovery';

/**
 * Zero-setup discovery via the (cloud) agent's device-code flow.
 * The agent runs the OAuth device code + read-only Graph collection server-side;
 * the portal just shows the code and polls for the result. No app registration,
 * nothing installed locally — the agent is hosted in the cloud.
 */

export interface DeviceSession {
  id: string;
  phase: 'starting' | 'awaiting-auth' | 'collecting' | 'done' | 'error';
  userCode: string;
  verificationUri: string;
  message: string;
  tenantName: string;
  tenantId: string;
  username: string;
  log: { t: string; level: string; text: string }[];
  result: DiscoveryResult | null;
  error: string | null;
}

function cfg() {
  const a = getIntegrations().agent;
  if (!a.enabled || !a.url) throw new Error('No cloud agent configured (Settings → Integrations → Migration Agent).');
  return a;
}

export function cloudAgentConfigured(): boolean {
  const a = getIntegrations().agent;
  return a.enabled && !!a.url && !!a.apiKey;
}

async function call<T>(method: string, path: string): Promise<T> {
  const a = cfg();
  const res = await fetch(`${a.url.replace(/\/+$/, '')}${path}`, {
    method,
    headers: { 'content-type': 'application/json', 'x-api-key': a.apiKey },
  });
  const text = await res.text();
  const data = text ? JSON.parse(text) : null;
  if (!res.ok) throw new Error(data?.error ?? `Agent ${res.status}`);
  return data as T;
}

export async function startDeviceDiscovery(): Promise<DeviceSession> {
  return call<DeviceSession>('POST', '/discovery/start');
}

/** Poll the session until done/error, surfacing new log lines. */
export async function trackDeviceDiscovery(id: string, onLog: (text: string, level: LogLevel) => void): Promise<DeviceSession> {
  let seen = 0;
  for (;;) {
    const s = await call<DeviceSession>('GET', `/discovery/status/${id}`);
    for (; seen < s.log.length; seen++) onLog(s.log[seen].text, (s.log[seen].level as LogLevel) || 'info');
    if (s.phase === 'done' || s.phase === 'error') return s;
    await new Promise((r) => setTimeout(r, 1500));
  }
}
