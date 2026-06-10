import { getIntegrations } from '../store/settings';

/**
 * Client for the WorkPilot Migration Agent (see /agent in the repo).
 * The agent runs on a trusted host with access to AD / Exchange / Graph and
 * does the real execution; this is the portal's thin HTTPS client to it.
 */

export interface AgentHealth {
  ok: boolean;
  name: string;
  version: string;
  host: string;
  capabilities: { powershell: boolean; graph: boolean; modules: string[] };
  time: string;
}

export interface AgentJob {
  id: string;
  type: string;
  status: 'queued' | 'running' | 'completed' | 'failed';
  progress: number;
  log: { t: string; level: string; text: string }[];
  result: unknown;
  error: string | null;
  startedAt: string;
  finishedAt: string | null;
}

function cfg() {
  const a = getIntegrations().agent;
  if (!a.enabled || !a.url) throw new Error('Migration Agent is not configured (Settings → Integrations → Migration Agent).');
  return a;
}

export function agentConfigured(): boolean {
  const a = getIntegrations().agent;
  return a.enabled && !!a.url && !!a.apiKey;
}

async function call<T>(method: string, path: string, body?: unknown, withKey = true): Promise<T> {
  const a = cfg();
  const headers: Record<string, string> = { 'content-type': 'application/json' };
  if (withKey) headers['x-api-key'] = a.apiKey;
  const res = await fetch(`${a.url.replace(/\/+$/, '')}${path}`, {
    method,
    headers,
    body: body ? JSON.stringify(body) : undefined,
  });
  const text = await res.text();
  const data = text ? JSON.parse(text) : null;
  if (!res.ok) throw new Error(data?.error ?? `Agent ${res.status}`);
  return data as T;
}

export async function agentHealth(): Promise<AgentHealth> {
  return call<AgentHealth>('GET', '/health', undefined, false);
}

export async function verifyEndpoint(type: string): Promise<{ verified: boolean; detail?: string; error?: string }> {
  return call('POST', '/endpoints/verify', { type });
}

export async function startProvision(payload: unknown): Promise<{ jobId: string }> {
  return call('POST', '/provision/user', payload);
}

export async function startMigration(payload: unknown): Promise<{ jobId: string }> {
  return call('POST', '/migrate/start', payload);
}

export async function getJob(id: string): Promise<AgentJob> {
  return call<AgentJob>('GET', `/jobs/${id}`);
}

/** Poll a job until terminal state, invoking onLog for each new log line. */
export async function trackJob(id: string, onLog: (line: { t: string; level: string; text: string }) => void): Promise<AgentJob> {
  let seen = 0;
  for (;;) {
    const job = await getJob(id);
    for (; seen < job.log.length; seen++) onLog(job.log[seen]);
    if (job.status === 'completed' || job.status === 'failed') return job;
    await new Promise((r) => setTimeout(r, 700));
  }
}
