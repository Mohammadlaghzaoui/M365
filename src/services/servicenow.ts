import { Ticket } from '../types';
import { getServiceNowSettings } from '../store/settings';

/**
 * ServiceNow Table API integration (REST, basic auth).
 *
 * IMPORTANT — browser CORS: ServiceNow blocks cross-origin calls by default.
 * One-time setup on the instance (admin):
 *   System Web Services > REST > CORS Rules > New
 *   - REST API: Table API
 *   - Domain: the portal origin (e.g. https://<your-host>)
 *   - HTTP methods: GET, POST, PATCH
 * For production, prefer routing through a small proxy/MID-server instead of
 * storing instance credentials in the browser.
 */

function cfg() {
  const s = getServiceNowSettings();
  if (!s.enabled || !s.instanceUrl || !s.username) {
    throw new Error('ServiceNow is not configured. Set instance URL and credentials in Settings.');
  }
  return s;
}

function authHeader(user: string, pass: string): string {
  return 'Basic ' + btoa(`${user}:${pass}`);
}

async function snFetch(path: string, init?: RequestInit) {
  const s = cfg();
  const base = s.instanceUrl.replace(/\/+$/, '');
  const res = await fetch(`${base}${path}`, {
    ...init,
    headers: {
      'content-type': 'application/json',
      accept: 'application/json',
      authorization: authHeader(s.username, s.password),
      ...(init?.headers ?? {}),
    },
  });
  if (!res.ok) throw new Error(`ServiceNow ${res.status}: ${await res.text()}`);
  return res.json();
}

export async function testServiceNow(): Promise<string> {
  const data = await snFetch('/api/now/table/incident?sysparm_limit=1&sysparm_fields=number');
  const n = data?.result?.[0]?.number;
  return n ? `Connected — latest incident visible: ${n}` : 'Connected — Table API reachable.';
}

const urgencyMap: Record<Ticket['urgency'], string> = { critical: '1', high: '2', medium: '3', low: '3' };

export interface SNIncidentResult {
  number: string;
  sysId: string;
  link: string;
}

export async function createIncident(t: Ticket, description: string): Promise<SNIncidentResult> {
  const s = cfg();
  const data = await snFetch('/api/now/table/incident', {
    method: 'POST',
    body: JSON.stringify({
      short_description: `[${t.service}] ${t.category} — ${t.userName || t.userEmail} (${t.customer})`,
      description,
      urgency: urgencyMap[t.urgency],
      impact: urgencyMap[t.urgency],
      category: 'software',
      subcategory: 'email',
      contact_type: 'self-service',
      caller_id: t.userEmail || undefined,
    }),
  });
  const r = data.result;
  const base = s.instanceUrl.replace(/\/+$/, '');
  return {
    number: r.number,
    sysId: r.sys_id,
    link: `${base}/nav_to.do?uri=incident.do?sys_id=${r.sys_id}`,
  };
}

export async function addWorkNote(sysId: string, note: string): Promise<void> {
  await snFetch(`/api/now/table/incident/${sysId}`, {
    method: 'PATCH',
    body: JSON.stringify({ work_notes: note }),
  });
}
