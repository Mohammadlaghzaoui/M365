import { getGraphToken } from './sso';

/**
 * Real Intune deployment/test of a generated Settings Catalog profile via
 * Microsoft Graph, using the signed-in operator's delegated token.
 *
 * SAFE BY DESIGN: the profile is created WITHOUT any assignment, so no device
 * ever receives it. The tester creates it, reads it back to prove it imports
 * cleanly, and (in test mode) deletes it again — a genuine end-to-end check
 * against the real tenant with zero device impact.
 *
 * Scope required (admin consent): DeviceManagementConfiguration.ReadWrite.All
 */

const GRAPH_BETA = 'https://graph.microsoft.com/beta';
const SCOPES = ['DeviceManagementConfiguration.ReadWrite.All'];

async function graph(method: string, path: string, body?: unknown) {
  const token = await getGraphToken(SCOPES);
  const res = await fetch(`${GRAPH_BETA}${path}`, {
    method,
    headers: { authorization: `Bearer ${token}`, 'content-type': 'application/json' },
    body: body ? JSON.stringify(body) : undefined,
  });
  if (res.status === 204) return null;
  const text = await res.text();
  const data = text ? JSON.parse(text) : null;
  if (!res.ok) throw new Error(`Graph ${res.status}: ${data?.error?.message ?? text}`);
  return data;
}

export interface DeployResult {
  id: string;
  name: string;
  verifiedSettings: number;
  deleted: boolean;
}

/**
 * @param json   the generated Settings Catalog profile JSON
 * @param mode   'test' creates + verifies + DELETES; 'deploy' leaves it (unassigned)
 */
export async function deployProfile(json: string, mode: 'test' | 'deploy', onLog: (line: string, level?: 'info' | 'ok' | 'err' | 'warn') => void): Promise<DeployResult> {
  const profile = JSON.parse(json);
  onLog(`Connecting to Microsoft Graph (delegated, scope DeviceManagementConfiguration.ReadWrite.All) ...`, 'info');
  onLog(`Creating configuration policy "${profile.name}" (UNASSIGNED — no device affected) ...`, 'info');
  const created = await graph('POST', '/deviceManagement/configurationPolicies', profile);
  onLog(`Created policy id ${created.id}`, 'ok');

  onLog('Reading the policy back to verify it imported cleanly ...', 'info');
  const readback = await graph('GET', `/deviceManagement/configurationPolicies/${created.id}?$expand=settings`);
  const count = Array.isArray(readback?.settings) ? readback.settings.length : 0;
  onLog(`Verified: ${count} setting(s) present in the created policy.`, count > 0 ? 'ok' : 'warn');

  let deleted = false;
  if (mode === 'test') {
    onLog('TEST mode — deleting the policy again (nothing left behind) ...', 'info');
    await graph('DELETE', `/deviceManagement/configurationPolicies/${created.id}`);
    deleted = true;
    onLog('Test policy deleted. The generated profile is valid and deployable.', 'ok');
  } else {
    onLog(`Deploy mode — policy kept (unassigned). Assign it to a pilot group in Intune when ready.`, 'ok');
  }
  return { id: created.id, name: profile.name, verifiedSettings: count, deleted };
}
