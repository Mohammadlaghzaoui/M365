import { getGraphToken } from './sso';

/**
 * Live tenant insights via Microsoft Graph (delegated, read-only).
 * Real data from the signed-in tenant — replaces any placeholder numbers on
 * the dashboard once the operator is signed in with Microsoft SSO.
 * Scopes: User.Read.All (counts), Directory.Read.All recommended.
 */

const GRAPH = 'https://graph.microsoft.com/v1.0';

async function count(path: string): Promise<number> {
  const token = await getGraphToken(['User.Read.All']);
  const res = await fetch(`${GRAPH}${path}`, {
    headers: { authorization: `Bearer ${token}`, ConsistencyLevel: 'eventual' },
  });
  if (!res.ok) throw new Error(`Graph ${res.status}: ${await res.text()}`);
  const text = await res.text();
  const n = Number(text);
  if (!Number.isNaN(n)) return n;
  try { return Number(JSON.parse(text)['@odata.count'] ?? 0); } catch { return 0; }
}

export interface TenantInsights {
  totalUsers: number;
  guests: number;
  disabled: number;
  fetchedAt: string;
}

export async function getTenantInsights(): Promise<TenantInsights> {
  const [totalUsers, guests, disabled] = await Promise.all([
    count('/users/$count'),
    count(`/users/$count?$filter=${encodeURIComponent("userType eq 'Guest'")}`),
    count(`/users/$count?$filter=${encodeURIComponent('accountEnabled eq false')}`),
  ]);
  return { totalUsers, guests, disabled, fetchedAt: new Date().toISOString() };
}
