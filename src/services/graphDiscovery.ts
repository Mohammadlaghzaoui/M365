import { getGraphToken } from './sso';

/**
 * Microsoft Graph DISCOVERY — strictly READ-ONLY.
 *
 * Every call here uses only *.Read.All delegated scopes and only GET requests.
 * Nothing is written back to the tenant. Used to build a migration analysis
 * and to fill the discovery Excel workbook.
 */

// Read-only scopes only. The consent screen will show these as read permissions.
export const DISCOVERY_SCOPES = [
  'User.Read.All',
  'Group.Read.All',
  'Directory.Read.All',
  'Organization.Read.All',
  'Domain.Read.All',
  'Reports.Read.All',
  'DeviceManagementManagedDevices.Read.All',
  'Sites.Read.All',
];

const V1 = 'https://graph.microsoft.com/v1.0';

async function get<T = unknown>(path: string, scopes = DISCOVERY_SCOPES): Promise<T> {
  const token = await getGraphToken(scopes);
  const res = await fetch(`${V1}${path}`, {
    method: 'GET',
    headers: { authorization: `Bearer ${token}`, ConsistencyLevel: 'eventual' },
  });
  if (!res.ok) throw new Error(`Graph ${res.status} on ${path}: ${(await res.text()).slice(0, 200)}`);
  return res.json() as Promise<T>;
}

/** Page through a collection (read-only), capped to keep it browser-friendly. */
async function getAll<T = Record<string, unknown>>(path: string, cap = 5000): Promise<T[]> {
  const out: T[] = [];
  let url: string | null = path;
  const token = await getGraphToken();
  while (url && out.length < cap) {
    const full: string = url.startsWith('http') ? url : `${V1}${url}`;
    const res = await fetch(full, { headers: { authorization: `Bearer ${token}`, ConsistencyLevel: 'eventual' } });
    if (!res.ok) throw new Error(`Graph ${res.status} on ${url}: ${(await res.text()).slice(0, 200)}`);
    const data = await res.json();
    out.push(...(data.value ?? []));
    url = data['@odata.nextLink'] ?? null;
  }
  return out;
}

export interface DiscoveryUser {
  displayName: string;
  userPrincipalName: string;
  mail: string;
  userType: string;
  accountEnabled: boolean;
  department: string;
  jobTitle: string;
  usageLocation: string;
  licenses: string;
  createdDateTime: string;
  lastSignIn: string;
}

export interface DiscoveryGroup {
  displayName: string;
  mail: string;
  groupType: string;
  membershipType: string;
  members: number;
  visibility: string;
  isTeam: boolean;
}

export interface DiscoveryLicense {
  skuPartNumber: string;
  enabled: number;
  consumed: number;
  available: number;
}

export interface DiscoveryDomain {
  id: string;
  isDefault: boolean;
  isVerified: boolean;
  supportedServices: string;
}

export interface UsageStats {
  mailboxCount: number;
  mailboxTotalGB: number;
  mailboxLargest: { upn: string; gb: number }[];
  mailboxOver50GB: number;
  archiveCount: number;
  oneDriveCount: number;
  oneDriveTotalGB: number;
  oneDriveOver100GB: number;
  spoSiteCount: number;
  spoTotalGB: number;
  available: boolean;
  note?: string;
}

export interface DiscoveryResult {
  org: { displayName: string; tenantId: string; verifiedDomains: number };
  users: DiscoveryUser[];
  guests: number;
  disabled: number;
  groups: DiscoveryGroup[];
  m365Groups: number;
  teams: number;
  securityGroups: number;
  distributionGroups: number;
  licenses: DiscoveryLicense[];
  domains: DiscoveryDomain[];
  devices: { total: number; byOs: Record<string, number> };
  usage: UsageStats;
  fetchedAt: string;
}

/** Fetch a Graph usage report (CSV) read-only and parse it into rows. */
async function getReportCsv(path: string): Promise<Record<string, string>[]> {
  const token = await getGraphToken(['Reports.Read.All']);
  const res = await fetch(`${V1}${path}`, { headers: { authorization: `Bearer ${token}` } });
  if (!res.ok) throw new Error(`Graph ${res.status} on ${path}`);
  const text = await res.text();
  const lines = text.trim().split(/\r?\n/);
  if (lines.length < 2) return [];
  const headers = lines[0].split(',').map((h) => h.replace(/^﻿/, '').trim());
  return lines.slice(1).map((line) => {
    // Simple CSV split (report fields don't contain commas in the columns we use).
    const cells = line.split(',');
    const row: Record<string, string> = {};
    headers.forEach((h, i) => { row[h] = (cells[i] ?? '').trim(); });
    return row;
  });
}

const toGB = (bytes: string | number) => Math.round((Number(bytes) || 0) / 1073741824 * 100) / 100;

export async function runUsage(onProgress: (s: string) => void): Promise<UsageStats> {
  const empty: UsageStats = { mailboxCount: 0, mailboxTotalGB: 0, mailboxLargest: [], mailboxOver50GB: 0, archiveCount: 0, oneDriveCount: 0, oneDriveTotalGB: 0, oneDriveOver100GB: 0, spoSiteCount: 0, spoTotalGB: 0, available: false };
  try {
    onProgress('Reading mailbox usage report (read-only) …');
    const mbx = await getReportCsv("/reports/getMailboxUsageDetail(period='D30')?$format=text/csv");
    const sizeKey = Object.keys(mbx[0] ?? {}).find((k) => /Storage Used/i.test(k)) ?? 'Storage Used (Byte)';
    const upnKey = Object.keys(mbx[0] ?? {}).find((k) => /User Principal Name/i.test(k)) ?? 'User Principal Name';
    const archiveKey = Object.keys(mbx[0] ?? {}).find((k) => /Has Archive/i.test(k));
    const mailboxes = mbx.map((r) => ({ upn: r[upnKey] || '(hidden)', gb: toGB(r[sizeKey]), archive: archiveKey ? /true|yes/i.test(r[archiveKey]) : false }));
    const mailboxTotalGB = Math.round(mailboxes.reduce((a, m) => a + m.gb, 0));
    const largest = [...mailboxes].sort((a, b) => b.gb - a.gb).slice(0, 10).map((m) => ({ upn: m.upn, gb: m.gb }));

    onProgress('Reading OneDrive usage report …');
    const od = await getReportCsv("/reports/getOneDriveUsageAccountDetail(period='D30')?$format=text/csv");
    const odSizeKey = Object.keys(od[0] ?? {}).find((k) => /Storage Used/i.test(k)) ?? 'Storage Used (Byte)';
    const odSizes = od.map((r) => toGB(r[odSizeKey]));
    const oneDriveTotalGB = Math.round(odSizes.reduce((a, b) => a + b, 0));

    onProgress('Reading SharePoint site usage report …');
    const spo = await getReportCsv("/reports/getSharePointSiteUsageDetail(period='D30')?$format=text/csv");
    const spoSizeKey = Object.keys(spo[0] ?? {}).find((k) => /Storage Used/i.test(k)) ?? 'Storage Used (Byte)';
    const spoTotalGB = Math.round(spo.reduce((a, r) => a + toGB(r[spoSizeKey]), 0));

    return {
      mailboxCount: mailboxes.length,
      mailboxTotalGB,
      mailboxLargest: largest,
      mailboxOver50GB: mailboxes.filter((m) => m.gb > 50).length,
      archiveCount: mailboxes.filter((m) => m.archive).length,
      oneDriveCount: od.length,
      oneDriveTotalGB,
      oneDriveOver100GB: odSizes.filter((g) => g > 100).length,
      spoSiteCount: spo.length,
      spoTotalGB,
      available: true,
    };
  } catch (e) {
    return { ...empty, note: `Usage reports unavailable: ${e instanceof Error ? e.message : e}. Needs Reports.Read.All; check tenant concealment setting.` };
  }
}

function skuName(skuId: string, skus: { skuId: string; skuPartNumber: string }[]): string {
  return skus.find((s) => s.skuId === skuId)?.skuPartNumber ?? skuId;
}

export async function runDiscovery(onProgress: (step: string) => void): Promise<DiscoveryResult> {
  onProgress('Reading organization & domains …');
  const orgData = await get<{ value: { displayName: string; id: string }[] }>('/organization?$select=displayName,id');
  const org = orgData.value?.[0] ?? { displayName: 'Unknown', id: '' };
  const domainData = await get<{ value: { id: string; isDefault: boolean; isVerified: boolean; supportedServices: string[] }[] }>('/domains');
  const domains: DiscoveryDomain[] = (domainData.value ?? []).map((d) => ({
    id: d.id, isDefault: d.isDefault, isVerified: d.isVerified, supportedServices: (d.supportedServices ?? []).join(', '),
  }));

  onProgress('Reading subscribed licenses …');
  const skuData = await get<{ value: { skuId: string; skuPartNumber: string; consumedUnits: number; prepaidUnits: { enabled: number } }[] }>('/subscribedSkus');
  const skus = skuData.value ?? [];
  const licenses: DiscoveryLicense[] = skus.map((s) => ({
    skuPartNumber: s.skuPartNumber,
    enabled: s.prepaidUnits?.enabled ?? 0,
    consumed: s.consumedUnits ?? 0,
    available: (s.prepaidUnits?.enabled ?? 0) - (s.consumedUnits ?? 0),
  }));

  onProgress('Reading users & assigned licenses …');
  const rawUsers = await getAll<Record<string, unknown>>('/users?$select=displayName,userPrincipalName,mail,userType,accountEnabled,department,jobTitle,usageLocation,assignedLicenses,createdDateTime,signInActivity&$top=999');
  const users: DiscoveryUser[] = rawUsers.map((u) => ({
    displayName: String(u.displayName ?? ''),
    userPrincipalName: String(u.userPrincipalName ?? ''),
    mail: String(u.mail ?? ''),
    userType: String(u.userType ?? 'Member'),
    accountEnabled: u.accountEnabled !== false,
    department: String(u.department ?? ''),
    jobTitle: String(u.jobTitle ?? ''),
    usageLocation: String(u.usageLocation ?? ''),
    licenses: ((u.assignedLicenses as { skuId: string }[]) ?? []).map((l) => skuName(l.skuId, skus)).join(', '),
    createdDateTime: String(u.createdDateTime ?? '').slice(0, 10),
    lastSignIn: String((u.signInActivity as { lastSignInDateTime?: string })?.lastSignInDateTime ?? '').slice(0, 10),
  }));
  const guests = users.filter((u) => u.userType === 'Guest').length;
  const disabled = users.filter((u) => !u.accountEnabled).length;

  onProgress('Reading groups & teams …');
  const rawGroups = await getAll<Record<string, unknown>>('/groups?$select=displayName,mail,groupTypes,securityEnabled,mailEnabled,visibility,resourceProvisioningOptions&$top=999');
  const groups: DiscoveryGroup[] = rawGroups.map((g) => {
    const types = (g.groupTypes as string[]) ?? [];
    const isM365 = types.includes('Unified');
    const isTeam = ((g.resourceProvisioningOptions as string[]) ?? []).includes('Team');
    const groupType = isM365 ? 'Microsoft 365' : g.securityEnabled && g.mailEnabled ? 'Mail-enabled security' : g.securityEnabled ? 'Security' : g.mailEnabled ? 'Distribution' : 'Other';
    return {
      displayName: String(g.displayName ?? ''),
      mail: String(g.mail ?? ''),
      groupType,
      membershipType: types.includes('DynamicMembership') ? 'Dynamic' : 'Assigned',
      members: 0,
      visibility: String(g.visibility ?? ''),
      isTeam,
    };
  });
  const m365Groups = groups.filter((g) => g.groupType === 'Microsoft 365').length;
  const teams = groups.filter((g) => g.isTeam).length;
  const securityGroups = groups.filter((g) => g.groupType.includes('Security')).length;
  const distributionGroups = groups.filter((g) => g.groupType === 'Distribution').length;

  onProgress('Reading managed devices …');
  let devices = { total: 0, byOs: {} as Record<string, number> };
  try {
    const dev = await getAll<{ operatingSystem?: string }>('/deviceManagement/managedDevices?$select=operatingSystem&$top=999', 10000);
    const byOs: Record<string, number> = {};
    for (const d of dev) { const os = d.operatingSystem || 'Unknown'; byOs[os] = (byOs[os] ?? 0) + 1; }
    devices = { total: dev.length, byOs };
  } catch {
    onProgress('Device read skipped (no Intune scope/licence).');
  }

  const usage = await runUsage(onProgress);

  return {
    org: { displayName: org.displayName, tenantId: org.id, verifiedDomains: domains.filter((d) => d.isVerified).length },
    users, guests, disabled,
    groups, m365Groups, teams, securityGroups, distributionGroups,
    licenses, domains, devices, usage,
    fetchedAt: new Date().toISOString(),
  };
}
