import { getGraphToken } from './sso';

/**
 * Microsoft Graph DISCOVERY — strictly READ-ONLY.
 *
 * Every call here uses only *.Read.All delegated scopes and only GET requests.
 * Nothing is written back to the tenant. Used to build a migration analysis
 * and to fill the discovery Excel workbook.
 *
 * The token provider is pluggable: by default it uses the portal SSO token, but
 * the Discovery page swaps in the per-customer-tenant token (discoveryAuth) so
 * the same portal can analyze many different tenants interactively.
 */

type TokenProvider = (scopes: string[]) => Promise<string>;
let tokenProvider: TokenProvider = getGraphToken;
export function useDiscoveryToken(fn: TokenProvider | null) {
  tokenProvider = fn ?? getGraphToken;
}

// Read-only scopes only. The consent screen will show these as read permissions.
export const DISCOVERY_SCOPES = [
  'User.Read.All',
  'Group.Read.All',
  'Directory.Read.All',
  'Organization.Read.All',
  'Domain.Read.All',
  'Reports.Read.All',
  'Sites.Read.All',
  'Files.Read.All',
  'Team.ReadBasic.All',
  'Channel.ReadBasic.All',
  'TeamMember.Read.All',
  'Application.Read.All',
  'Policy.Read.All',
  'AuditLog.Read.All',
  'DeviceManagementConfiguration.Read.All',
  'DeviceManagementManagedDevices.Read.All',
];

const V1 = 'https://graph.microsoft.com/v1.0';

export type LogLevel = 'cmd' | 'info' | 'ok' | 'warn' | 'err';
export type Logger = (text: string, level?: LogLevel) => void;
const noop: Logger = () => {};

/** GraphError carries the status so callers can turn 403/404 into validation items. */
export class GraphError extends Error {
  constructor(public status: number, message: string) { super(message); }
}

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

/** Fetch with automatic 429 (rate limit) retry honouring Retry-After. */
async function graphFetch(full: string, token: string, tries = 4): Promise<Response> {
  for (let attempt = 0; ; attempt++) {
    const res = await fetch(full, { headers: { authorization: `Bearer ${token}`, ConsistencyLevel: 'eventual' } });
    if (res.status === 429 && attempt < tries) {
      const wait = Number(res.headers.get('Retry-After')) || Math.min(30, 2 ** attempt);
      await sleep(wait * 1000);
      continue;
    }
    return res;
  }
}

interface BatchReq { id: string; url: string; headers?: Record<string, string> }
interface BatchResp { id: string; status: number; body: unknown }

/** Graph $batch — up to 20 GET requests per call, 429-aware. Read-only. */
async function graphBatch(requests: BatchReq[]): Promise<Map<string, BatchResp>> {
  const token = await tokenProvider(DISCOVERY_SCOPES);
  const out = new Map<string, BatchResp>();
  for (let i = 0; i < requests.length; i += 20) {
    const chunk = requests.slice(i, i + 20);
    const body = { requests: chunk.map((r) => ({ id: r.id, method: 'GET', url: r.url, ...(r.headers ? { headers: r.headers } : {}) })) };
    let res: Response;
    for (let attempt = 0; ; attempt++) {
      res = await fetch(`${V1}/$batch`, { method: 'POST', headers: { authorization: `Bearer ${token}`, 'content-type': 'application/json' }, body: JSON.stringify(body) });
      if (res.status === 429 && attempt < 4) { await sleep((Number(res.headers.get('Retry-After')) || 2 ** attempt) * 1000); continue; }
      break;
    }
    if (!res.ok) throw new GraphError(res.status, `Graph $batch ${res.status}`);
    const data = await res.json();
    for (const r of data.responses ?? []) out.set(r.id, { id: r.id, status: r.status, body: r.body });
    // inner 429s inside a batch: honor the largest Retry-After then continue
    const throttled = (data.responses ?? []).some((r: BatchResp) => r.status === 429);
    if (throttled) await sleep(3000);
  }
  return out;
}

async function get<T = unknown>(path: string, scopes = DISCOVERY_SCOPES): Promise<T> {
  const token = await tokenProvider(scopes);
  const res = await graphFetch(path.startsWith('http') ? path : `${V1}${path}`, token);
  if (!res.ok) throw new GraphError(res.status, `Graph ${res.status} on ${path}: ${(await res.text()).slice(0, 200)}`);
  return res.json() as Promise<T>;
}

/** Page through a collection (read-only), capped, with 429 + nextLink handling. */
async function getAll<T = Record<string, unknown>>(path: string, cap = 5000): Promise<T[]> {
  const out: T[] = [];
  let url: string | null = path;
  const token = await tokenProvider(DISCOVERY_SCOPES);
  while (url && out.length < cap) {
    const full: string = url.startsWith('http') ? url : `${V1}${url}`;
    const res = await graphFetch(full, token);
    if (!res.ok) throw new GraphError(res.status, `Graph ${res.status} on ${url}: ${(await res.text()).slice(0, 200)}`);
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
  // M365 Apps usage (who logs in with the desktop app vs mobile only)
  appPlatforms: string;     // e.g. "Windows, Mobile"
  userCategory: string;     // "Office (desktop app)" | "Mobile only" | "Web only" | "No Office activity"
  workerType: string;       // "Office" (desktop/laptop) | "Field" (mobile-only) | "" (unknown)
  lastOfficeActivity: string;
  deviceCount: number;      // managed/registered devices owned by this user
  deviceTypes: string;      // e.g. "Windows, iOS"
  // Migration-critical identity & mailbox detail
  aliases: string;          // all SMTP proxy addresses
  hybrid: string;           // "Synced (AD)" | "Cloud only"
  company: string;
  office: string;
  mobile: string;
  manager: string;          // manager UPN
  mailboxType: string;      // "User" | "Shared (likely)" | "No mailbox"
  mailboxGB: number;        // mailbox size from the usage report
  mailboxItems: number;     // mailbox item count
  oneDriveGB: number;       // OneDrive size from the usage report
  mfa: string;              // "Registered" | "Not registered" | ""
  authMethods: string;      // registered auth methods
}

export interface DiscoveryGroup {
  id: string;
  displayName: string;
  mail: string;
  groupType: string;
  membershipType: string;
  members: number;
  owners: string;
  visibility: string;
  isTeam: boolean;
}

export interface DiscoveryLicense {
  skuPartNumber: string;
  enabled: number;
  consumed: number;
  available: number;
  plans: string;            // enabled service plans
}

export interface DiscoveryDomain {
  id: string;
  authType: string;         // "Managed" | "Federated"
  isDefault: boolean;
  isVerified: boolean;
  supportedServices: string;
}

export interface UsageStats {
  mailboxCount: number;
  mailboxTotalGB: number;
  mailboxUpns: string[];     // lowercased UPNs that have a mailbox (for shared-mailbox detection)
  mailboxSizes: Record<string, number>;  // lowercased UPN -> GB
  mailboxItems: Record<string, number>;   // lowercased UPN -> item count
  oneDriveSizes: Record<string, number>;  // lowercased UPN -> GB
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

export type WorkloadStatus = 'Exported' | 'Partial' | 'Blocked' | 'Needs validation' | 'Not run';

export interface ValidationItem {
  workload: string;
  object: string;
  status: number;
  note: string;
}

export interface WorkloadReadiness {
  workload: string;
  status: WorkloadStatus;
  count: number;
  note: string;
}

export interface DeviceRecord {
  deviceName: string;
  user: string;
  os: string;
  osVersion: string;
  compliance: string;
  ownership: string;
  manufacturer: string;
  model: string;
  serialNumber: string;
  managementAgent: string;
  encrypted: boolean;
  lastSync: string;
  enrolled: string;
  formFactor: string;       // Desktop | Laptop | Phone | Tablet | Server | Other
  typeCode: string;         // naming-convention device-type code (L/D/W/X/E/F/M/P/T/...)
  suggestedName: string;    // skeleton name per the workstation convention
  source: string;           // Intune | Entra
  isVM: boolean;            // virtual machine / Cloud PC
}

/**
 * Workstation naming convention (from the customer's "Workstation Naming" sheet).
 * Worker types: Field=F, Office=O, Temp=T, Kiosk/Shared=K.
 * Device types: Office Laptop=L, Office Desktop=D, Engineering Laptop=W,
 * Engineering Desktop=X, Executive Laptop=E, Executive Desktop=F, Mac=M,
 * Server=S, Network=N, Appliance/IOT=A, Phone=P, Tablet=T.
 * Format: <SITE><n>-<WorkerType><DeviceType>-<Serial> e.g. AHA1-OW-BC349BC34.
 */
function classifyDevice(os: string, model: string): { formFactor: string; typeCode: string } {
  const o = os.toLowerCase(), m = model.toLowerCase();
  if (/ipad|tab\b|tablet|surface pro/.test(m)) return { formFactor: 'Tablet', typeCode: 'T' };
  if (o.includes('ios') || o.includes('ipados')) return /ipad/.test(m) ? { formFactor: 'Tablet', typeCode: 'T' } : { formFactor: 'Phone', typeCode: 'P' };
  if (o.includes('android')) return { formFactor: 'Phone', typeCode: 'P' };
  if (o.includes('mac')) return { formFactor: 'Mac', typeCode: 'M' };
  if (o.includes('windows')) {
    const desktop = /optiplex|tower|desktop|sff|micro|workstation|precision t|mini/.test(m);
    return desktop ? { formFactor: 'Desktop', typeCode: 'D' } : { formFactor: 'Laptop', typeCode: 'L' };
  }
  if (o.includes('server')) return { formFactor: 'Server', typeCode: 'S' };
  return { formFactor: 'Other', typeCode: 'A' };
}

export interface DiscoveryResult {
  org: { displayName: string; tenantId: string; verifiedDomains: number; tenantType?: string };
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
  devices: { total: number; byOs: Record<string, number>; compliant: number; nonCompliant: number; byCompliance: Record<string, number> };
  deviceInventory: DeviceRecord[];
  appUsage: { desktopApp: number; mobileOnly: number; webOnly: number; noActivity: number; classified: number; office: number; field: number };
  usage: UsageStats;
  // Extended assessment workloads
  sharePointSites: { id: string; name: string; webUrl: string; created: string; lastModified: string }[];
  teamsDetail: { name: string; visibility: string; owners: number; members: number; guests: number; channels: number; privateChannels: number }[];
  siteDrives: { siteName: string; driveName: string; driveType: string; usedGB: number; totalGB: number }[];
  sharing: { anonymous: number; organization: number; users: number; total: number; sampledDrives: number };
  oneDriveSample: { sampled: number; readable: number; notReadable: number; usedGB: number };
  caPolicies: { id: string; displayName: string; state: string }[];
  appRegistrations: { appId: string; displayName: string; signInAudience: string; created: string }[];
  servicePrincipals: { appId: string; displayName: string; type: string; enabled: boolean }[];
  intune: { configs: number; compliance: number; devices: number };
  compliancePolicies: { name: string; platform: string }[];
  adminRoles: { role: string; members: string[] }[];
  oneDrive: { readable: number; notReadable: number };
  validations: ValidationItem[];
  workloads: WorkloadReadiness[];
  fetchedAt: string;
}

/** Fetch a Graph usage report (CSV) read-only and parse it into rows. */
async function getReportCsv(path: string): Promise<Record<string, string>[]> {
  const token = await tokenProvider(['Reports.Read.All']);
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

export async function runUsage(log: Logger = noop): Promise<UsageStats> {
  const empty: UsageStats = { mailboxCount: 0, mailboxTotalGB: 0, mailboxUpns: [], mailboxSizes: {}, mailboxItems: {}, oneDriveSizes: {}, mailboxLargest: [], mailboxOver50GB: 0, archiveCount: 0, oneDriveCount: 0, oneDriveTotalGB: 0, oneDriveOver100GB: 0, spoSiteCount: 0, spoTotalGB: 0, available: false };
  try {
    log("Get-MgReportMailboxUsageDetail -Period D30", 'cmd');
    const mbx = await getReportCsv("/reports/getMailboxUsageDetail(period='D30')?$format=text/csv");
    const sizeKey = Object.keys(mbx[0] ?? {}).find((k) => /Storage Used/i.test(k)) ?? 'Storage Used (Byte)';
    const upnKey = Object.keys(mbx[0] ?? {}).find((k) => /User Principal Name/i.test(k)) ?? 'User Principal Name';
    const archiveKey = Object.keys(mbx[0] ?? {}).find((k) => /Has Archive/i.test(k));
    const itemKey = Object.keys(mbx[0] ?? {}).find((k) => /Item Count/i.test(k));
    const mailboxes = mbx.map((r) => ({ upn: r[upnKey] || '(hidden)', gb: toGB(r[sizeKey]), items: itemKey ? Number(r[itemKey]) || 0 : 0, archive: archiveKey ? /true|yes/i.test(r[archiveKey]) : false }));
    const mailboxTotalGB = Math.round(mailboxes.reduce((a, m) => a + m.gb, 0));
    const largest = [...mailboxes].sort((a, b) => b.gb - a.gb).slice(0, 10).map((m) => ({ upn: m.upn, gb: m.gb }));
    log(`→ ${mailboxes.length} mailbox(es), ${mailboxTotalGB.toLocaleString()} GB total (largest ${largest[0]?.gb ?? 0} GB)`, 'ok');

    log("Get-MgReportOneDriveUsageAccountDetail -Period D30", 'cmd');
    const od = await getReportCsv("/reports/getOneDriveUsageAccountDetail(period='D30')?$format=text/csv");
    const odSizeKey = Object.keys(od[0] ?? {}).find((k) => /Storage Used/i.test(k)) ?? 'Storage Used (Byte)';
    const odOwnerKey = Object.keys(od[0] ?? {}).find((k) => /Owner Principal Name/i.test(k));
    const odSizes = od.map((r) => toGB(r[odSizeKey]));
    const oneDriveTotalGB = Math.round(odSizes.reduce((a, b) => a + b, 0));
    const oneDriveSizes: Record<string, number> = {};
    if (odOwnerKey) for (const r of od) { const o = (r[odOwnerKey] || '').toLowerCase(); if (o) oneDriveSizes[o] = toGB(r[odSizeKey]); }
    log(`→ ${od.length} OneDrive account(s), ${oneDriveTotalGB.toLocaleString()} GB total`, 'ok');

    log("Get-MgReportSharePointSiteUsageDetail -Period D30", 'cmd');
    const spo = await getReportCsv("/reports/getSharePointSiteUsageDetail(period='D30')?$format=text/csv");
    const spoSizeKey = Object.keys(spo[0] ?? {}).find((k) => /Storage Used/i.test(k)) ?? 'Storage Used (Byte)';
    const spoTotalGB = Math.round(spo.reduce((a, r) => a + toGB(r[spoSizeKey]), 0));
    log(`→ ${spo.length} SharePoint site(s), ${spoTotalGB.toLocaleString()} GB total`, 'ok');

    return {
      mailboxCount: mailboxes.length,
      mailboxTotalGB,
      mailboxUpns: mailboxes.map((m) => m.upn.toLowerCase()),
      mailboxSizes: Object.fromEntries(mailboxes.map((m) => [m.upn.toLowerCase(), m.gb])),
      mailboxItems: Object.fromEntries(mailboxes.map((m) => [m.upn.toLowerCase(), m.items])),
      oneDriveSizes,
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
    log(`→ usage reports unavailable: ${e instanceof Error ? e.message : e}`, 'warn');
    return { ...empty, note: `Usage reports unavailable: ${e instanceof Error ? e.message : e}. Needs Reports.Read.All; check tenant concealment setting.` };
  }
}

function skuName(skuId: string, skus: { skuId: string; skuPartNumber: string }[]): string {
  return skus.find((s) => s.skuId === skuId)?.skuPartNumber ?? skuId;
}

export async function runDiscovery(log: Logger = noop): Promise<DiscoveryResult> {
  log(`Connect-MgGraph -Scopes ${DISCOVERY_SCOPES.map((s) => `'${s}'`).join(',')}`, 'cmd');
  log('Authenticating with delegated, READ-ONLY scopes (GET only) ...', 'info');

  log('Get-MgOrganization | Select DisplayName,Id', 'cmd');
  const orgData = await get<{ value: { displayName: string; id: string }[] }>('/organization?$select=displayName,id');
  const org = orgData.value?.[0] ?? { displayName: 'Unknown', id: '' };
  log(`→ Tenant "${org.displayName}" (${org.id})`, 'ok');

  log('Get-MgDomain', 'cmd');
  const domainData = await get<{ value: { id: string; authenticationType?: string; isDefault: boolean; isVerified: boolean; supportedServices: string[] }[] }>('/domains');
  const domains: DiscoveryDomain[] = (domainData.value ?? []).map((d) => ({
    id: d.id, authType: String(d.authenticationType ?? 'Managed'), isDefault: d.isDefault, isVerified: d.isVerified, supportedServices: (d.supportedServices ?? []).join(', '),
  }));
  const federated = domains.filter((d) => /federat/i.test(d.authType)).length;
  log(`→ ${domains.length} domain(s): ${domains.map((d) => d.id).slice(0, 5).join(', ')}${domains.length > 5 ? ' …' : ''}${federated ? ` — ${federated} FEDERATED (ADFS)` : ''}`, federated ? 'warn' : 'ok');

  log('Get-MgSubscribedSku', 'cmd');
  const skuData = await get<{ value: { skuId: string; skuPartNumber: string; consumedUnits: number; prepaidUnits: { enabled: number }; servicePlans?: { servicePlanName: string; provisioningStatus: string }[] }[] }>('/subscribedSkus');
  const skus = skuData.value ?? [];
  const licenses: DiscoveryLicense[] = skus.map((s) => ({
    skuPartNumber: s.skuPartNumber,
    enabled: s.prepaidUnits?.enabled ?? 0,
    consumed: s.consumedUnits ?? 0,
    available: (s.prepaidUnits?.enabled ?? 0) - (s.consumedUnits ?? 0),
    plans: (s.servicePlans ?? []).filter((p) => /success/i.test(p.provisioningStatus)).map((p) => p.servicePlanName).slice(0, 30).join(', '),
  }));
  log(`→ ${licenses.length} license SKU(s), ${licenses.reduce((a, l) => a + l.consumed, 0)} seats consumed`, 'ok');

  log('Get-MgUser -All -Property displayName,upn,mail,userType,licenses,signInActivity …', 'cmd');
  const USER_SELECT = 'displayName,userPrincipalName,mail,userType,accountEnabled,department,jobTitle,usageLocation,assignedLicenses,createdDateTime,proxyAddresses,onPremisesSyncEnabled,companyName,officeLocation,mobilePhone';
  const USER_EXPAND = '&$expand=manager($select=userPrincipalName,displayName)';
  // signInActivity needs AuditLog.Read.All. If that consent is missing the whole
  // /users call 403s, so fall back to the same query without it (last sign-in
  // dates are a nice-to-have; the rest of the assessment must still run).
  let signInBlocked = false;
  let rawUsers: Record<string, unknown>[];
  try {
    rawUsers = await getAll<Record<string, unknown>>(`/users?$select=${USER_SELECT},signInActivity${USER_EXPAND}&$top=999`);
  } catch (e) {
    if (e instanceof GraphError && (e.status === 403 || e.status === 401)) {
      signInBlocked = true;
      log('→ signInActivity needs AuditLog.Read.All (not consented) — retrying without last sign-in dates', 'warn');
      rawUsers = await getAll<Record<string, unknown>>(`/users?$select=${USER_SELECT}${USER_EXPAND}&$top=999`);
    } else {
      throw e;
    }
  }
  log(`→ ${rawUsers.length} user object(s) retrieved`, 'ok');
  const smtpAliases = (px: unknown): string => ((px as string[]) ?? []).filter((p) => /^smtp:/i.test(p)).map((p) => p.replace(/^smtp:/i, '')).join(', ');
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
    appPlatforms: '', userCategory: '', workerType: '', lastOfficeActivity: '', deviceCount: 0, deviceTypes: '',
    aliases: smtpAliases(u.proxyAddresses),
    hybrid: u.onPremisesSyncEnabled === true ? 'Synced (AD)' : 'Cloud only',
    company: String(u.companyName ?? ''),
    office: String(u.officeLocation ?? ''),
    mobile: String(u.mobilePhone ?? ''),
    manager: String((u.manager as { userPrincipalName?: string })?.userPrincipalName ?? ''),
    mailboxType: '', mailboxGB: 0, mailboxItems: 0, oneDriveGB: 0, mfa: '', authMethods: '',
  }));
  const guests = users.filter((u) => u.userType === 'Guest').length;
  const disabled = users.filter((u) => !u.accountEnabled).length;
  log(`   ${guests} guest(s), ${disabled} disabled, ${users.length - guests - disabled} active member(s)`, 'info');

  log('Get-MgGroup -All -Property id,displayName,groupTypes,resourceProvisioningOptions …', 'cmd');
  const rawGroups = await getAll<Record<string, unknown>>('/groups?$select=id,displayName,mail,groupTypes,securityEnabled,mailEnabled,visibility,resourceProvisioningOptions&$top=999');
  const groups: DiscoveryGroup[] = rawGroups.map((g) => {
    const types = (g.groupTypes as string[]) ?? [];
    const isM365 = types.includes('Unified');
    const isTeam = ((g.resourceProvisioningOptions as string[]) ?? []).includes('Team');
    const groupType = isM365 ? 'Microsoft 365' : g.securityEnabled && g.mailEnabled ? 'Mail-enabled security' : g.securityEnabled ? 'Security' : g.mailEnabled ? 'Distribution' : 'Other';
    return {
      id: String(g.id ?? ''),
      displayName: String(g.displayName ?? ''),
      mail: String(g.mail ?? ''),
      groupType,
      membershipType: types.includes('DynamicMembership') ? 'Dynamic' : 'Assigned',
      members: 0,
      owners: '',
      visibility: String(g.visibility ?? ''),
      isTeam,
    };
  });
  const m365Groups = groups.filter((g) => g.groupType === 'Microsoft 365').length;
  const teams = groups.filter((g) => g.isTeam).length;
  const securityGroups = groups.filter((g) => g.groupType.includes('Security')).length;
  const distributionGroups = groups.filter((g) => g.groupType === 'Distribution').length;
  log(`→ ${groups.length} group(s): ${m365Groups} M365, ${teams} Teams, ${securityGroups} security, ${distributionGroups} distribution`, 'ok');

  // Owners + member counts per group (batched, capped) — needed to recreate groups.
  try {
    const GROUP_CAP = 200;
    const targets = groups.slice(0, GROUP_CAP).filter((g) => g.id);
    const reqs: BatchReq[] = [];
    targets.forEach((g, i) => {
      reqs.push({ id: `o${i}`, url: `/groups/${g.id}/owners?$select=userPrincipalName&$top=20` });
      reqs.push({ id: `m${i}`, url: `/groups/${g.id}/members/$count`, headers: { ConsistencyLevel: 'eventual' } });
    });
    if (reqs.length) {
      log(`Get-MgGroupOwner / Get-MgGroupMember -Count  (${targets.length} groups)`, 'cmd');
      const res = await graphBatch(reqs);
      targets.forEach((g, i) => {
        const ow = res.get(`o${i}`); const mc = res.get(`m${i}`);
        if (ow?.status === 200) g.owners = (((ow.body as { value?: { userPrincipalName?: string }[] })?.value) ?? []).map((o) => o.userPrincipalName ?? '').filter(Boolean).join(', ');
        if (mc && mc.status === 200) g.members = Number(mc.body) || 0;
      });
      log(`→ owners + member counts resolved for ${targets.length} group(s)`, 'ok');
    }
  } catch (e) {
    log(`→ group owners/members partial: ${e instanceof Error ? e.message : e}`, 'warn');
  }

  // Devices come from TWO sources so nothing is missed:
  //  1. Intune managed devices (rich: serial, compliance, encryption) — needs Intune.
  //  2. Entra registered/joined devices (/devices) — works with Directory.Read.All,
  //     so even tenants without Intune still get a full device inventory.
  let devices = { total: 0, byOs: {} as Record<string, number>, compliant: 0, nonCompliant: 0, byCompliance: {} as Record<string, number> };
  let deviceInventory: DeviceRecord[] = [];
  const deviceErrors: string[] = [];

  log('Get-MgDeviceManagementManagedDevice -All  (Intune)', 'cmd');
  try {
    const dev = await getAll<Record<string, unknown>>('/deviceManagement/managedDevices?$select=deviceName,operatingSystem,osVersion,complianceState,managedDeviceOwnerType,manufacturer,model,userPrincipalName,emailAddress,lastSyncDateTime,enrolledDateTime,serialNumber,managementAgent,isEncrypted&$top=999', 20000);
    for (const d of dev) {
      const os = String(d.operatingSystem || 'Unknown');
      const model = String(d.model ?? '');
      const serial = String(d.serialNumber ?? '');
      const { formFactor, typeCode } = classifyDevice(os, model);
      deviceInventory.push({
        deviceName: String(d.deviceName ?? ''), user: String(d.userPrincipalName ?? d.emailAddress ?? ''),
        os, osVersion: String(d.osVersion ?? ''), compliance: String(d.complianceState || 'unknown'),
        ownership: String(d.managedDeviceOwnerType ?? ''), manufacturer: String(d.manufacturer ?? ''), model,
        serialNumber: serial, managementAgent: String(d.managementAgent ?? ''), encrypted: d.isEncrypted === true,
        lastSync: String(d.lastSyncDateTime ?? '').slice(0, 10), enrolled: String(d.enrolledDateTime ?? '').slice(0, 10),
        formFactor, typeCode, suggestedName: `<SITE>-<W>${typeCode}-${serial || '<SERIAL>'}`, source: 'Intune',
        isVM: /virtual|vmware|hyper-?v|cloud pc|cloudpc|azure|kvm|virtualbox|xen/i.test(`${model} ${String(d.manufacturer ?? '')}`),
      });
    }
    log(`→ ${dev.length} Intune-managed device(s)`, dev.length ? 'ok' : 'warn');
  } catch (e) {
    const status = e instanceof GraphError ? e.status : 0;
    deviceErrors.push(`Intune devices ${status || ''}: ${e instanceof Error ? e.message : e}`.trim());
    log(`→ Intune device read failed (${status || 'error'}) — will try Entra devices`, 'warn');
  }

  log('Get-MgDevice -All  (Entra registered/joined)', 'cmd');
  try {
    const seen = new Set(deviceInventory.map((d) => d.deviceName.toLowerCase()));
    const ed = await getAll<Record<string, unknown>>('/devices?$select=displayName,operatingSystem,operatingSystemVersion,trustType,isCompliant,isManaged,manufacturer,model,approximateLastSignInDateTime,accountEnabled&$expand=registeredOwners($select=userPrincipalName,displayName)&$top=999', 20000);
    let added = 0;
    for (const d of ed) {
      const name = String(d.displayName ?? '');
      if (seen.has(name.toLowerCase())) continue; // already have richer Intune record
      const os = String(d.operatingSystem || 'Unknown');
      const model = String(d.model ?? '');
      const { formFactor, typeCode } = classifyDevice(os, model);
      const owners = (d.registeredOwners as { userPrincipalName?: string }[]) ?? [];
      const comp = d.isCompliant === true ? 'compliant' : d.isCompliant === false ? 'noncompliant' : 'unknown';
      deviceInventory.push({
        deviceName: name, user: String(owners[0]?.userPrincipalName ?? ''),
        os, osVersion: String(d.operatingSystemVersion ?? ''), compliance: comp,
        ownership: String(d.trustType ?? ''), manufacturer: String(d.manufacturer ?? ''), model,
        serialNumber: '', managementAgent: d.isManaged === true ? 'managed' : 'registered', encrypted: false,
        lastSync: String(d.approximateLastSignInDateTime ?? '').slice(0, 10), enrolled: '',
        formFactor, typeCode, suggestedName: `<SITE>-<W>${typeCode}-<SERIAL>`, source: 'Entra',
        isVM: /virtual|vmware|hyper-?v|cloud pc|cloudpc|azure|kvm|virtualbox|xen/i.test(`${model} ${String(d.manufacturer ?? '')}`),
      });
      added++;
    }
    log(`→ ${added} additional Entra device(s)`, 'ok');
  } catch (e) {
    const status = e instanceof GraphError ? e.status : 0;
    deviceErrors.push(`Entra devices ${status || ''}: ${e instanceof Error ? e.message : e}`.trim());
    log(`→ Entra device read failed (${status || 'error'})`, 'warn');
  }

  // Aggregate counts over the merged inventory.
  {
    const byOs: Record<string, number> = {};
    const byCompliance: Record<string, number> = {};
    for (const d of deviceInventory) { byOs[d.os] = (byOs[d.os] ?? 0) + 1; byCompliance[d.compliance] = (byCompliance[d.compliance] ?? 0) + 1; }
    const compliant = byCompliance['compliant'] ?? 0;
    devices = { total: deviceInventory.length, byOs, compliant, nonCompliant: deviceInventory.length - compliant, byCompliance };
    log(`→ ${deviceInventory.length} device(s) total: ${Object.entries(byOs).map(([o, n]) => `${o} ${n}`).join(', ') || 'none'} (${compliant} compliant)`, deviceInventory.length ? 'ok' : 'warn');
  }

  // ---- Who logs in with the Office desktop app vs mobile only ----
  // Microsoft 365 Apps usage report: per-user platform flags (Windows/Mac/Mobile/Web).
  log("Get-MgReportM365AppUserDetail -Period D30", 'cmd');
  let officeClassified = 0;
  try {
    const appRows = await getReportCsv("/reports/getM365AppUserDetail(period='D30')?$format=text/csv");
    const keys = Object.keys(appRows[0] ?? {});
    const upnKey = keys.find((k) => /User Principal Name/i.test(k)) ?? 'User Principal Name';
    const winKey = keys.find((k) => /^Windows$/i.test(k));
    const macKey = keys.find((k) => /^Mac$/i.test(k));
    const mobKey = keys.find((k) => /^Mobile$/i.test(k));
    const webKey = keys.find((k) => /^Web$/i.test(k));
    const actKey = keys.find((k) => /Last Activity Date/i.test(k));
    const used = (v?: string) => !!v && !/^(no|false|0|)$/i.test(v.trim());
    const map = new Map<string, { platforms: string[]; desktop: boolean; mobile: boolean; web: boolean; last: string }>();
    for (const r of appRows) {
      const upn = (r[upnKey] || '').toLowerCase();
      if (!upn) continue;
      const desktop = used(winKey ? r[winKey] : '') || used(macKey ? r[macKey] : '');
      const mobile = used(mobKey ? r[mobKey] : '');
      const web = used(webKey ? r[webKey] : '');
      const platforms: string[] = [];
      if (used(winKey ? r[winKey] : '')) platforms.push('Windows');
      if (used(macKey ? r[macKey] : '')) platforms.push('Mac');
      if (mobile) platforms.push('Mobile');
      if (web) platforms.push('Web');
      map.set(upn, { platforms, desktop, mobile, web, last: actKey ? (r[actKey] || '') : '' });
    }
    // Per-user device types/count from the merged inventory.
    const devByUser = new Map<string, Set<string>>();
    for (const d of deviceInventory) {
      const u = d.user.toLowerCase();
      if (!u) continue;
      if (!devByUser.has(u)) devByUser.set(u, new Set());
      devByUser.get(u)!.add(d.os);
    }
    const isDesktopOs = (s: string) => /windows|mac/i.test(s);
    const isMobileOs = (s: string) => /ios|ipados|android/i.test(s);
    for (const u of users) {
      const key = u.userPrincipalName.toLowerCase();
      const a = map.get(key);
      const dt = devByUser.get(key);
      const hasDesktopDev = dt ? [...dt].some(isDesktopOs) : false;
      const hasMobileDev = dt ? [...dt].some(isMobileOs) : false;
      if (dt) { u.deviceCount = deviceInventory.filter((d) => d.user.toLowerCase() === key).length; u.deviceTypes = [...dt].join(', '); }
      if (a) {
        u.appPlatforms = a.platforms.join(', ');
        u.lastOfficeActivity = String(a.last).slice(0, 10);
        u.userCategory = a.desktop ? 'Office (desktop app)' : a.mobile ? 'Mobile only' : a.web ? 'Web only' : 'No Office activity';
        officeClassified++;
      } else {
        u.userCategory = u.deviceTypes ? 'Has device, no Office activity' : 'No Office activity';
      }
      // Worker type per the naming convention: Office = desktop/laptop, Field = mobile-only.
      u.workerType = (a?.desktop || hasDesktopDev) ? 'Office' : (a?.mobile || hasMobileDev) ? 'Field' : '';
    }
    // Now that worker types are known, fill the worker-type code into each device's suggested name.
    const workerByUser = new Map(users.map((u) => [u.userPrincipalName.toLowerCase(), u.workerType] as const));
    for (const d of deviceInventory) {
      const wt = workerByUser.get(d.user.toLowerCase());
      const wcode = wt === 'Office' ? 'O' : wt === 'Field' ? 'F' : '<W>';
      d.suggestedName = d.suggestedName.replace('<W>', wcode);
    }
    const deskCount = users.filter((u) => u.workerType === 'Office').length;
    const fieldCount = users.filter((u) => u.workerType === 'Field').length;
    log(`→ M365 Apps usage classified ${officeClassified} user(s): ${deskCount} office, ${fieldCount} field`, 'ok');
  } catch (e) {
    log(`→ M365 Apps usage report unavailable: ${e instanceof Error ? e.message : e}`, 'warn');
  }

  const usage = await runUsage(log);

  // Per-user mailbox / OneDrive detail from the usage reports.
  if (usage.mailboxUpns.length || Object.keys(usage.oneDriveSizes).length) {
    const mbxSet = new Set(usage.mailboxUpns);
    for (const u of users) {
      const upn = u.userPrincipalName.toLowerCase();
      const mailLc = u.mail ? u.mail.toLowerCase() : '';
      const has = mbxSet.has(upn) || (mailLc && mbxSet.has(mailLc));
      u.mailboxType = has ? (u.licenses ? 'User' : 'Shared (likely)') : 'No mailbox';
      u.mailboxGB = usage.mailboxSizes[upn] ?? (mailLc ? usage.mailboxSizes[mailLc] : 0) ?? 0;
      u.mailboxItems = usage.mailboxItems[upn] ?? (mailLc ? usage.mailboxItems[mailLc] : 0) ?? 0;
      u.oneDriveGB = usage.oneDriveSizes[upn] ?? (mailLc ? usage.oneDriveSizes[mailLc] : 0) ?? 0;
    }
    log(`→ ${users.filter((u) => u.mailboxType === 'Shared (likely)').length} likely shared mailbox(es)`, 'info');
  }

  // MFA / authentication-method registration per user (migration needs re-registration planning).
  try {
    log('Get-MgReportAuthenticationMethodUserRegistrationDetail', 'cmd');
    const reg = await getAll<Record<string, unknown>>('/reports/authenticationMethods/userRegistrationDetails?$top=999');
    const regMap = new Map(reg.map((x) => [String(x.userPrincipalName ?? '').toLowerCase(), x] as const));
    let mfaReg = 0;
    for (const u of users) {
      const x = regMap.get(u.userPrincipalName.toLowerCase());
      if (!x) continue;
      u.mfa = x.isMfaRegistered === true ? 'Registered' : 'Not registered';
      u.authMethods = ((x.methodsRegistered as string[]) ?? []).join(', ');
      if (x.isMfaRegistered === true) mfaReg++;
    }
    log(`→ MFA registration: ${mfaReg} of ${users.length} users registered`, 'ok');
  } catch (e) {
    log(`→ MFA registration report unavailable: ${e instanceof Error ? e.message : e}`, 'warn');
  }

  // ---- Extended workloads (each resilient: 403/404 → validation item) ----
  const validations: ValidationItem[] = [];
  if (signInBlocked) validations.push({ workload: 'Users (sign-in activity)', object: 'signInActivity', status: 403, note: 'Last sign-in dates require AuditLog.Read.All consent; all other user attributes were collected.' });
  if (deviceInventory.length === 0 && deviceErrors.length) validations.push({ workload: 'Devices', object: 'managedDevices / devices', status: 403, note: `No devices returned. ${deviceErrors.join(' | ')}. Sign in with an admin who can consent to Intune (DeviceManagementManagedDevices.Read.All) or Directory.Read.All.` });
  const workloads: WorkloadReadiness[] = [];
  const collect = async <T>(workload: string, cmd: string, fn: () => Promise<T[]>, okNote: (n: number) => string): Promise<T[]> => {
    log(cmd, 'cmd');
    try {
      const items = await fn();
      log(`→ ${okNote(items.length)}`, 'ok');
      workloads.push({ workload, status: 'Exported', count: items.length, note: okNote(items.length) });
      return items;
    } catch (e) {
      const status = e instanceof GraphError ? e.status : 0;
      const blocked = status === 403 || status === 401;
      log(`→ ${workload}: ${blocked ? 'blocked' : 'needs validation'} (${status || 'error'})`, 'warn');
      validations.push({ workload, object: cmd, status, note: e instanceof Error ? e.message : String(e) });
      workloads.push({ workload, status: blocked ? 'Blocked' : 'Needs validation', count: 0, note: blocked ? 'Insufficient permission/consent for this read scope.' : 'Endpoint not available or out of scope.' });
      return [];
    }
  };

  const sites = await collect('SharePoint', 'Get-MgSite -Search *',
    () => getAll<Record<string, unknown>>('/sites?search=*&$top=200', 2000),
    (n) => `${n} SharePoint site(s)`);
  const sharePointSites = sites.map((s) => ({
    id: String(s.id ?? ''), name: String(s.displayName ?? s.name ?? ''), webUrl: String(s.webUrl ?? ''),
    created: String(s.createdDateTime ?? '').slice(0, 10), lastModified: String(s.lastModifiedDateTime ?? '').slice(0, 10),
  }));

  const caRaw = await collect('Identity (CA)', 'Get-MgIdentityConditionalAccessPolicy',
    () => getAll<Record<string, unknown>>('/identity/conditionalAccess/policies'),
    (n) => `${n} Conditional Access policy(ies)`);
  const caPolicies = caRaw.map((p) => ({ id: String(p.id ?? ''), displayName: String(p.displayName ?? ''), state: String(p.state ?? '') }));

  const appsRaw = await collect('Apps', 'Get-MgApplication',
    () => getAll<Record<string, unknown>>('/applications?$select=appId,displayName,signInAudience,createdDateTime&$top=999'),
    (n) => `${n} app registration(s)`);
  const appRegistrations = appsRaw.map((a) => ({ appId: String(a.appId ?? ''), displayName: String(a.displayName ?? ''), signInAudience: String(a.signInAudience ?? ''), created: String(a.createdDateTime ?? '').slice(0, 10) }));

  const spRaw = await collect('Enterprise apps', 'Get-MgServicePrincipal',
    () => getAll<Record<string, unknown>>('/servicePrincipals?$select=appId,displayName,servicePrincipalType,accountEnabled&$top=999'),
    (n) => `${n} service principal(s)`);
  const servicePrincipals = spRaw.map((s) => ({ appId: String(s.appId ?? ''), displayName: String(s.displayName ?? ''), type: String(s.servicePrincipalType ?? ''), enabled: s.accountEnabled !== false }));

  const rolesRaw = await collect<Record<string, unknown>>('Admin roles', 'Get-MgDirectoryRole -ExpandProperty members',
    () => getAll<Record<string, unknown>>('/directoryRoles?$expand=members($select=userPrincipalName,displayName)'),
    (n) => `${n} active admin role(s)`);
  const adminRoles = rolesRaw
    .map((r) => ({ role: String(r.displayName ?? ''), members: ((r.members as { userPrincipalName?: string }[]) ?? []).map((m) => String(m.userPrincipalName ?? '')).filter(Boolean) }))
    .filter((r) => r.members.length > 0);

  const intuneConfigs = await collect('Intune (config)', 'Get-MgDeviceManagementDeviceConfiguration',
    () => getAll<unknown>('/deviceManagement/deviceConfigurations'), (n) => `${n} device configuration(s)`);
  const intuneCompliance = await collect<Record<string, unknown>>('Intune (compliance)', 'Get-MgDeviceManagementDeviceCompliancePolicy',
    () => getAll<Record<string, unknown>>('/deviceManagement/deviceCompliancePolicies'), (n) => `${n} compliance policy(ies)`);
  const compliancePolicies = intuneCompliance.map((p) => {
    const odataType = String(p['@odata.type'] ?? '');
    const platform = /android/i.test(odataType) ? 'Android' : /ios/i.test(odataType) ? 'iOS/iPadOS' : /macOS/i.test(odataType) ? 'macOS' : /windows/i.test(odataType) ? 'Windows' : 'Other';
    return { name: String(p.displayName ?? ''), platform };
  });

  // ---- Per-object detail collectors via $batch (capped, 429-safe) ----
  const TEAM_CAP = 150, SITE_CAP = 150, OD_SAMPLE = 100;
  const gb = (b: unknown) => Math.round((Number(b) || 0) / 1073741824 * 100) / 100;

  // Teams detail (channels, owners, members, guests) for Teams-enabled groups.
  const teamsDetail: DiscoveryResult['teamsDetail'] = [];
  try {
    const teamGroups = rawGroups.filter((g) => ((g.resourceProvisioningOptions as string[]) ?? []).includes('Team')).slice(0, TEAM_CAP) as Record<string, unknown>[];
    if (teamGroups.length) {
      log(`Get-MgTeamChannel / owners / members for ${teamGroups.length} team(s) (batched) ...`, 'cmd');
      const reqs: BatchReq[] = [];
      teamGroups.forEach((g, i) => {
        reqs.push({ id: `c${i}`, url: `/teams/${g.id}/channels?$select=membershipType` });
        reqs.push({ id: `o${i}`, url: `/groups/${g.id}/owners/$count` });
        reqs.push({ id: `m${i}`, url: `/groups/${g.id}/members?$select=userType&$top=999` });
      });
      const resp = await graphBatch(reqs);
      teamGroups.forEach((g, i) => {
        const ch = resp.get(`c${i}`)?.body as { value?: { membershipType?: string }[] } | undefined;
        const channels = ch?.value ?? [];
        const mem = resp.get(`m${i}`)?.body as { value?: { userType?: string }[] } | undefined;
        const members = mem?.value ?? [];
        const ownersBody = resp.get(`o${i}`)?.body;
        teamsDetail.push({
          name: String(g.displayName ?? ''),
          visibility: String(g.visibility ?? ''),
          owners: typeof ownersBody === 'number' ? ownersBody : Number(ownersBody) || 0,
          members: members.length,
          guests: members.filter((m) => m.userType === 'Guest').length,
          channels: channels.length,
          privateChannels: channels.filter((c) => c.membershipType === 'private' || c.membershipType === 'shared').length,
        });
      });
      log(`→ detailed ${teamsDetail.length} team(s)`, 'ok');
    }
  } catch (e) {
    validations.push({ workload: 'Teams detail', object: 'channels/owners/members', status: e instanceof GraphError ? e.status : 0, note: e instanceof Error ? e.message : String(e) });
  }

  // SharePoint drives + sharing baseline for sampled sites.
  const siteDrives: DiscoveryResult['siteDrives'] = [];
  const sharing = { anonymous: 0, organization: 0, users: 0, total: 0, sampledDrives: 0 };
  try {
    const sampleSites = sharePointSites.slice(0, SITE_CAP);
    if (sampleSites.length) {
      log(`Get-MgSiteDrive for ${sampleSites.length} site(s) (batched) ...`, 'cmd');
      const driveResp = await graphBatch(sampleSites.map((s, i) => ({ id: `d${i}`, url: `/sites/${s.id}/drives?$select=name,driveType,quota,id` })));
      const driveIds: { id: string; siteName: string; driveName: string }[] = [];
      sampleSites.forEach((s, i) => {
        const body = driveResp.get(`d${i}`)?.body as { value?: Record<string, unknown>[] } | undefined;
        for (const dr of body?.value ?? []) {
          const quota = (dr.quota as { used?: number; total?: number }) ?? {};
          siteDrives.push({ siteName: s.name, driveName: String(dr.name ?? ''), driveType: String(dr.driveType ?? ''), usedGB: gb(quota.used), totalGB: gb(quota.total) });
          if (dr.id) driveIds.push({ id: String(dr.id), siteName: s.name, driveName: String(dr.name ?? '') });
        }
      });
      // Sharing baseline: root permissions for a sample of drives.
      const permSample = driveIds.slice(0, 100);
      if (permSample.length) {
        log(`Get-MgDriveRootPermission for ${permSample.length} drive(s) (sharing baseline) ...`, 'cmd');
        const permResp = await graphBatch(permSample.map((d, i) => ({ id: `p${i}`, url: `/drives/${d.id}/root/permissions` })));
        permSample.forEach((_d, i) => {
          const body = permResp.get(`p${i}`)?.body as { value?: { link?: { scope?: string }; grantedToV2?: unknown }[] } | undefined;
          for (const perm of body?.value ?? []) {
            sharing.total++;
            const scope = perm.link?.scope;
            if (scope === 'anonymous') sharing.anonymous++;
            else if (scope === 'organization') sharing.organization++;
            else if (perm.grantedToV2) sharing.users++;
          }
          sharing.sampledDrives++;
        });
        log(`→ sharing: ${sharing.anonymous} anonymous, ${sharing.organization} org, ${sharing.users} user link(s) across ${sharing.sampledDrives} drive(s)`, 'ok');
      }
    }
  } catch (e) {
    validations.push({ workload: 'SharePoint drives/sharing', object: 'drives/permissions', status: e instanceof GraphError ? e.status : 0, note: e instanceof Error ? e.message : String(e) });
  }

  // OneDrive per user (sampled) — 403/404 = not readable, not a hard error.
  const oneDriveSample = { sampled: 0, readable: 0, notReadable: 0, usedGB: 0 };
  try {
    const sample = users.filter((u) => u.accountEnabled && u.userType !== 'Guest').slice(0, OD_SAMPLE);
    const idUsers = (rawUsers as Record<string, unknown>[]).filter((u) => sample.some((s) => s.userPrincipalName === u.userPrincipalName)).slice(0, OD_SAMPLE);
    if (idUsers.length) {
      log(`Get-MgUserDefaultDrive for ${idUsers.length} user(s) sample (OneDrive readability) ...`, 'cmd');
      const odResp = await graphBatch(idUsers.map((u, i) => ({ id: `u${i}`, url: `/users/${u.id}/drive?$select=quota` })));
      idUsers.forEach((_u, i) => {
        const r = odResp.get(`u${i}`);
        oneDriveSample.sampled++;
        if (r && r.status >= 200 && r.status < 300) {
          oneDriveSample.readable++;
          const q = (r.body as { quota?: { used?: number } })?.quota;
          oneDriveSample.usedGB += gb(q?.used);
        } else {
          oneDriveSample.notReadable++;
        }
      });
      oneDriveSample.usedGB = Math.round(oneDriveSample.usedGB);
      log(`→ OneDrive sample: ${oneDriveSample.readable} readable, ${oneDriveSample.notReadable} not readable (of ${oneDriveSample.sampled})`, 'ok');
    }
  } catch (e) {
    validations.push({ workload: 'OneDrive', object: 'user drive', status: e instanceof GraphError ? e.status : 0, note: e instanceof Error ? e.message : String(e) });
  }

  // ---- Workload readiness summary ----
  workloads.unshift({ workload: 'Identity', status: 'Exported', count: users.length, note: `${users.length} users, ${guests} guests` });
  workloads.push({ workload: 'Teams', status: teams > 0 ? 'Exported' : 'Needs validation', count: teams, note: `${teams} Teams` });
  workloads.push({ workload: 'OneDrive', status: usage.available ? 'Exported' : 'Needs validation', count: usage.oneDriveCount, note: usage.available ? `${usage.oneDriveCount} accounts sized` : 'OneDrive sizes need Reports.Read.All' });
  workloads.push({ workload: 'Exchange', status: 'Needs validation', count: usage.mailboxCount, note: 'Deep Exchange (permissions, rules, connectors) requires Exchange Online PowerShell — Graph baseline only.' });
  workloads.push({ workload: 'Power Platform', status: 'Needs validation', count: 0, note: 'Power Platform inventory not covered by Graph — separate connector planned.' });

  log(`Assessment complete: ${validations.length} validation item(s) logged.`, validations.length ? 'warn' : 'ok');

  return {
    org: { displayName: org.displayName, tenantId: org.id, verifiedDomains: domains.filter((d) => d.isVerified).length },
    users, guests, disabled,
    groups, m365Groups, teams, securityGroups, distributionGroups,
    licenses, domains, devices, deviceInventory,
    appUsage: {
      desktopApp: users.filter((u) => u.userCategory === 'Office (desktop app)').length,
      mobileOnly: users.filter((u) => u.userCategory === 'Mobile only').length,
      webOnly: users.filter((u) => u.userCategory === 'Web only').length,
      noActivity: users.filter((u) => u.userCategory.includes('No Office activity')).length,
      classified: users.filter((u) => u.appPlatforms).length,
      office: users.filter((u) => u.workerType === 'Office').length,
      field: users.filter((u) => u.workerType === 'Field').length,
    },
    usage,
    sharePointSites, teamsDetail, siteDrives, sharing, oneDriveSample,
    caPolicies, appRegistrations, servicePrincipals,
    intune: { configs: intuneConfigs.length, compliance: intuneCompliance.length, devices: devices.total },
    compliancePolicies,
    adminRoles,
    oneDrive: { readable: oneDriveSample.readable || usage.oneDriveCount, notReadable: oneDriveSample.notReadable },
    validations, workloads,
    fetchedAt: new Date().toISOString(),
  };
}
