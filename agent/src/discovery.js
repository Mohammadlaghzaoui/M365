/**
 * Read-only Microsoft Graph discovery — agent (Node) implementation.
 * Mirrors the portal collectors but runs server-side (no CORS), used by the
 * device-code zero-setup flow. STRICTLY read-only: only GET requests, only
 * *.Read.All scopes. Failures become validation items (never fatal).
 */

export const DISCOVERY_SCOPES = [
  'User.Read.All', 'Group.Read.All', 'Directory.Read.All', 'Organization.Read.All',
  'Domain.Read.All', 'Reports.Read.All', 'Sites.Read.All', 'Application.Read.All',
  'Policy.Read.All', 'AuditLog.Read.All', 'DeviceManagementConfiguration.Read.All', 'DeviceManagementManagedDevices.Read.All',
];

const V1 = 'https://graph.microsoft.com/v1.0';
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function gget(token, path, tries = 4) {
  const url = path.startsWith('http') ? path : `${V1}${path}`;
  for (let attempt = 0; ; attempt++) {
    const res = await fetch(url, { headers: { authorization: `Bearer ${token}`, ConsistencyLevel: 'eventual' } });
    if (res.status === 429 && attempt < tries) {
      await sleep((Number(res.headers.get('Retry-After')) || 2 ** attempt) * 1000);
      continue;
    }
    if (!res.ok) { const err = new Error(`Graph ${res.status} on ${path}`); err.status = res.status; throw err; }
    return res.json();
  }
}

async function gall(token, path, cap = 5000) {
  const out = [];
  let url = path;
  while (url && out.length < cap) {
    const data = await gget(token, url);
    out.push(...(data.value ?? []));
    url = data['@odata.nextLink'] ?? null;
  }
  return out;
}

async function reportCsv(token, path) {
  const res = await fetch(`${V1}${path}`, { headers: { authorization: `Bearer ${token}` } });
  if (!res.ok) { const e = new Error(`Graph ${res.status} on ${path}`); e.status = res.status; throw e; }
  const text = await res.text();
  const lines = text.trim().split(/\r?\n/);
  if (lines.length < 2) return [];
  const headers = lines[0].split(',').map((h) => h.replace(/^﻿/, '').trim());
  return lines.slice(1).map((l) => {
    const cells = l.split(',');
    const row = {};
    headers.forEach((h, i) => { row[h] = (cells[i] ?? '').trim(); });
    return row;
  });
}

const toGB = (b) => Math.round((Number(b) || 0) / 1073741824 * 100) / 100;
const skuName = (id, skus) => skus.find((s) => s.skuId === id)?.skuPartNumber ?? id;

export async function runReadOnlyDiscovery(token, log = () => {}) {
  const validations = [];
  const workloads = [];
  const collect = async (workload, cmd, fn, note) => {
    log(cmd, 'cmd');
    try {
      const items = await fn();
      log(`-> ${note(items.length)}`, 'ok');
      workloads.push({ workload, status: 'Exported', count: items.length, note: note(items.length) });
      return items;
    } catch (e) {
      const blocked = e.status === 403 || e.status === 401;
      log(`-> ${workload}: ${blocked ? 'blocked' : 'needs validation'} (${e.status || 'error'})`, 'warn');
      validations.push({ workload, object: cmd, status: e.status || 0, note: String(e.message || e) });
      workloads.push({ workload, status: blocked ? 'Blocked' : 'Needs validation', count: 0, note: blocked ? 'Insufficient permission/consent.' : 'Endpoint not available.' });
      return [];
    }
  };

  log('Connect-MgGraph (device code, read-only)', 'cmd');
  log('Authenticated. Collecting read-only tenant data ...', 'info');

  const orgData = await gget(token, '/organization?$select=displayName,id');
  const org = orgData.value?.[0] ?? { displayName: 'Unknown', id: '' };
  log(`Get-MgOrganization -> ${org.displayName} (${org.id})`, 'ok');

  const domainData = await gget(token, '/domains');
  const domains = (domainData.value ?? []).map((d) => ({ id: d.id, isDefault: d.isDefault, isVerified: d.isVerified, supportedServices: (d.supportedServices ?? []).join(', ') }));
  log(`Get-MgDomain -> ${domains.length} domain(s)`, 'ok');

  const skuData = await gget(token, '/subscribedSkus');
  const skus = skuData.value ?? [];
  const licenses = skus.map((s) => ({ skuPartNumber: s.skuPartNumber, enabled: s.prepaidUnits?.enabled ?? 0, consumed: s.consumedUnits ?? 0, available: (s.prepaidUnits?.enabled ?? 0) - (s.consumedUnits ?? 0) }));
  log(`Get-MgSubscribedSku -> ${licenses.length} SKU(s)`, 'ok');

  const USER_SELECT = 'displayName,userPrincipalName,mail,userType,accountEnabled,department,jobTitle,usageLocation,assignedLicenses,createdDateTime';
  // signInActivity needs AuditLog.Read.All — fall back without it if not consented.
  let rawUsers;
  try {
    rawUsers = await gall(token, `/users?$select=${USER_SELECT},signInActivity&$top=999`);
  } catch (e) {
    if (e.status === 403 || e.status === 401) {
      log('signInActivity needs AuditLog.Read.All (not consented) — retrying without last sign-in dates', 'warn');
      validations.push({ workload: 'Users (sign-in activity)', object: 'signInActivity', status: 403, note: 'Last sign-in dates require AuditLog.Read.All consent; all other user attributes were collected.' });
      rawUsers = await gall(token, `/users?$select=${USER_SELECT}&$top=999`);
    } else { throw e; }
  }
  const users = rawUsers.map((u) => ({
    displayName: u.displayName ?? '', userPrincipalName: u.userPrincipalName ?? '', mail: u.mail ?? '',
    userType: u.userType ?? 'Member', accountEnabled: u.accountEnabled !== false, department: u.department ?? '',
    jobTitle: u.jobTitle ?? '', usageLocation: u.usageLocation ?? '',
    licenses: (u.assignedLicenses ?? []).map((l) => skuName(l.skuId, skus)).join(', '),
    createdDateTime: String(u.createdDateTime ?? '').slice(0, 10),
    lastSignIn: String(u.signInActivity?.lastSignInDateTime ?? '').slice(0, 10),
  }));
  const guests = users.filter((u) => u.userType === 'Guest').length;
  const disabled = users.filter((u) => !u.accountEnabled).length;
  log(`Get-MgUser -> ${users.length} users (${guests} guests, ${disabled} disabled)`, 'ok');

  const rawGroups = await gall(token, '/groups?$select=displayName,mail,groupTypes,securityEnabled,mailEnabled,visibility,resourceProvisioningOptions&$top=999');
  const groups = rawGroups.map((g) => {
    const types = g.groupTypes ?? [];
    const isM365 = types.includes('Unified');
    const isTeam = (g.resourceProvisioningOptions ?? []).includes('Team');
    const groupType = isM365 ? 'Microsoft 365' : g.securityEnabled && g.mailEnabled ? 'Mail-enabled security' : g.securityEnabled ? 'Security' : g.mailEnabled ? 'Distribution' : 'Other';
    return { displayName: g.displayName ?? '', mail: g.mail ?? '', groupType, membershipType: types.includes('DynamicMembership') ? 'Dynamic' : 'Assigned', members: 0, visibility: g.visibility ?? '', isTeam };
  });
  const m365Groups = groups.filter((g) => g.groupType === 'Microsoft 365').length;
  const teams = groups.filter((g) => g.isTeam).length;
  const securityGroups = groups.filter((g) => g.groupType.includes('Security')).length;
  const distributionGroups = groups.filter((g) => g.groupType === 'Distribution').length;
  log(`Get-MgGroup -> ${groups.length} groups (${m365Groups} M365, ${teams} Teams)`, 'ok');

  let devices = { total: 0, byOs: {}, compliant: 0, nonCompliant: 0, byCompliance: {} };
  let deviceInventory = [];
  try {
    const dev = await gall(token, '/deviceManagement/managedDevices?$select=deviceName,operatingSystem,osVersion,complianceState,managedDeviceOwnerType,manufacturer,model,userPrincipalName,emailAddress,lastSyncDateTime,enrolledDateTime,serialNumber,managementAgent,isEncrypted&$top=999', 20000);
    const byOs = {}; const byCompliance = {};
    deviceInventory = dev.map((d) => {
      const os = String(d.operatingSystem || 'Unknown'); byOs[os] = (byOs[os] ?? 0) + 1;
      const comp = String(d.complianceState || 'unknown'); byCompliance[comp] = (byCompliance[comp] ?? 0) + 1;
      return {
        deviceName: d.deviceName ?? '', user: d.userPrincipalName ?? d.emailAddress ?? '', os, osVersion: d.osVersion ?? '',
        compliance: comp, ownership: d.managedDeviceOwnerType ?? '', manufacturer: d.manufacturer ?? '', model: d.model ?? '',
        serialNumber: d.serialNumber ?? '', managementAgent: d.managementAgent ?? '', encrypted: d.isEncrypted === true,
        lastSync: String(d.lastSyncDateTime ?? '').slice(0, 10), enrolled: String(d.enrolledDateTime ?? '').slice(0, 10),
      };
    });
    const compliant = byCompliance['compliant'] ?? 0;
    devices = { total: dev.length, byOs, compliant, nonCompliant: dev.length - compliant, byCompliance };
    log(`Get-MgDeviceManagementManagedDevice -> ${dev.length} (${compliant} compliant)`, 'ok');
  } catch (e) {
    validations.push({ workload: 'Intune (devices)', object: 'managedDevices', status: e.status || 0, note: String(e.message) });
  }

  // Usage reports (sizes)
  let usage = { mailboxCount: 0, mailboxTotalGB: 0, mailboxLargest: [], mailboxOver50GB: 0, archiveCount: 0, oneDriveCount: 0, oneDriveTotalGB: 0, oneDriveOver100GB: 0, spoSiteCount: 0, spoTotalGB: 0, available: false };
  try {
    log('Get-MgReportMailboxUsageDetail -Period D30', 'cmd');
    const mbx = await reportCsv(token, "/reports/getMailboxUsageDetail(period='D30')?$format=text/csv");
    const sizeKey = Object.keys(mbx[0] ?? {}).find((k) => /Storage Used/i.test(k)) ?? 'Storage Used (Byte)';
    const upnKey = Object.keys(mbx[0] ?? {}).find((k) => /User Principal Name/i.test(k)) ?? 'User Principal Name';
    const archiveKey = Object.keys(mbx[0] ?? {}).find((k) => /Has Archive/i.test(k));
    const mailboxes = mbx.map((r) => ({ upn: r[upnKey] || '(hidden)', gb: toGB(r[sizeKey]), archive: archiveKey ? /true|yes/i.test(r[archiveKey]) : false }));
    const mailboxTotalGB = Math.round(mailboxes.reduce((a, m) => a + m.gb, 0));
    const od = await reportCsv(token, "/reports/getOneDriveUsageAccountDetail(period='D30')?$format=text/csv");
    const odKey = Object.keys(od[0] ?? {}).find((k) => /Storage Used/i.test(k)) ?? 'Storage Used (Byte)';
    const odSizes = od.map((r) => toGB(r[odKey]));
    const spo = await reportCsv(token, "/reports/getSharePointSiteUsageDetail(period='D30')?$format=text/csv");
    const spoKey = Object.keys(spo[0] ?? {}).find((k) => /Storage Used/i.test(k)) ?? 'Storage Used (Byte)';
    usage = {
      mailboxCount: mailboxes.length, mailboxTotalGB,
      mailboxLargest: [...mailboxes].sort((a, b) => b.gb - a.gb).slice(0, 10).map((m) => ({ upn: m.upn, gb: m.gb })),
      mailboxOver50GB: mailboxes.filter((m) => m.gb > 50).length, archiveCount: mailboxes.filter((m) => m.archive).length,
      oneDriveCount: od.length, oneDriveTotalGB: Math.round(odSizes.reduce((a, b) => a + b, 0)), oneDriveOver100GB: odSizes.filter((g) => g > 100).length,
      spoSiteCount: spo.length, spoTotalGB: Math.round(spo.reduce((a, r) => a + toGB(r[spoKey]), 0)), available: true,
    };
    log(`Usage reports -> ${mailboxTotalGB} GB mailboxes, ${usage.oneDriveTotalGB} GB OneDrive, ${usage.spoTotalGB} GB SharePoint`, 'ok');
  } catch (e) {
    usage.note = `Usage reports unavailable: ${e.message}`;
    log(`-> usage reports unavailable (${e.status || 'error'})`, 'warn');
  }

  const sites = await collect('SharePoint', 'Get-MgSite -Search *', () => gall(token, '/sites?search=*&$top=200', 2000), (n) => `${n} site(s)`);
  const sharePointSites = sites.map((s) => ({ id: s.id ?? '', name: s.displayName ?? s.name ?? '', webUrl: s.webUrl ?? '', created: String(s.createdDateTime ?? '').slice(0, 10), lastModified: String(s.lastModifiedDateTime ?? '').slice(0, 10) }));

  const caRaw = await collect('Identity (CA)', 'Get-MgIdentityConditionalAccessPolicy', () => gall(token, '/identity/conditionalAccess/policies'), (n) => `${n} CA policy(ies)`);
  const caPolicies = caRaw.map((p) => ({ id: p.id ?? '', displayName: p.displayName ?? '', state: p.state ?? '' }));

  const appsRaw = await collect('Apps', 'Get-MgApplication', () => gall(token, '/applications?$select=appId,displayName,signInAudience,createdDateTime&$top=999'), (n) => `${n} app registration(s)`);
  const appRegistrations = appsRaw.map((a) => ({ appId: a.appId ?? '', displayName: a.displayName ?? '', signInAudience: a.signInAudience ?? '', created: String(a.createdDateTime ?? '').slice(0, 10) }));

  const spRaw = await collect('Enterprise apps', 'Get-MgServicePrincipal', () => gall(token, '/servicePrincipals?$select=appId,displayName,servicePrincipalType,accountEnabled&$top=999'), (n) => `${n} service principal(s)`);
  const servicePrincipals = spRaw.map((s) => ({ appId: s.appId ?? '', displayName: s.displayName ?? '', type: s.servicePrincipalType ?? '', enabled: s.accountEnabled !== false }));

  const intuneConfigs = await collect('Intune (config)', 'Get-MgDeviceManagementDeviceConfiguration', () => gall(token, '/deviceManagement/deviceConfigurations'), (n) => `${n} config(s)`);
  const intuneCompliance = await collect('Intune (compliance)', 'Get-MgDeviceManagementDeviceCompliancePolicy', () => gall(token, '/deviceManagement/deviceCompliancePolicies'), (n) => `${n} compliance policy(ies)`);

  workloads.unshift({ workload: 'Identity', status: 'Exported', count: users.length, note: `${users.length} users, ${guests} guests` });
  workloads.push({ workload: 'Teams', status: teams > 0 ? 'Exported' : 'Needs validation', count: teams, note: `${teams} Teams` });
  workloads.push({ workload: 'OneDrive', status: usage.available ? 'Exported' : 'Needs validation', count: usage.oneDriveCount, note: usage.available ? `${usage.oneDriveCount} accounts sized` : 'OneDrive sizes need Reports.Read.All' });
  workloads.push({ workload: 'Exchange', status: 'Needs validation', count: usage.mailboxCount, note: 'Deep Exchange needs Exchange Online PowerShell.' });
  workloads.push({ workload: 'Power Platform', status: 'Needs validation', count: 0, note: 'Not covered by Graph.' });

  return {
    org: { displayName: org.displayName, tenantId: org.id, verifiedDomains: domains.filter((d) => d.isVerified).length },
    users, guests, disabled, groups, m365Groups, teams, securityGroups, distributionGroups,
    licenses, domains, devices, deviceInventory, usage,
    sharePointSites, teamsDetail: [], siteDrives: [], sharing: { anonymous: 0, organization: 0, users: 0, total: 0, sampledDrives: 0 },
    oneDriveSample: { sampled: 0, readable: usage.oneDriveCount, notReadable: 0, usedGB: usage.oneDriveTotalGB },
    caPolicies, appRegistrations, servicePrincipals,
    intune: { configs: intuneConfigs.length, compliance: intuneCompliance.length, devices: devices.total },
    compliancePolicies: intuneCompliance.map((p) => ({ name: p.displayName ?? '', platform: /android/i.test(String(p['@odata.type'] ?? '')) ? 'Android' : /ios/i.test(String(p['@odata.type'] ?? '')) ? 'iOS/iPadOS' : /macOS/i.test(String(p['@odata.type'] ?? '')) ? 'macOS' : /windows/i.test(String(p['@odata.type'] ?? '')) ? 'Windows' : 'Other' })),
    oneDrive: { readable: usage.oneDriveCount, notReadable: 0 },
    validations, workloads,
    fetchedAt: new Date().toISOString(),
  };
}
