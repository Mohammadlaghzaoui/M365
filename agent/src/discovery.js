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
  'DeviceManagementApps.Read.All', 'SecurityEvents.Read.All', 'CloudPC.Read.All', 'InformationProtectionPolicy.Read.All',
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

// Workstation naming convention: map OS+model to device-type code + form factor.
function classifyDevice(os, model) {
  const o = os.toLowerCase(), m = model.toLowerCase();
  if (/ipad|tab\b|tablet|surface pro/.test(m)) return { formFactor: 'Tablet', typeCode: 'T' };
  if (o.includes('ios') || o.includes('ipados')) return /ipad/.test(m) ? { formFactor: 'Tablet', typeCode: 'T' } : { formFactor: 'Phone', typeCode: 'P' };
  if (o.includes('android')) return { formFactor: 'Phone', typeCode: 'P' };
  if (o.includes('mac')) return { formFactor: 'Mac', typeCode: 'M' };
  if (o.includes('windows')) { const desktop = /optiplex|tower|desktop|sff|micro|workstation|precision t|mini/.test(m); return desktop ? { formFactor: 'Desktop', typeCode: 'D' } : { formFactor: 'Laptop', typeCode: 'L' }; }
  if (o.includes('server')) return { formFactor: 'Server', typeCode: 'S' };
  return { formFactor: 'Other', typeCode: 'A' };
}

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
  const domains = (domainData.value ?? []).map((d) => ({ id: d.id, authType: String(d.authenticationType ?? 'Managed'), isDefault: d.isDefault, isVerified: d.isVerified, supportedServices: (d.supportedServices ?? []).join(', ') }));
  log(`Get-MgDomain -> ${domains.length} domain(s)`, 'ok');

  const skuData = await gget(token, '/subscribedSkus');
  const skus = skuData.value ?? [];
  const licenses = skus.map((s) => ({ skuPartNumber: s.skuPartNumber, enabled: s.prepaidUnits?.enabled ?? 0, consumed: s.consumedUnits ?? 0, available: (s.prepaidUnits?.enabled ?? 0) - (s.consumedUnits ?? 0), plans: (s.servicePlans ?? []).filter((x) => /success/i.test(x.provisioningStatus)).map((x) => x.servicePlanName).slice(0, 30).join(', ') }));
  log(`Get-MgSubscribedSku -> ${licenses.length} SKU(s)`, 'ok');

  const USER_SELECT = 'displayName,userPrincipalName,mail,userType,accountEnabled,department,jobTitle,usageLocation,assignedLicenses,createdDateTime,proxyAddresses,onPremisesSyncEnabled,companyName,officeLocation,mobilePhone';
  const USER_EXPAND = '&$expand=manager($select=userPrincipalName,displayName)';
  // signInActivity needs AuditLog.Read.All — fall back without it if not consented.
  let rawUsers;
  try {
    rawUsers = await gall(token, `/users?$select=${USER_SELECT},signInActivity${USER_EXPAND}&$top=999`);
  } catch (e) {
    if (e.status === 403 || e.status === 401) {
      log('signInActivity needs AuditLog.Read.All (not consented) — retrying without last sign-in dates', 'warn');
      validations.push({ workload: 'Users (sign-in activity)', object: 'signInActivity', status: 403, note: 'Last sign-in dates require AuditLog.Read.All consent; all other user attributes were collected.' });
      rawUsers = await gall(token, `/users?$select=${USER_SELECT}${USER_EXPAND}&$top=999`);
    } else { throw e; }
  }
  const users = rawUsers.map((u) => ({
    displayName: u.displayName ?? '', userPrincipalName: u.userPrincipalName ?? '', mail: u.mail ?? '',
    userType: u.userType ?? 'Member', accountEnabled: u.accountEnabled !== false, department: u.department ?? '',
    jobTitle: u.jobTitle ?? '', usageLocation: u.usageLocation ?? '',
    licenses: (u.assignedLicenses ?? []).map((l) => skuName(l.skuId, skus)).join(', '),
    createdDateTime: String(u.createdDateTime ?? '').slice(0, 10),
    lastSignIn: String(u.signInActivity?.lastSignInDateTime ?? '').slice(0, 10),
    appPlatforms: '', userCategory: '', workerType: '', lastOfficeActivity: '', deviceCount: 0, deviceTypes: '',
    aliases: ((u.proxyAddresses ?? []).filter((p) => /^smtp:/i.test(p)).map((p) => p.replace(/^smtp:/i, '')).join(', ')),
    hybrid: u.onPremisesSyncEnabled === true ? 'Synced (AD)' : 'Cloud only', company: u.companyName ?? '', office: u.officeLocation ?? '', mobile: u.mobilePhone ?? '', manager: u.manager?.userPrincipalName ?? '', mailboxType: '', mailboxGB: 0, mailboxItems: 0, oneDriveGB: 0, mfa: '', authMethods: '',
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
    return { id: g.id ?? '', displayName: g.displayName ?? '', mail: g.mail ?? '', groupType, membershipType: types.includes('DynamicMembership') ? 'Dynamic' : 'Assigned', members: 0, owners: '', visibility: g.visibility ?? '', isTeam };
  });
  const m365Groups = groups.filter((g) => g.groupType === 'Microsoft 365').length;
  const teams = groups.filter((g) => g.isTeam).length;
  const securityGroups = groups.filter((g) => g.groupType.includes('Security')).length;
  const distributionGroups = groups.filter((g) => g.groupType === 'Distribution').length;
  log(`Get-MgGroup -> ${groups.length} groups (${m365Groups} M365, ${teams} Teams)`, 'ok');

  let devices = { total: 0, byOs: {}, compliant: 0, nonCompliant: 0, byCompliance: {} };
  let deviceInventory = [];
  const deviceErrors = [];
  try {
    const dev = await gall(token, '/deviceManagement/managedDevices?$select=deviceName,operatingSystem,osVersion,complianceState,managedDeviceOwnerType,manufacturer,model,userPrincipalName,emailAddress,lastSyncDateTime,enrolledDateTime,serialNumber,managementAgent,isEncrypted&$top=999', 20000);
    for (const d of dev) {
      const os = String(d.operatingSystem || 'Unknown'); const model = String(d.model ?? ''); const serial = String(d.serialNumber ?? '');
      const { formFactor, typeCode } = classifyDevice(os, model);
      deviceInventory.push({
        deviceName: d.deviceName ?? '', user: d.userPrincipalName ?? d.emailAddress ?? '', os, osVersion: d.osVersion ?? '',
        compliance: String(d.complianceState || 'unknown'), ownership: d.managedDeviceOwnerType ?? '', manufacturer: d.manufacturer ?? '', model,
        serialNumber: serial, managementAgent: d.managementAgent ?? '', encrypted: d.isEncrypted === true,
        lastSync: String(d.lastSyncDateTime ?? '').slice(0, 10), enrolled: String(d.enrolledDateTime ?? '').slice(0, 10),
        formFactor, typeCode, suggestedName: `<SITE>-<W>${typeCode}-${serial || '<SERIAL>'}`, source: 'Intune', isVM: /virtual|vmware|hyper-?v|cloud pc|cloudpc|azure|kvm|virtualbox|xen/i.test(`${model} ${String(d.manufacturer ?? '')}`),
      });
    }
    log(`Get-MgDeviceManagementManagedDevice -> ${dev.length}`, dev.length ? 'ok' : 'warn');
  } catch (e) { deviceErrors.push(`Intune ${e.status || ''}: ${e.message}`); }
  try {
    const seen = new Set(deviceInventory.map((d) => d.deviceName.toLowerCase()));
    const ed = await gall(token, '/devices?$select=displayName,operatingSystem,operatingSystemVersion,trustType,isCompliant,isManaged,manufacturer,model,approximateLastSignInDateTime&$expand=registeredOwners($select=userPrincipalName)&$top=999', 20000);
    let added = 0;
    for (const d of ed) {
      const name = String(d.displayName ?? ''); if (seen.has(name.toLowerCase())) continue;
      const os = String(d.operatingSystem || 'Unknown'); const model = String(d.model ?? '');
      const { formFactor, typeCode } = classifyDevice(os, model);
      const owners = d.registeredOwners ?? [];
      deviceInventory.push({
        deviceName: name, user: owners[0]?.userPrincipalName ?? '', os, osVersion: String(d.operatingSystemVersion ?? ''),
        compliance: d.isCompliant === true ? 'compliant' : d.isCompliant === false ? 'noncompliant' : 'unknown', ownership: d.trustType ?? '',
        manufacturer: d.manufacturer ?? '', model, serialNumber: '', managementAgent: d.isManaged ? 'managed' : 'registered', encrypted: false,
        lastSync: String(d.approximateLastSignInDateTime ?? '').slice(0, 10), enrolled: '',
        formFactor, typeCode, suggestedName: `<SITE>-<W>${typeCode}-<SERIAL>`, source: 'Entra', isVM: /virtual|vmware|hyper-?v|cloud pc|cloudpc|azure|kvm|virtualbox|xen/i.test(`${model} ${String(d.manufacturer ?? '')}`),
      });
      added++;
    }
    log(`Get-MgDevice (Entra) -> +${added}`, 'ok');
  } catch (e) { deviceErrors.push(`Entra ${e.status || ''}: ${e.message}`); }
  {
    const byOs = {}; const byCompliance = {};
    for (const d of deviceInventory) { byOs[d.os] = (byOs[d.os] ?? 0) + 1; byCompliance[d.compliance] = (byCompliance[d.compliance] ?? 0) + 1; }
    const compliant = byCompliance['compliant'] ?? 0;
    devices = { total: deviceInventory.length, byOs, compliant, nonCompliant: deviceInventory.length - compliant, byCompliance };
    if (deviceInventory.length === 0 && deviceErrors.length) validations.push({ workload: 'Devices', object: 'managedDevices / devices', status: 403, note: `No devices returned. ${deviceErrors.join(' | ')}` });
    log(`Devices total -> ${deviceInventory.length} (${compliant} compliant)`, deviceInventory.length ? 'ok' : 'warn');
  }

  // Who logs in with the Office desktop app vs mobile only (M365 Apps usage report).
  try {
    log('Get-MgReportM365AppUserDetail -Period D30', 'cmd');
    const appRows = await reportCsv(token, "/reports/getM365AppUserDetail(period='D30')?$format=text/csv");
    const keys = Object.keys(appRows[0] ?? {});
    const upnK = keys.find((k) => /User Principal Name/i.test(k)) ?? 'User Principal Name';
    const winK = keys.find((k) => /^Windows$/i.test(k)), macK = keys.find((k) => /^Mac$/i.test(k)), mobK = keys.find((k) => /^Mobile$/i.test(k)), webK = keys.find((k) => /^Web$/i.test(k)), actK = keys.find((k) => /Last Activity Date/i.test(k));
    const usedV = (v) => !!v && !/^(no|false|0|)$/i.test(String(v).trim());
    const map = new Map();
    for (const r of appRows) { const u = (r[upnK] || '').toLowerCase(); if (!u) continue;
      const plat = []; if (usedV(winK && r[winK])) plat.push('Windows'); if (usedV(macK && r[macK])) plat.push('Mac'); if (usedV(mobK && r[mobK])) plat.push('Mobile'); if (usedV(webK && r[webK])) plat.push('Web');
      map.set(u, { plat, desktop: usedV(winK && r[winK]) || usedV(macK && r[macK]), mobile: usedV(mobK && r[mobK]), web: usedV(webK && r[webK]), last: actK ? (r[actK] || '') : '' }); }
    const devByUser = new Map();
    for (const d of deviceInventory) { const u = (d.user || '').toLowerCase(); if (!u) continue; if (!devByUser.has(u)) devByUser.set(u, new Set()); devByUser.get(u).add(d.os); }
    const isDesk = (s) => /windows|mac/i.test(s), isMob = (s) => /ios|ipados|android/i.test(s);
    for (const u of users) { const k = u.userPrincipalName.toLowerCase(); const a = map.get(k); const dt = devByUser.get(k);
      const hasDesk = dt ? [...dt].some(isDesk) : false, hasMob = dt ? [...dt].some(isMob) : false;
      if (dt) { u.deviceCount = deviceInventory.filter((d) => (d.user || '').toLowerCase() === k).length; u.deviceTypes = [...dt].join(', '); }
      if (a) { u.appPlatforms = a.plat.join(', '); u.lastOfficeActivity = String(a.last).slice(0, 10); u.userCategory = a.desktop ? 'Office (desktop app)' : a.mobile ? 'Mobile only' : a.web ? 'Web only' : 'No Office activity'; }
      else { u.userCategory = u.deviceTypes ? 'Has device, no Office activity' : 'No Office activity'; }
      u.workerType = (a?.desktop || hasDesk) ? 'Office' : (a?.mobile || hasMob) ? 'Field' : ''; }
    const workerByUser = new Map(users.map((u) => [u.userPrincipalName.toLowerCase(), u.workerType]));
    for (const d of deviceInventory) { const wt = workerByUser.get((d.user || '').toLowerCase()); d.suggestedName = d.suggestedName.replace('<W>', wt === 'Office' ? 'O' : wt === 'Field' ? 'F' : '<W>'); }
    log(`Get-MgReportM365AppUserDetail -> ${users.filter((u) => u.workerType === 'Office').length} office, ${users.filter((u) => u.workerType === 'Field').length} field`, 'ok');
  } catch (e) { log(`M365 Apps usage report unavailable: ${e.message}`, 'warn'); }

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
  let secureScore = null;
  try { const ss = await gget(token, '/security/secureScores?$top=1'); const x = ss.value?.[0]; if (x) secureScore = { current: Math.round(x.currentScore), max: Math.round(x.maxScore), percent: x.maxScore ? Math.round((x.currentScore / x.maxScore) * 100) : 0 }; } catch (e) { log(`Secure Score unavailable: ${e.message}`, 'warn'); }
  const cloudPCsRaw = await collect('Cloud PCs', 'Get-MgVirtualEndpointCloudPC', () => gall(token, '/deviceManagement/virtualEndpoint/cloudPCs'), (n) => `${n} Cloud PC(s)`);
  const appProtRaw = await collect('App protection', 'Get-MgManagedAppPolicy', () => gall(token, '/deviceAppManagement/managedAppPolicies'), (n) => `${n} app protection policy(ies)`);
  const namedLocRaw = await collect('Named locations', 'Get-MgNamedLocation', () => gall(token, '/identity/conditionalAccess/namedLocations'), (n) => `${n} named location(s)`);
  const labelsRaw = await collect('Sensitivity labels', 'Get-MgSensitivityLabel', () => gall(token, '/security/informationProtection/sensitivityLabels'), (n) => `${n} label(s)`);

  workloads.unshift({ workload: 'Identity', status: 'Exported', count: users.length, note: `${users.length} users, ${guests} guests` });
  workloads.push({ workload: 'Teams', status: teams > 0 ? 'Exported' : 'Needs validation', count: teams, note: `${teams} Teams` });
  workloads.push({ workload: 'OneDrive', status: usage.available ? 'Exported' : 'Needs validation', count: usage.oneDriveCount, note: usage.available ? `${usage.oneDriveCount} accounts sized` : 'OneDrive sizes need Reports.Read.All' });
  workloads.push({ workload: 'Exchange', status: 'Needs validation', count: usage.mailboxCount, note: 'Deep Exchange needs Exchange Online PowerShell.' });
  workloads.push({ workload: 'Power Platform', status: 'Needs validation', count: 0, note: 'Not covered by Graph.' });

  return {
    org: { displayName: org.displayName, tenantId: org.id, verifiedDomains: domains.filter((d) => d.isVerified).length },
    users, guests, disabled, groups, m365Groups, teams, securityGroups, distributionGroups,
    licenses, domains, devices, deviceInventory,
    appUsage: {
      desktopApp: users.filter((u) => u.userCategory === 'Office (desktop app)').length,
      mobileOnly: users.filter((u) => u.userCategory === 'Mobile only').length,
      webOnly: users.filter((u) => u.userCategory === 'Web only').length,
      noActivity: users.filter((u) => (u.userCategory || '').includes('No Office activity')).length,
      classified: users.filter((u) => u.appPlatforms).length,
      office: users.filter((u) => u.workerType === 'Office').length,
      field: users.filter((u) => u.workerType === 'Field').length,
    },
    usage,
    sharePointSites, teamsDetail: [], siteDrives: [], sharing: { anonymous: 0, organization: 0, users: 0, total: 0, sampledDrives: 0 },
    oneDriveSample: { sampled: 0, readable: usage.oneDriveCount, notReadable: 0, usedGB: usage.oneDriveTotalGB },
    caPolicies, appRegistrations, servicePrincipals,
    intune: { configs: intuneConfigs.length, compliance: intuneCompliance.length, devices: devices.total },
    compliancePolicies: intuneCompliance.map((p) => ({ name: p.displayName ?? '', platform: /android/i.test(String(p['@odata.type'] ?? '')) ? 'Android' : /ios/i.test(String(p['@odata.type'] ?? '')) ? 'iOS/iPadOS' : /macOS/i.test(String(p['@odata.type'] ?? '')) ? 'macOS' : /windows/i.test(String(p['@odata.type'] ?? '')) ? 'Windows' : 'Other' })),
    adminRoles: [],
    secureScore, cloudPCs: cloudPCsRaw.length, appProtection: appProtRaw.length, namedLocations: namedLocRaw.length, sensitivityLabels: labelsRaw.length,
    oneDrive: { readable: usage.oneDriveCount, notReadable: 0 },
    validations, workloads,
    fetchedAt: new Date().toISOString(),
  };
}
