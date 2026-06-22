import { DiscoveryResult, DISCOVERY_SCOPES } from './graphDiscovery';
import { Assessment } from './migrationAssessment';
import { Sheet, downloadWorkbook } from './excelExport';
import { recordExportAudit } from './tenantStore';
import { getSession } from './auth';
import { getSiteCode, getMigrationTarget, getTargetDomain, mapUpnToTarget, endpointClassOf, deviceTypeLabel, suggestedName, proposedLicenseOf } from './naming';

/** Build the full multi-sheet discovery workbook (shared by Discovery + Tenant detail). */
export function buildDiscoverySheets(result: DiscoveryResult, analysis: Assessment | null): Sheet[] {
  const tid = result.org.tenantId;
  const site = getSiteCode(tid);
  const target = getMigrationTarget(tid);
  const targetDomain = getTargetDomain(tid);
  const workerByUser = new Map(result.users.map((u) => [u.userPrincipalName.toLowerCase(), u.workerType] as const));
  return [
    {
      name: 'Summary',
      columns: ['Metric', 'Value'],
      rows: [
        ['Tenant', result.org.displayName],
        ['Tenant ID', result.org.tenantId],
        ['Verified domains', result.org.verifiedDomains],
        ['Total users', result.users.length],
        ['Guests', result.guests],
        ['Disabled accounts', result.disabled],
        ['Microsoft 365 groups', result.m365Groups],
        ['Teams', result.teams],
        ['Security groups', result.securityGroups],
        ['Distribution groups', result.distributionGroups],
        ['Managed devices', result.devices.total],
        ['Compliant devices', result.devices.compliant ?? 0],
        ['Non-compliant devices', result.devices.nonCompliant ?? 0],
        ['Compliance policies', (result.compliancePolicies ?? []).length],
        ['Conditional Access policies', result.caPolicies.length],
        ['License SKUs', result.licenses.length],
        ['Licensed seats consumed', result.licenses.reduce((a, l) => a + l.consumed, 0)],
        ['SharePoint sites', result.sharePointSites.length],
        ['Total data (GB)', Math.round(result.usage.mailboxTotalGB + result.usage.oneDriveTotalGB + result.usage.spoTotalGB)],
        ['Office workers (desktop/laptop)', result.appUsage?.office ?? 0],
        ['Field workers (mobile-only)', result.appUsage?.field ?? 0],
        ['Office desktop-app users', result.appUsage?.desktopApp ?? 0],
        ['Mobile-only users', result.appUsage?.mobileOnly ?? 0],
        ['Web-only users', result.appUsage?.webOnly ?? 0],
        ['Users with no Office activity', result.appUsage?.noActivity ?? 0],
        ['Generated', new Date(result.fetchedAt).toLocaleString()],
      ],
    },
    {
      name: 'Users & Licensing',
      columns: ['Display name', 'UPN', 'Type', 'Enabled', 'Department', 'Worker type', 'Mailbox type', 'Mailbox GB', 'Login evidence', 'Current licenses', 'Proposed target license', 'Action', 'Last sign-in'],
      rows: result.users.filter((u) => u.userType !== 'Guest').map((u) => [u.displayName, u.userPrincipalName, u.userType, u.accountEnabled ? 'Yes' : 'No', u.department, u.workerType || '', u.mailboxType || '', (u.mailboxGB ?? 0) || '', u.appPlatforms || 'No evidence', u.licenses || '', proposedLicenseOf(u.workerType || ''), u.workerType ? 'Assign on cutover' : 'Validate manually', u.lastSignIn]),
    },
    {
      name: 'Mailboxes & Identity',
      columns: ['UPN', 'Primary mail', 'Mailbox type', 'Mailbox GB', 'Aliases (SMTP)', 'Identity (AD sync)', 'Manager', 'Company', 'Office', 'Mobile'],
      rows: result.users.filter((u) => u.userType !== 'Guest').map((u) => [u.userPrincipalName, u.mail || '', u.mailboxType || '', (u.mailboxGB ?? 0) || '', u.aliases || '', u.hybrid || '', u.manager || '', u.company || '', u.office || '', u.mobile || '']),
    },
    {
      name: 'Test Migration (Pilot)',
      columns: ['Source UPN', `Target UPN (${targetDomain})`, 'Display name', 'Worker type', 'Proposed license', 'Mailbox GB', 'Aliases (SMTP)', 'Action'],
      rows: result.users.filter((u) => u.accountEnabled && u.userType !== 'Guest').slice(0, 25).map((u) => [u.userPrincipalName, mapUpnToTarget(u.userPrincipalName, targetDomain), u.displayName, u.workerType || 'Validate', proposedLicenseOf(u.workerType || ''), (u.mailboxGB ?? 0) || '', u.aliases || '', u.workerType ? 'Pilot wave 1' : 'Validate then pilot']),
    },
    {
      name: 'Admin roles',
      columns: ['Admin role', 'Member UPN'],
      rows: (result.adminRoles ?? []).flatMap((x) => x.members.map((m) => [x.role, m] as (string | number)[])),
    },
    {
      name: 'Endpoint Plan',
      columns: ['Current device name', 'Endpoint class', 'Operating system', 'Likely user', 'Suggested new name', 'Device type', 'Naming convention', 'Validation / action', 'Source'],
      rows: (result.deviceInventory ?? []).map((d) => {
        const wt = workerByUser.get((d.user || '').toLowerCase()) ?? '';
        const sug = suggestedName(site, d, wt);
        return [d.deviceName, endpointClassOf(d.formFactor) || '', `${d.os}${d.osVersion ? ' ' + d.osVersion : ''}`.trim(), d.user || '', sug, deviceTypeLabel(d.formFactor), sug ? `${site}-[Worker][Device][Last5]` : '', sug ? 'Rename to suggested' : 'Validate manually', d.source ?? ''];
      }),
    },
    {
      name: 'Groups',
      columns: ['Display name', 'Mail', 'Type', 'Membership', 'Visibility', 'Is Team'],
      rows: result.groups.map((g) => [g.displayName, g.mail, g.groupType, g.membershipType, g.visibility, g.isTeam ? 'Yes' : 'No']),
    },
    {
      name: 'Licenses',
      columns: ['SKU', 'Enabled', 'Consumed', 'Available'],
      rows: result.licenses.map((l) => [l.skuPartNumber, l.enabled, l.consumed, l.available]),
    },
    {
      name: 'Domains',
      columns: ['Domain', 'Default', 'Verified', 'Services'],
      rows: result.domains.map((d) => [d.id, d.isDefault ? 'Yes' : 'No', d.isVerified ? 'Yes' : 'No', d.supportedServices]),
    },
    {
      name: 'Devices by OS',
      columns: ['Operating system', 'Count'],
      rows: Object.entries(result.devices.byOs).map(([os, n]) => [os, n]),
    },
    {
      name: 'Device inventory',
      columns: ['Device', 'User', 'OS', 'OS version', 'Form factor', 'Type code', 'Suggested name', 'Compliance', 'Ownership', 'Manufacturer', 'Model', 'Serial', 'Encrypted', 'Source', 'Last sync', 'Enrolled'],
      rows: (result.deviceInventory ?? []).map((d) => [d.deviceName, d.user, d.os, d.osVersion, d.formFactor ?? '', d.typeCode ?? '', d.suggestedName ?? '', d.compliance, d.ownership, d.manufacturer, d.model, d.serialNumber, d.encrypted ? 'Yes' : 'No', d.source ?? '', d.lastSync, d.enrolled]),
    },
    {
      name: 'Naming convention',
      columns: ['Category', 'Name', 'Code'],
      rows: [
        ['Format', `Workstation name = ${site}-[WorkerCode][DeviceCode][Last5]  (e.g. ${site}-OL12345)`, ''],
        ['Site / company code', `${result.org.displayName} (source) → ${target} (target)`, site],
        ['Worker code', 'Office', 'O'], ['Worker code', 'Field', 'F'], ['Worker code', 'Temp', 'T'], ['Worker code', 'Kiosk', 'K'],
        ['Device code', 'Laptop', 'L'], ['Device code', 'Desktop', 'D'], ['Device code', 'Phone', 'P'], ['Device code', 'Tablet', 'T'], ['Device code', 'Mac', 'M'],
      ],
    },
    {
      name: 'Compliance policies',
      columns: ['Policy', 'Platform'],
      rows: (result.compliancePolicies ?? []).map((p) => [p.name, p.platform]),
    },
    {
      name: 'SharePoint sites',
      columns: ['Name', 'URL', 'Created', 'Last modified'],
      rows: result.sharePointSites.map((s) => [s.name, s.webUrl, s.created, s.lastModified]),
    },
    {
      name: 'Conditional Access',
      columns: ['Policy', 'State'],
      rows: result.caPolicies.map((p) => [p.displayName, p.state]),
    },
    {
      name: 'App registrations',
      columns: ['Display name', 'App ID', 'Audience', 'Created'],
      rows: result.appRegistrations.map((a) => [a.displayName, a.appId, a.signInAudience, a.created]),
    },
    {
      name: 'Enterprise apps',
      columns: ['Display name', 'App ID', 'Type', 'Enabled'],
      rows: result.servicePrincipals.map((s) => [s.displayName, s.appId, s.type, s.enabled ? 'Yes' : 'No']),
    },
    {
      name: 'Teams detail',
      columns: ['Team', 'Visibility', 'Owners', 'Members', 'Guests', 'Channels', 'Private/shared channels'],
      rows: result.teamsDetail.map((t) => [t.name, t.visibility, t.owners, t.members, t.guests, t.channels, t.privateChannels]),
    },
    {
      name: 'Site drives',
      columns: ['Site', 'Drive', 'Type', 'Used GB', 'Total GB'],
      rows: result.siteDrives.map((d) => [d.siteName, d.driveName, d.driveType, d.usedGB, d.totalGB]),
    },
    {
      name: 'Workload readiness',
      columns: ['Workload', 'Status', 'Count', 'Note'],
      rows: result.workloads.map((w) => [w.workload, w.status, w.count, w.note]),
    },
    {
      name: 'Validation items',
      columns: ['Workload', 'Status', 'Note'],
      rows: result.validations.map((v) => [v.workload, v.status, v.note]),
    },
    {
      name: 'Data sizing',
      columns: ['Workload', 'Total GB', 'Count', 'Notes'],
      rows: [
        ['Mailboxes', result.usage.mailboxTotalGB, result.usage.mailboxCount, `${result.usage.mailboxOver50GB} over 50GB, ${result.usage.archiveCount} archives`],
        ['OneDrive', result.usage.oneDriveTotalGB, result.usage.oneDriveCount, `${result.usage.oneDriveOver100GB} over 100GB`],
        ['SharePoint', result.usage.spoTotalGB, result.usage.spoSiteCount, ''],
      ],
    },
    {
      name: 'Largest mailboxes',
      columns: ['User', 'Size (GB)'],
      rows: result.usage.mailboxLargest.map((m) => [m.upn, m.gb]),
    },
    ...(analysis ? [{
      name: 'Migration analysis',
      columns: ['Category', 'Detail'],
      rows: [
        ['Complexity', `${analysis.complexityLabel} (${analysis.complexityScore}/100)`],
        ['Total data', `~${analysis.totalDataGB} GB`],
        ['Estimated window', `~${analysis.estimatedDays} working days`],
        ['Recommended approach', analysis.recommendedApproach],
        ...analysis.findings.map((f) => [f.level.toUpperCase() + ' · ' + f.area, f.text] as (string | number)[]),
      ],
    }, {
      name: 'Prep checklist',
      columns: ['Phase', 'Task'],
      rows: analysis.checklist.map((c) => [c.phase, c.task]),
    }] : []),
  ];
}

/** Build + download the workbook and record the export in the audit trail. */
export function exportDiscoveryWorkbook(result: DiscoveryResult, analysis: Assessment | null): void {
  const sheets = buildDiscoverySheets(result, analysis);
  downloadWorkbook(sheets, `migration-discovery-${result.org.displayName.replace(/[^a-z0-9]+/gi, '-')}-${new Date().toISOString().slice(0, 10)}`);
  recordExportAudit({ operator: getSession()?.email ?? 'local', tenantId: result.org.tenantId, tenantName: result.org.displayName, scopes: DISCOVERY_SCOPES.length, objects: result.users.length + result.groups.length + result.sharePointSites.length, format: 'xlsx' });
}
