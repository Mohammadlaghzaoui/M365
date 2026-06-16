import { DiscoveryResult } from './graphDiscovery';

/**
 * Migration assessment engine — turns read-only discovery data into a full
 * migration analysis: complexity score, data sizing & duration estimate,
 * risks, blockers, and a preparation checklist. Pure functions, no I/O.
 */

export interface Finding {
  level: 'info' | 'warn' | 'risk' | 'blocker';
  area: string;
  text: string;
}

export interface ChecklistItem {
  phase: string;
  task: string;
  done?: boolean;
}

export interface Assessment {
  complexityScore: number;       // 0-100
  complexityLabel: 'Low' | 'Medium' | 'High' | 'Very High';
  totalDataGB: number;
  estimatedDays: number;
  recommendedApproach: string;
  findings: Finding[];
  checklist: ChecklistItem[];
  sizing: { label: string; value: string }[];
}

const labelFor = (s: number): Assessment['complexityLabel'] =>
  s >= 75 ? 'Very High' : s >= 50 ? 'High' : s >= 25 ? 'Medium' : 'Low';

export function assess(d: DiscoveryResult): Assessment {
  const u = d.usage;
  const activeUsers = d.users.length - d.guests - d.disabled;
  const totalDataGB = (u.mailboxTotalGB || 0) + (u.oneDriveTotalGB || 0) + (u.spoTotalGB || 0);

  const findings: Finding[] = [];
  const add = (level: Finding['level'], area: string, text: string) => findings.push({ level, area, text });

  // ---- Complexity scoring ----
  let score = 0;
  score += Math.min(25, Math.round(d.users.length / 40));        // user volume
  score += Math.min(15, Math.round(totalDataGB / 500));          // data volume
  score += Math.min(10, d.teams);                                // Teams collaboration
  score += d.domains.length > 3 ? 8 : d.domains.length > 1 ? 4 : 0;
  score += d.guests > 50 ? 8 : d.guests > 0 ? 3 : 0;
  score += d.distributionGroups > 50 ? 6 : 0;
  score += (u.mailboxOver50GB || 0) > 0 ? 6 : 0;
  score += (u.archiveCount || 0) > 0 ? 5 : 0;
  score += d.devices.total > 100 ? 8 : d.devices.total > 0 ? 4 : 0;
  score = Math.min(100, score);

  // ---- Findings ----
  add('info', 'Scope', `${d.users.length} users (${activeUsers} active, ${d.guests} guests, ${d.disabled} disabled), ${d.groups.length} groups, ${d.teams} Teams, ${d.devices.total} managed devices.`);
  if (u.available) {
    add('info', 'Data', `~${totalDataGB.toLocaleString()} GB total: ${u.mailboxTotalGB} GB mailboxes, ${u.oneDriveTotalGB} GB OneDrive, ${u.spoTotalGB} GB SharePoint.`);
  } else {
    add('warn', 'Data', `Mailbox/OneDrive/SharePoint sizes not loaded (${u.note ?? 'Reports.Read.All needed'}). Data sizing is incomplete.`);
  }

  if ((u.mailboxOver50GB || 0) > 0) add('risk', 'Exchange', `${u.mailboxOver50GB} mailbox(es) over 50 GB — plan extended migration windows and verify destination quotas/Plan 2 licences.`);
  if ((u.archiveCount || 0) > 0) add('risk', 'Exchange', `${u.archiveCount} mailbox(es) have an online archive — archives migrate separately and auto-expanding archives can block cross-tenant moves.`);
  if (d.distributionGroups > 0) add('warn', 'Exchange', `${d.distributionGroups} distribution group(s) — recreate or migrate; check external sender settings and mail-enabled membership.`);
  if (d.teams > 0) add('risk', 'Teams', `${d.teams} Team(s) — Teams chat, private channels and meeting links do not migrate cleanly cross-tenant; plan recreation and user comms.`);
  if (d.m365Groups > 0) add('warn', 'Groups', `${d.m365Groups} Microsoft 365 group(s) connect mailbox + SharePoint + Teams — migrate as a unit.`);
  if (d.guests > 0) add('warn', 'Identity', `${d.guests} guest account(s) — B2B guests are re-invited in the target tenant, not migrated; refresh external sharing.`);
  if ((u.oneDriveOver100GB || 0) > 0) add('risk', 'OneDrive', `${u.oneDriveOver100GB} OneDrive(s) over 100 GB — long initial sync; pre-stage well before cutover.`);
  if (d.domains.length > 1) add('warn', 'Domains', `${d.domains.length} domains — a custom domain exists in only one tenant at a time; plan the domain cutover carefully.`);
  if (d.devices.total > 0) add('risk', 'Devices', `${d.devices.total} Intune-managed device(s) — enrollment cannot be moved; plan re-enrollment (some methods factory-reset). See Migration Edge Cases.`);
  if ((d.devices.nonCompliant || 0) > 0) add('warn', 'Devices', `${d.devices.nonCompliant} device(s) are non-compliant/unknown — Conditional Access can block these accounts during and after cutover; remediate before migrating their users.`);
  if ((d.compliancePolicies?.length || 0) > 0) add('info', 'Devices', `${d.compliancePolicies.length} compliance policy(ies) across ${[...new Set(d.compliancePolicies.map((p) => p.platform))].join(', ')} — recreate these in the target tenant before re-enrollment.`);
  const unverified = d.domains.filter((dom) => !dom.isVerified).length;
  if (unverified > 0) add('blocker', 'Domains', `${unverified} unverified domain(s) — resolve before any cutover.`);
  const noLocation = d.users.filter((x) => !x.usageLocation && x.accountEnabled).length;
  if (noLocation > 0) add('warn', 'Licensing', `${noLocation} active user(s) without a usage location — set it before assigning licences in the target.`);

  // ---- Sizing & duration (very rough planning estimate) ----
  // Assume ~25 GB/hour effective throughput per concurrent stream, 3 streams.
  const throughputGBperDay = 25 * 3 * 8; // ~600 GB/day effective
  const dataDays = Math.ceil(totalDataGB / throughputGBperDay);
  const userDays = Math.ceil(d.users.length / 200); // ~200 mailboxes/wave/day
  const estimatedDays = Math.max(5, dataDays + userDays + (d.teams > 0 ? 5 : 0) + (d.devices.total > 0 ? 10 : 0));

  const sizing: Assessment['sizing'] = [
    { label: 'Users to migrate', value: `${activeUsers} active (+${d.guests} guests re-invited)` },
    { label: 'Mailbox data', value: u.available ? `${u.mailboxTotalGB.toLocaleString()} GB / ${u.mailboxCount} mailboxes` : 'unknown' },
    { label: 'OneDrive data', value: u.available ? `${u.oneDriveTotalGB.toLocaleString()} GB / ${u.oneDriveCount} accounts` : 'unknown' },
    { label: 'SharePoint data', value: u.available ? `${u.spoTotalGB.toLocaleString()} GB / ${u.spoSiteCount} sites` : 'unknown' },
    { label: 'Total data', value: `~${totalDataGB.toLocaleString()} GB` },
    { label: 'Teams', value: `${d.teams}` },
    { label: 'Managed devices', value: `${d.devices.total}` },
    { label: 'Rough migration window', value: `~${estimatedDays} working days (planning estimate)` },
  ];

  // ---- Recommended approach ----
  let recommendedApproach: string;
  if (d.users.length < 150 && d.teams === 0 && totalDataGB < 500) {
    recommendedApproach = 'Small/simple — a single-wave cutover migration (native or BitTitan) is viable. Pilot first, then one cutover weekend.';
  } else if (score >= 50) {
    recommendedApproach = 'Complex — phased, wave-based migration with pre-staging. Use cross-tenant mailbox migration (native) or BitTitan, migrate Teams/SharePoint with a dedicated tool, and run device re-enrollment as a parallel workstream. Allow hypercare after each wave.';
  } else {
    recommendedApproach = 'Medium — wave-based migration with pre-staging of mailboxes and OneDrive, scheduled domain cutover, and recreated distribution groups. Pilot, then 2–4 waves.';
  }

  // ---- Preparation checklist ----
  const checklist: ChecklistItem[] = [
    { phase: 'Discovery', task: 'Validate this read-only inventory with the customer (users, data, Teams, devices).' },
    { phase: 'Discovery', task: 'Confirm source/target tenant IDs, admin accounts and licensing availability.' },
    { phase: 'Planning', task: `Decide migration method and tool (${score >= 50 ? 'cross-tenant native / BitTitan + Teams/SPO tooling' : 'native or BitTitan'}).` },
    { phase: 'Planning', task: 'Define waves/pilot group and the cutover date; write the comms plan.' },
    { phase: 'Identity', task: `Set usage location on the ${noLocation} user(s) missing it; map UPNs and primary SMTP.` },
    { phase: 'Domains', task: 'Plan domain cutover (a custom domain lives in one tenant only); lower MX/Autodiscover TTL.' },
    { phase: 'Exchange', task: `Prepare ${u.mailboxOver50GB || 0} large mailbox(es) and ${u.archiveCount || 0} archive(s); verify target quotas.` },
    { phase: 'Exchange', task: `Recreate ${d.distributionGroups} distribution group(s) and ${d.m365Groups} M365 group(s).` },
    { phase: 'Collaboration', task: d.teams > 0 ? `Plan recreation of ${d.teams} Team(s) and SharePoint sites; warn users about chat/meeting-link limits.` : 'No Teams detected — confirm with customer.' },
    { phase: 'Devices', task: d.devices.total > 0 ? `Plan re-enrollment of ${d.devices.total} device(s) — check enrollment methods for reset risk (Edge Cases module).` : 'No managed devices detected.' },
    { phase: 'Cutover', task: 'Pre-stage data, run the cutover wave, switch DNS, run final delta, validate mail flow.' },
    { phase: 'Post', task: 'Rebuild profiles, recreate permissions/signatures, hypercare period, decommission source per plan.' },
  ];

  return {
    complexityScore: score,
    complexityLabel: labelFor(score),
    totalDataGB,
    estimatedDays,
    recommendedApproach,
    findings,
    checklist,
    sizing,
  };
}
