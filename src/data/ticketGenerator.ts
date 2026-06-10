import { Ticket } from '../types';

// Per-service guidance used to build the generated ticket sections without AI.
const serviceGuides: Record<string, {
  firstSteps: string[]; causes: string[]; portals: string[]; powershell: string[]; evidence: string[];
}> = {
  'Entra ID': {
    firstSteps: ['Check Entra sign-in logs for the user and note the error code', 'Verify account enabled + license + group memberships', 'Check Conditional Access result on the failing sign-in', 'Check Risky users / risk detections', 'Test in private browser at office.com'],
    causes: ['Wrong/expired password', 'Conditional Access policy block', 'MFA method lost or misconfigured', 'Account disabled or sign-in blocked', 'License/service plan missing', 'Identity Protection risk block', 'Hybrid sync issue'],
    portals: ['Entra admin center > Protection > Sign-in logs', 'Entra admin center > Identity > Users', 'Entra admin center > Protection > Risky users'],
    powershell: ['Get-MgUser -UserId <upn> -Property AccountEnabled,OnPremisesSyncEnabled | fl', "Get-MgAuditLogSignIn -Filter \"userPrincipalName eq '<upn>'\" -Top 10"],
    evidence: ['Screenshot of the exact error', 'Sign-in log entry with Correlation ID', 'Time + device + network of the failure'],
  },
  'Exchange Online': {
    firstSteps: ['Run a message trace for the reported mail (sender/recipient/time)', 'Check inbox rules + forwarding on the mailbox', 'Check quarantine for the missing mail', 'Verify mailbox quota and license', 'Test in OWA to rule out the Outlook profile'],
    causes: ['Mail in quarantine/Junk', 'Inbox rule moving/deleting mail', 'Forwarding misconfigured', 'Mailbox over quota', 'Transport rule interference', 'Permission missing (shared mailbox/Send As)', 'DNS/SPF issue at sender side'],
    portals: ['Exchange admin center > Mail flow > Message trace', 'Defender portal > Review > Quarantine', 'Exchange admin center > Recipients > Mailboxes'],
    powershell: ['Get-MessageTrace -RecipientAddress <upn> -StartDate (Get-Date).AddDays(-2) -EndDate (Get-Date)', 'Get-InboxRule -Mailbox <upn>', 'Get-Mailbox <upn> | fl Forwarding*,*Quota'],
    evidence: ['Exact sender, recipient, subject, timestamp of an example mail', 'NDR text if any', 'Screenshot of the error in Outlook/OWA'],
  },
  'SharePoint Online': {
    firstSteps: ['Run Check Permissions for the user on the site/library', 'Determine site type (group-connected vs communication)', 'Check sharing settings tenant + site level', 'Check sign-in logs for CA/device blocks', 'Verify with a working colleague account for comparison'],
    causes: ['User not in the right group', 'Broken inheritance on library/folder', 'Sharing policy blocks external/guest', 'Conditional Access device restriction', 'Sensitivity label restriction', 'Link expired or wrong identity used'],
    portals: ['SharePoint admin center > Active sites', 'Site > Settings > Site permissions > Advanced', 'SharePoint admin center > Policies > Sharing'],
    powershell: ['Get-SPOSite <url> | fl SharingCapability,LockState', 'Get-SPOUser -Site <url> -LoginName <upn>'],
    evidence: ['Exact URL failing', 'Screenshot of the access denied page', 'Check Permissions output'],
  },
  Teams: {
    firstSteps: ['Test Teams on the web (separates client vs account)', 'Check Teams license/service plan', 'Check assigned Teams policies for the blocked feature', 'Check Service Health for Teams incidents', 'Clear Teams cache if client-only'],
    causes: ['Client cache corruption', 'Policy restricting the feature', 'License/service plan missing', 'CA block on Teams app', 'Guest/external settings', 'Microsoft service incident'],
    portals: ['Teams admin center > Users', 'Teams admin center > Meetings/Messaging policies', 'M365 admin center > Service health'],
    powershell: ['Get-CsOnlineUser -Identity <upn> | fl *Policy*', 'Connect-MicrosoftTeams'],
    evidence: ['Screenshot/error text', 'Web vs client behaviour', 'Meeting link/ID for meeting issues'],
  },
  'M365 Migration': {
    firstSteps: ['Check the migration batch/user status in the tool (EAC or MigrationWiz)', 'Read the exact error on the failed item/user', 'Verify endpoint credentials still valid', 'Check the relevant error helper (BitTitan/Cross-tenant modules)', 'Verify source/destination object state (license, GUIDs, scope)'],
    causes: ['Endpoint/credential failure', 'Throttling', 'Missing prerequisite (license, GUID, scope membership)', 'Oversized/corrupt items', 'Hold blocking the move'],
    portals: ['EAC > Migration', 'MigrationWiz project dashboard', 'M365 admin center > Service health'],
    powershell: ['Get-MigrationBatch | ft Identity,Status,FailedCount', 'Get-MigrationUser -Identity <user> | Get-MigrationUserStatistics | fl Status,Error'],
    evidence: ['Batch + user identity', 'Full error text from statistics', 'Timeline of passes already run'],
  },
  'Cross-Tenant Migration': {
    firstSteps: ['Get-MigrationUserStatistics for the failing user — read the exception name', 'Match it in the Cross-Tenant error helper', 'Verify MailUser prep: ExchangeGUID, X500, targetAddress', 'Verify license + scope group membership + no holds', 'Test-MigrationServerAvailability to validate the chain'],
    causes: ['Missing CT User Data Migration license', 'User not in scope group', 'Missing/wrong ExchangeGUID', 'Stale move request from earlier attempt', 'Hold on source mailbox', 'Domain not accepted in target'],
    portals: ['Target EAC > Migration', 'Entra admin center > App registrations (consent)', 'Cross-Tenant module in this portal'],
    powershell: ['Get-MigrationUser -Identity <user> | Get-MigrationUserStatistics | fl Status,Error', 'Get-MailUser <user> | fl ExchangeGuid,EmailAddresses,ExternalEmailAddress'],
    evidence: ['Exception name + full statistics output', 'MailUser property dump', 'Batch CSV row'],
  },
  BitTitan: {
    firstSteps: ['Open item statistics in MigrationWiz and read the error category', 'Re-verify endpoint credentials', 'Retry transient errors (throttling) once', 'Match persistent errors in the BitTitan error helper', 'Check destination quota/limits'],
    causes: ['Endpoint auth (modern auth, expired secret)', 'EWS throttling', 'Oversized items vs destination limits', 'Corrupt source items', 'Wrong source/destination mapping'],
    portals: ['MigrationWiz > Project > item statistics', 'BitTitan module in this portal'],
    powershell: ['Set-Mailbox <upn> -MaxReceiveSize 150MB # for oversized item errors'],
    evidence: ['Project name + item ID', 'Full error text', 'Retry history'],
  },
  Syskit: {
    firstSteps: ['Open the relevant Syskit report and export current data', 'Compare findings against agreed thresholds/baseline', 'Identify content/owners affected', 'Map findings to recommended actions in the Syskit module'],
    causes: ['Governance drift (sharing, guests, ownership)', 'Process gap (no owner lifecycle)', 'Configuration change unreported'],
    portals: ['Syskit Points > Reports/Governance', 'Syskit module in this portal'],
    powershell: [],
    evidence: ['Syskit report export', 'Affected workspace list', 'Baseline comparison'],
  },
  Security: {
    firstSteps: ['Triage the alert severity and affected entities', 'Contain if compromise indicated: reset password, revoke sessions', 'Preserve evidence (do NOT delete rules/mails)', 'Check related sign-ins and inbox rules', 'Escalate per the security escalation matrix'],
    causes: ['Phishing-led credential compromise', 'Malicious OAuth consent', 'Legacy auth abuse', 'Misconfiguration (forwarding, sharing)'],
    portals: ['Defender portal > Incidents & alerts', 'Entra admin center > Risky users/sign-ins'],
    powershell: ['Revoke-MgUserSignInSession -UserId <upn>', 'Get-InboxRule -Mailbox <upn>'],
    evidence: ['Alert IDs', 'Sign-in log exports', 'Timeline of containment actions'],
  },
  'Tenant Administration': {
    firstSteps: ['Confirm authorization/approval for the requested change', 'Identify the correct procedure in the Tenant Administration module', 'Check current state before changing anything', 'Execute with the documented steps and verify'],
    causes: ['Standard service request (not an incident)'],
    portals: ['M365 admin center', 'Entra admin center', 'Exchange admin center'],
    powershell: ['# See Tenant Administration module for the task-specific command'],
    evidence: ['Approval reference', 'Before/after state'],
  },
};

const nl = {
  ticketUpdate: (t: Ticket) => `Beste ${t.userName || 'gebruiker'},

Bedankt voor je melding. We hebben je ticket over "${t.category || t.service}" in behandeling genomen.

Wat we tot nu toe hebben gedaan:
- Eerste analyse uitgevoerd op basis van je melding
- <onderzoeksstappen invullen>

Volgende stappen:
- <volgende stap>

We houden je op de hoogte. Heb je aanvullende informatie (zoals een screenshot van de foutmelding), dan helpt dat ons enorm.

Met vriendelijke groet,
De servicedesk`,
  closure: (t: Ticket) => `Beste ${t.userName || 'gebruiker'},

Het probleem "${t.category || t.service}" is opgelost.

Oorzaak: <oorzaak>
Oplossing: <uitgevoerde actie>

Kun je bevestigen dat alles weer naar verwachting werkt? Zonder reactie sluiten we het ticket over 3 werkdagen.

Met vriendelijke groet,
De servicedesk`,
};

const en = {
  ticketUpdate: (t: Ticket) => `Dear ${t.userName || 'user'},

Thank you for your report. We are now working on your ticket regarding "${t.category || t.service}".

What we have done so far:
- Performed initial analysis based on your report
- <fill in investigation steps>

Next steps:
- <next step>

We will keep you informed. If you have additional information (such as a screenshot of the error), that would help us a lot.

Kind regards,
The service desk`,
  closure: (t: Ticket) => `Dear ${t.userName || 'user'},

The issue "${t.category || t.service}" has been resolved.

Root cause: <cause>
Resolution: <action taken>

Could you confirm everything works as expected? Without a response we will close the ticket in 3 business days.

Kind regards,
The service desk`,
};

export interface GeneratedTicket {
  shortDescription: string;
  fullDescription: string;
  troubleshooting: string;
  causes: string;
  portals: string;
  powershell: string;
  evidence: string;
  customerUpdate: string;
  escalationNote: string;
  closureNote: string;
}

export function generateTicket(t: Ticket): GeneratedTicket {
  const g = serviceGuides[t.service] ?? serviceGuides['Entra ID'];
  const lang = t.language === 'nl' ? nl : en;
  const upn = t.userEmail || '<upn>';

  return {
    shortDescription: `[${t.service}] ${t.category || 'Issue'} — ${t.userName || t.userEmail || 'user'} (${t.customer || 'customer'})`,
    fullDescription: `Customer/Tenant: ${t.customer || '-'}
Affected user: ${t.userName || '-'} (${t.userEmail || '-'})
Service: ${t.service}
Category: ${t.category || '-'}
Urgency: ${t.urgency.toUpperCase()} | Business impact: ${t.impact || '-'}

Description:
User reports: ${t.category || 'issue'} in ${t.service}.
Error message: ${t.errorMessage || 'none provided'}

Already tried (by user/requester):
${t.tried || '-'}

Screenshots available: ${t.screenshots ? 'yes (attached)' : 'no — requested'}
Reported: ${new Date(t.createdAt).toLocaleString()}`,
    troubleshooting: g.firstSteps.map((s, i) => `${i + 1}. ${s.replace(/<upn>/g, upn)}`).join('\n'),
    causes: g.causes.map((c) => `- ${c}`).join('\n'),
    portals: g.portals.map((p) => `- ${p}`).join('\n'),
    powershell: g.powershell.map((p) => p.replace(/<upn>/g, upn)).join('\n'),
    evidence: g.evidence.map((e) => `- ${e}`).join('\n'),
    customerUpdate: lang.ticketUpdate(t),
    escalationNote: `[Escalation — ${t.service}: ${t.category}]
Customer: ${t.customer} | User: ${t.userEmail} | Urgency: ${t.urgency.toUpperCase()}
Issue: ${t.category} — ${t.errorMessage || 'see description'}
Investigation done:
- ${g.firstSteps.slice(0, 3).join('\n- ')}
Findings: <fill in findings per step>
Evidence attached: <list>
Blocked because: <reason — role/permission/specialist knowledge>
Requested from escalation team: <specific request>`,
    closureNote: lang.closure(t),
  };
}

export const serviceOptions = Object.keys(serviceGuides);

export const categoryOptions: Record<string, string[]> = {
  'Entra ID': ['Login failure', 'Password reset', 'MFA issue', 'Conditional Access block', 'Account disabled', 'Risky user/sign-in', 'License issue', 'Group membership', 'Guest access', 'Sync issue', 'Admin role request'],
  'Exchange Online': ['Mail not received', 'Cannot send mail', 'Mailbox access', 'Shared mailbox', 'Send As / Send on Behalf', 'Mailbox full', 'Quarantine', 'Forwarding', 'Calendar/delegation', 'Room booking', 'Auto reply', 'Restore deleted mailbox', 'Connector/mail flow', 'SPF/DKIM/DMARC'],
  'SharePoint Online': ['Site access denied', 'Library access', 'Permission level', 'External sharing', 'Guest access', 'File restore', 'Site restore', 'Owner change', 'Storage', 'Sharing link', 'OneDrive sync'],
  Teams: ['Cannot access Teams', 'Client issue', 'Calendar missing', 'Cannot join meeting', 'Audio/video', 'Cannot create team', 'Channel issue', 'Guest access', 'External access', 'Recording', 'Room device', 'Policy'],
  'M365 Migration': ['Batch failure', 'User migration failure', 'Endpoint issue', 'Planning question', 'Cutover issue', 'Post-migration issue'],
  'Cross-Tenant Migration': ['Setup/consent issue', 'MailUser preparation', 'Batch failure', 'Move request error', 'License issue', 'Post-migration issue'],
  BitTitan: ['Endpoint/auth failure', 'Failed items', 'Throttling', 'Oversized items', 'Licensing', 'Reporting'],
  Syskit: ['Report request', 'Governance finding', 'Access review', 'Configuration'],
  Security: ['Suspected compromise', 'Phishing report', 'Defender alert', 'Risky user', 'Policy question', 'Hardening request'],
  'Tenant Administration': ['Create user', 'Offboard user', 'License change', 'Group change', 'Shared mailbox', 'Distribution group', 'Domain', 'Other request'],
};
