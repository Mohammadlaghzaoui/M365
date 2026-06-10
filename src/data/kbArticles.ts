import { KBArticle } from '../types';

export const seedArticles: KBArticle[] = [
  {
    id: 'kb-reset-mfa', title: 'Reset MFA for a user', service: 'Entra ID', tags: ['mfa', 'authentication', 'identity'], favorite: false,
    body: `1. VERIFY IDENTITY first (callback/manager confirmation) — MFA reset is the #1 social-engineering target.
2. Entra admin center > Users > [user] > Authentication methods.
3. Either "Require re-register MFA" or delete the lost method directly.
4. User registers the new device at https://aka.ms/mfasetup.
5. Confirm a successful MFA sign-in afterwards.

PowerShell:
Connect-MgGraph -Scopes 'UserAuthenticationMethod.ReadWrite.All'
Get-MgUserAuthenticationMethod -UserId user@contoso.com

Role: Authentication Administrator. Never reset for admins at L1 — escalate.`,
  },
  {
    id: 'kb-signin-logs', title: 'Check Entra sign-in logs', service: 'Entra ID', tags: ['signin', 'logs', 'troubleshooting'], favorite: false,
    body: `Entra admin center > Protection > Sign-in logs. Filter on user/app/status.
Key error codes:
- 50126 wrong password
- 50053 account locked / smart lockout
- 50057 account disabled
- 50076 / 50079 MFA required/registration
- 53003 blocked by Conditional Access
- 50105 not assigned to the application
- 65001 consent required
Open the entry > Conditional Access tab to see which policy failed.
Copy the Correlation ID into the ticket for escalations.`,
  },
  {
    id: 'kb-message-trace', title: 'Run a message trace', service: 'Exchange Online', tags: ['mailflow', 'trace'], favorite: false,
    body: `EAC > Mail flow > Message trace, or:
Connect-ExchangeOnline
Get-MessageTrace -SenderAddress a@ext.com -RecipientAddress u@contoso.com -StartDate (Get-Date).AddDays(-2) -EndDate (Get-Date)

Status meanings: Delivered (check Junk/rules) | FilteredAsSpam (quarantine) | Failed (read detail) | Pending (retrying) | None (never reached tenant).
Detail: ... | Get-MessageTraceDetail shows transport rule hits.
>10 days: Start-HistoricalSearch (up to 90 days).`,
  },
  {
    id: 'kb-quarantine', title: 'Check and release quarantine', service: 'Exchange Online', tags: ['quarantine', 'defender', 'spam'], favorite: false,
    body: `Defender portal > Email & collaboration > Review > Quarantine.
1. Search by recipient/sender. 2. Check the reason (spam/phish/malware). 3. Preview safely in portal.
4. Release only clearly legitimate mail. NEVER release malware/high-confidence phish — escalate.
PowerShell:
Get-QuarantineMessage -RecipientAddress u@contoso.com -StartReceivedDate (Get-Date).AddDays(-7)
Release-QuarantineMessage -Identity <id> -ReleaseToAll
Recurring FPs: fix sender SPF/DKIM; allow-listing is a security-team action.`,
  },
  {
    id: 'kb-shared-mailbox-perm', title: 'Add shared mailbox permission', service: 'Exchange Online', tags: ['shared mailbox', 'permissions'], favorite: false,
    body: `Full Access (open the mailbox):
Add-MailboxPermission shared@contoso.com -User user@contoso.com -AccessRights FullAccess -AutoMapping:$true
Send As (send as the address):
Add-RecipientPermission shared@contoso.com -Trustee user@contoso.com -AccessRights SendAs -Confirm:$false
Both cached up to 60 min. Test instantly via OWA > Open another mailbox.
Always record the approver in the ticket.`,
  },
  {
    id: 'kb-sendas-fix', title: 'Fix Send As issue', service: 'Exchange Online', tags: ['sendas', 'permissions'], favorite: false,
    body: `Symptom: NDR "you do not have permission to send on behalf of".
1. Full Access is NOT enough — Send As is separate.
2. Check: Get-RecipientPermission shared@contoso.com
3. Grant: Add-RecipientPermission shared@contoso.com -Trustee user@contoso.com -AccessRights SendAs -Confirm:$false
4. Wait up to 60 min (cache) / restart Outlook.
5. Verify the From field is set to the shared address.`,
  },
  {
    id: 'kb-room-booking', title: 'Troubleshoot room mailbox booking', service: 'Exchange Online', tags: ['room', 'calendar', 'booking'], favorite: false,
    body: `Get-CalendarProcessing room@contoso.com | fl
Checklist: AutomateProcessing=AutoAccept | BookingWindowInDays vs meeting date | MaximumDurationInMinutes vs length | AllowRecurringMeetings + ConflictPercentageAllowed for series | BookInPolicy if only some users fail | ResourceDelegates may decline later.
Room "disappears after booking": meeting UPDATES are re-evaluated (window/duration), delegates can decline afterwards, users with calendar rights can delete. Trace decline mails:
Get-MessageTrace -SenderAddress room@... -RecipientAddress organizer@...
Evidence for escalation: Get-CalendarDiagnosticObjects.`,
  },
  {
    id: 'kb-teams-calendar', title: 'Troubleshoot Teams calendar missing', service: 'Teams', tags: ['teams', 'calendar'], favorite: false,
    body: `1. Mailbox must be in Exchange ONLINE (on-prem mailbox = known limitation, needs hybrid EWS OAuth).
2. Check OWA calendar works.
3. Teams admin center > user > Policies: App setup policy must include/pin Calendar.
4. Clear Teams cache after fixes; policy changes take up to 24h.
Get-CsOnlineUser -Identity u@c.com | Select TeamsAppSetupPolicy`,
  },
  {
    id: 'kb-onedrive-sync', title: 'Troubleshoot OneDrive sync', service: 'SharePoint Online', tags: ['onedrive', 'sync'], favorite: false,
    body: `1. Read the actual state from the OneDrive cloud icon (errors list).
2. Verify browser access to the library (permission loss stops sync silently).
3. Common: path >400 chars, invalid chars, >300k items, locked files.
4. Reset: %localappdata%\\Microsoft\\OneDrive\\onedrive.exe /reset (no data loss, full resync).
5. Huge libraries: use "Add shortcut to OneDrive" instead of full sync.`,
  },
  {
    id: 'kb-spo-access-denied', title: 'Troubleshoot SharePoint access denied', service: 'SharePoint Online', tags: ['sharepoint', 'permissions'], favorite: false,
    body: `1. Site > Settings > Site permissions > Advanced > Check Permissions (user UPN) → effective access + source.
2. Group-connected site: fix membership via the M365 GROUP, not direct SPO grants.
3. Library/folder denies while site works: broken inheritance — check unique permissions at that level.
4. Permissions fine? Check CA policies (unmanaged device restrictions) in sign-in logs.
5. Sensitivity label / restricted access control can override — escalate to compliance.`,
  },
  {
    id: 'kb-prepare-migration', title: 'Prepare an M365 migration (checklist)', service: 'Migration', tags: ['migration', 'planning'], favorite: false,
    body: `Discovery: users, mailbox sizes, shared/resource mailboxes, groups, Teams/SPO inventory, domains+DNS export, third-party mail senders, holds/journaling.
Plan: method (native cross-tenant / BitTitan / hybrid), waves, pilot, cutover date, rollback, comms.
Prep: target licenses, identity model, MX/Autodiscover TTL lowering, admin/app credentials both sides.
Execute: pilot → pre-stage → cutover (DNS+identity+final pass) → delta → validate.
Close: permissions recreated, profiles rebuilt, reports exported, migration credentials removed.
Use the Migration module to track stages per project.`,
  },
  {
    id: 'kb-ct-checklist', title: 'Cross-tenant mailbox migration checklist', service: 'Cross-Tenant Migration', tags: ['cross-tenant', 'migration', 'exchange'], favorite: false,
    body: `Target: app registration + Mailbox.Migration permission + admin consent (source side via consent URL) | migration endpoint | org relationship (Inbound).
Source: consent granted | mail-enabled security group as scope | org relationship (RemoteOutbound + OAuthApplicationId + PublishedScopes).
Per user: MailUser in target with source ExchangeGUID + ArchiveGUID + X500(LegacyExchangeDN) | targetAddress→source | CT User Data Migration license BOTH sides | no holds | in scope group | no premature target mailbox.
Validate: Test-MigrationServerAvailability per batch.
Run: New-MigrationBatch (CSV header EmailAddress with TARGET identities) → monitor → Complete.
Remember: source mailbox is converted after completion — no rollback; Teams meeting links break; signatures don't migrate.
Full generator in the Cross-Tenant module.`,
  },
  {
    id: 'kb-t2t-manual', title: 'Tenant-to-tenant migration: manual domain move runbook', service: 'Migration', tags: ['tenant-to-tenant', 'domain', 'dns', 'cutover'], favorite: false,
    body: `A custom domain can exist in only ONE tenant. The manual move:
1. Lower MX/Autodiscover TTL 24-48h ahead.
2. Pre-stage mailbox data to target (tool of choice) while users work in source.
3. In SOURCE: switch all UPNs + primary/alias addresses to source.onmicrosoft.com (script proxyAddresses cleanup!).
4. Remove domain from source tenant (fails while ANY reference remains — find leftovers with Get-Recipient | ? {$_.EmailAddresses -match "domain.com"}).
5. Add + verify domain in TARGET (TXT record); set UPNs/primary SMTP in target.
6. DNS cutover: MX → target, Autodiscover CNAME, SPF include target, DKIM enable in target, DMARC keep.
7. Final delta pass; rebuild Outlook/mobile profiles; recreate permissions/signatures.
8. Validate mail flow both directions; keep source for fallback window.`,
  },
  {
    id: 'kb-gmail-m365', title: 'Google Workspace (Gmail) to M365 migration guide', service: 'Migration', tags: ['google', 'gmail', 'migration'], favorite: false,
    body: `Google side: Cloud project + service account + JSON key | enable Gmail/Calendar/People APIs | domain-wide delegation with Microsoft's scope list | note super admin.
M365 side: users + Exchange licenses FIRST (mailboxes must exist) | EAC > Migration > Google Workspace > endpoint with JSON key + super admin.
Batches: CSV header "EmailAddress" (M365 addresses) | pilot first | labels become folders (multi-label items duplicate into multiple folders — tell users!).
Cutover: MX from Google to M365, SPF switch, DKIM enable, complete batches (final delta automatic).
Drive→OneDrive: Migration Manager in SharePoint admin center (separate project). Not migrated: Chat, Sites, Forms.`,
  },
  {
    id: 'kb-bittitan-failed', title: 'BitTitan failed items handling', service: 'BitTitan', tags: ['bittitan', 'migrationwiz', 'errors'], favorite: false,
    body: `1. Open project > item > statistics: read error category.
2. Transient (throttling/503/ErrorServerBusy): just Retry Errors — usually clears.
3. Auth errors: endpoint problem — fix FIRST, blocks everything (modern auth app, secret expiry).
4. Oversized: raise destination limits (Set-Mailbox -MaxReceiveSize 150MB) then retry.
5. Corrupt items: retry once; persistent = document for customer (normal residue).
6. Export the error report for the project file. Escalate to BitTitan support with item IDs if systematic.`,
  },
  {
    id: 'kb-syskit-guest', title: 'Syskit guest review procedure', service: 'Syskit', tags: ['syskit', 'guests', 'governance'], favorite: false,
    body: `1. Syskit Points > Governance > Guest users (last activity + memberships).
2. Flag: inactive >90 days, unexpected domains, guests in many teams.
3. Send owner-driven review (Syskit access review) — owners confirm keep/remove.
4. Remove confirmed-stale guests; document list in ticket.
5. Escalate immediately: guest from competitor/unknown domain with sensitive access.
Frequency: quarterly minimum.`,
  },
  {
    id: 'kb-security-quickwins', title: 'Security quick wins (first week)', service: 'Security', tags: ['security', 'hardening', 'quickwins'], favorite: false,
    body: `1. MFA for all users + admins (CA or Security Defaults).
2. Block legacy authentication (check sign-in logs for legacy usage first).
3. Break-glass accounts (2, cloud-only, CA-excluded, monitored).
4. Block external auto-forwarding (outbound spam policy).
5. Admin role review: 2-4 GAs, least privilege for the rest.
6. Enable unified audit logging.
7. SPF/DKIM for all domains.
8. Restrict user app consent.
9. Anti-malware common attachment filter on.
10. Review Secure Score for tenant-specific items.
Track all of this in the Security Hardening module.`,
  },
];
