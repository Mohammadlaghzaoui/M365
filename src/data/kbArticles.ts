import { KBArticle } from '../types';

export const seedArticles: KBArticle[] = [
  {
    id: 'kb-reset-mfa', title: 'Reset MFA for a user', service: 'Entra ID', level: 'L1', tags: ['mfa', 'authentication', 'identity'], favorite: false,
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
    id: 'kb-signin-logs', title: 'Check Entra sign-in logs', service: 'Entra ID', level: 'L1', tags: ['signin', 'logs', 'troubleshooting'], favorite: false,
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
    id: 'kb-message-trace', title: 'Run a message trace', service: 'Exchange Online', level: 'L1', tags: ['mailflow', 'trace'], favorite: false,
    body: `EAC > Mail flow > Message trace, or:
Connect-ExchangeOnline
Get-MessageTrace -SenderAddress a@ext.com -RecipientAddress u@contoso.com -StartDate (Get-Date).AddDays(-2) -EndDate (Get-Date)

Status meanings: Delivered (check Junk/rules) | FilteredAsSpam (quarantine) | Failed (read detail) | Pending (retrying) | None (never reached tenant).
Detail: ... | Get-MessageTraceDetail shows transport rule hits.
>10 days: Start-HistoricalSearch (up to 90 days).`,
  },
  {
    id: 'kb-quarantine', title: 'Check and release quarantine', service: 'Exchange Online', level: 'L1', tags: ['quarantine', 'defender', 'spam'], favorite: false,
    body: `Defender portal > Email & collaboration > Review > Quarantine.
1. Search by recipient/sender. 2. Check the reason (spam/phish/malware). 3. Preview safely in portal.
4. Release only clearly legitimate mail. NEVER release malware/high-confidence phish — escalate.
PowerShell:
Get-QuarantineMessage -RecipientAddress u@contoso.com -StartReceivedDate (Get-Date).AddDays(-7)
Release-QuarantineMessage -Identity <id> -ReleaseToAll
Recurring FPs: fix sender SPF/DKIM; allow-listing is a security-team action.`,
  },
  {
    id: 'kb-shared-mailbox-perm', title: 'Add shared mailbox permission', service: 'Exchange Online', level: 'L1', tags: ['shared mailbox', 'permissions'], favorite: false,
    body: `Full Access (open the mailbox):
Add-MailboxPermission shared@contoso.com -User user@contoso.com -AccessRights FullAccess -AutoMapping:$true
Send As (send as the address):
Add-RecipientPermission shared@contoso.com -Trustee user@contoso.com -AccessRights SendAs -Confirm:$false
Both cached up to 60 min. Test instantly via OWA > Open another mailbox.
Always record the approver in the ticket.`,
  },
  {
    id: 'kb-sendas-fix', title: 'Fix Send As issue', service: 'Exchange Online', level: 'L1', tags: ['sendas', 'permissions'], favorite: false,
    body: `Symptom: NDR "you do not have permission to send on behalf of".
1. Full Access is NOT enough — Send As is separate.
2. Check: Get-RecipientPermission shared@contoso.com
3. Grant: Add-RecipientPermission shared@contoso.com -Trustee user@contoso.com -AccessRights SendAs -Confirm:$false
4. Wait up to 60 min (cache) / restart Outlook.
5. Verify the From field is set to the shared address.`,
  },
  {
    id: 'kb-room-booking', title: 'Troubleshoot room mailbox booking', service: 'Exchange Online', level: 'L2', tags: ['room', 'calendar', 'booking'], favorite: false,
    body: `Get-CalendarProcessing room@contoso.com | fl
Checklist: AutomateProcessing=AutoAccept | BookingWindowInDays vs meeting date | MaximumDurationInMinutes vs length | AllowRecurringMeetings + ConflictPercentageAllowed for series | BookInPolicy if only some users fail | ResourceDelegates may decline later.
Room "disappears after booking": meeting UPDATES are re-evaluated (window/duration), delegates can decline afterwards, users with calendar rights can delete. Trace decline mails:
Get-MessageTrace -SenderAddress room@... -RecipientAddress organizer@...
Evidence for escalation: Get-CalendarDiagnosticObjects.`,
  },
  {
    id: 'kb-teams-calendar', title: 'Troubleshoot Teams calendar missing', service: 'Teams', level: 'L2', tags: ['teams', 'calendar'], favorite: false,
    body: `1. Mailbox must be in Exchange ONLINE (on-prem mailbox = known limitation, needs hybrid EWS OAuth).
2. Check OWA calendar works.
3. Teams admin center > user > Policies: App setup policy must include/pin Calendar.
4. Clear Teams cache after fixes; policy changes take up to 24h.
Get-CsOnlineUser -Identity u@c.com | Select TeamsAppSetupPolicy`,
  },
  {
    id: 'kb-onedrive-sync', title: 'Troubleshoot OneDrive sync', service: 'SharePoint Online', level: 'L1', tags: ['onedrive', 'sync'], favorite: false,
    body: `1. Read the actual state from the OneDrive cloud icon (errors list).
2. Verify browser access to the library (permission loss stops sync silently).
3. Common: path >400 chars, invalid chars, >300k items, locked files.
4. Reset: %localappdata%\\Microsoft\\OneDrive\\onedrive.exe /reset (no data loss, full resync).
5. Huge libraries: use "Add shortcut to OneDrive" instead of full sync.`,
  },
  {
    id: 'kb-spo-access-denied', title: 'Troubleshoot SharePoint access denied', service: 'SharePoint Online', level: 'L1', tags: ['sharepoint', 'permissions'], favorite: false,
    body: `1. Site > Settings > Site permissions > Advanced > Check Permissions (user UPN) → effective access + source.
2. Group-connected site: fix membership via the M365 GROUP, not direct SPO grants.
3. Library/folder denies while site works: broken inheritance — check unique permissions at that level.
4. Permissions fine? Check CA policies (unmanaged device restrictions) in sign-in logs.
5. Sensitivity label / restricted access control can override — escalate to compliance.`,
  },
  {
    id: 'kb-prepare-migration', title: 'Prepare an M365 migration (checklist)', service: 'Migration', level: 'L2', tags: ['migration', 'planning'], favorite: false,
    body: `Discovery: users, mailbox sizes, shared/resource mailboxes, groups, Teams/SPO inventory, domains+DNS export, third-party mail senders, holds/journaling.
Plan: method (native cross-tenant / BitTitan / hybrid), waves, pilot, cutover date, rollback, comms.
Prep: target licenses, identity model, MX/Autodiscover TTL lowering, admin/app credentials both sides.
Execute: pilot → pre-stage → cutover (DNS+identity+final pass) → delta → validate.
Close: permissions recreated, profiles rebuilt, reports exported, migration credentials removed.
Use the Migration module to track stages per project.`,
  },
  {
    id: 'kb-ct-checklist', title: 'Cross-tenant mailbox migration checklist', service: 'Cross-Tenant Migration', level: 'L3', tags: ['cross-tenant', 'migration', 'exchange'], favorite: false,
    body: `Target: app registration + Mailbox.Migration permission + admin consent (source side via consent URL) | migration endpoint | org relationship (Inbound).
Source: consent granted | mail-enabled security group as scope | org relationship (RemoteOutbound + OAuthApplicationId + PublishedScopes).
Per user: MailUser in target with source ExchangeGUID + ArchiveGUID + X500(LegacyExchangeDN) | targetAddress→source | CT User Data Migration license BOTH sides | no holds | in scope group | no premature target mailbox.
Validate: Test-MigrationServerAvailability per batch.
Run: New-MigrationBatch (CSV header EmailAddress with TARGET identities) → monitor → Complete.
Remember: source mailbox is converted after completion — no rollback; Teams meeting links break; signatures don't migrate.
Full generator in the Cross-Tenant module.`,
  },
  {
    id: 'kb-t2t-manual', title: 'Tenant-to-tenant migration: manual domain move runbook', service: 'Migration', level: 'L3', tags: ['tenant-to-tenant', 'domain', 'dns', 'cutover'], favorite: false,
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
    id: 'kb-gmail-m365', title: 'Google Workspace (Gmail) to M365 migration guide', service: 'Migration', level: 'L2', tags: ['google', 'gmail', 'migration'], favorite: false,
    body: `Google side: Cloud project + service account + JSON key | enable Gmail/Calendar/People APIs | domain-wide delegation with Microsoft's scope list | note super admin.
M365 side: users + Exchange licenses FIRST (mailboxes must exist) | EAC > Migration > Google Workspace > endpoint with JSON key + super admin.
Batches: CSV header "EmailAddress" (M365 addresses) | pilot first | labels become folders (multi-label items duplicate into multiple folders — tell users!).
Cutover: MX from Google to M365, SPF switch, DKIM enable, complete batches (final delta automatic).
Drive→OneDrive: Migration Manager in SharePoint admin center (separate project). Not migrated: Chat, Sites, Forms.`,
  },
  {
    id: 'kb-bittitan-failed', title: 'BitTitan failed items handling', service: 'BitTitan', level: 'L2', tags: ['bittitan', 'migrationwiz', 'errors'], favorite: false,
    body: `1. Open project > item > statistics: read error category.
2. Transient (throttling/503/ErrorServerBusy): just Retry Errors — usually clears.
3. Auth errors: endpoint problem — fix FIRST, blocks everything (modern auth app, secret expiry).
4. Oversized: raise destination limits (Set-Mailbox -MaxReceiveSize 150MB) then retry.
5. Corrupt items: retry once; persistent = document for customer (normal residue).
6. Export the error report for the project file. Escalate to BitTitan support with item IDs if systematic.`,
  },
  {
    id: 'kb-syskit-guest', title: 'Syskit guest review procedure', service: 'Syskit', level: 'L2', tags: ['syskit', 'guests', 'governance'], favorite: false,
    body: `1. Syskit Points > Governance > Guest users (last activity + memberships).
2. Flag: inactive >90 days, unexpected domains, guests in many teams.
3. Send owner-driven review (Syskit access review) — owners confirm keep/remove.
4. Remove confirmed-stale guests; document list in ticket.
5. Escalate immediately: guest from competitor/unknown domain with sensitive access.
Frequency: quarterly minimum.`,
  },
  {
    id: 'kb-security-quickwins', title: 'Security quick wins (first week)', service: 'Security', level: 'L2', tags: ['security', 'hardening', 'quickwins'], favorite: false,
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
  {
    id: 'kb-outlook-password-prompt', title: 'Outlook keeps asking for password', service: 'Exchange Online', level: 'L1', tags: ['outlook', 'credentials', 'profile'], favorite: false,
    body: `1. Check OWA works (rules out account/mailbox issues).
2. Windows: clear stored credentials — Credential Manager > Windows Credentials > remove all MicrosoftOffice/Outlook entries.
3. Sign out/in of Office: File > Office Account > Sign out, restart, sign in.
4. Check sign-in logs: repeated prompts often = CA policy or MFA loop (token issue).
5. Still looping: new Outlook profile (Control Panel > Mail > Profiles).
6. Multiple users affected at once → check Service Health + CA changes (escalate).`,
  },
  {
    id: 'kb-account-lockout', title: 'Account lockout / smart lockout loop', service: 'Entra ID', level: 'L1', tags: ['lockout', 'password', 'signin'], favorite: false,
    body: `Error 50053 = smart lockout. Cause: bad password attempts (often an OLD password saved somewhere).
1. Hunt the stale credential: phone mail app, other PC, mapped drives, scheduled tasks, scan-to-mail apps.
2. Hybrid: check on-prem AD lockout too (badPwdCount, lockoutTime) — find source with Account Lockout Status tool / DC security log 4740.
3. Lockout clears automatically; do not just reset the password again without removing the stale credential.
4. Repeated external attempts (spray): check sign-in log IPs — unfamiliar = security escalation (L2/L3).`,
  },
  {
    id: 'kb-add-alias', title: 'Add email alias / change primary SMTP', service: 'Exchange Online', level: 'L1', tags: ['alias', 'smtp', 'proxyaddresses'], favorite: false,
    body: `Cloud user: M365 admin center > Users > [user] > Username/email > Manage email aliases.
PowerShell:
Set-Mailbox user@contoso.com -EmailAddresses @{Add='smtp:alias@contoso.com'}
Primary change: 'SMTP:' (capital) prefix sets primary:
Set-Mailbox user@contoso.com -WindowsEmailAddress newprimary@contoso.com
HYBRID: edit proxyAddresses in on-prem AD (cloud edit fails for synced users) and sync.
Warn user: primary change affects sign-in display and sent-from; old address stays as alias (keep it!).`,
  },
  {
    id: 'kb-outlook-profile-rebuild', title: 'Rebuild Outlook profile (standard procedure)', service: 'Exchange Online', level: 'L1', tags: ['outlook', 'profile', 'client'], favorite: false,
    body: `When: profile corruption, post-migration, persistent client weirdness with healthy OWA.
1. Warn: local PST archives stay, but note signatures and local settings.
2. Close Outlook > Control Panel > Mail (Microsoft Outlook) > Show Profiles > Add new profile.
3. Set the new profile as default ("Always use this profile"), start Outlook, autodiscover signs in.
4. Verify mail/calendar/shared mailboxes load; re-add PSTs if used.
5. Delete the old profile only after the user confirms everything works.`,
  },
  {
    id: 'kb-mobile-mail-setup', title: 'Set up mail on mobile (iOS/Android)', service: 'Exchange Online', level: 'L1', tags: ['mobile', 'outlook', 'activesync'], favorite: false,
    body: `Recommended: Outlook app (supports modern auth + app protection policies).
1. Install Outlook app > Add account > work address > M365 sign-in + MFA.
2. Post-migration: REMOVE the old account first, then add fresh.
3. Native iOS Mail: Settings > Mail > Accounts > Add > Microsoft Exchange (only if customer policy allows).
4. Blocked? Check CA policies (require approved app / compliant device) and Intune enrollment requirement — explain the customer policy, do not bypass.`,
  },
  {
    id: 'kb-teams-status-presence', title: 'Teams presence/status stuck or wrong', service: 'Teams', level: 'L1', tags: ['teams', 'presence', 'status'], favorite: false,
    body: `1. Have the user reset status manually (avatar > status > Reset status).
2. Presence follows Outlook calendar — check for a stale "in a meeting" appointment.
3. Quit Teams fully and clear cache (see Teams cache article).
4. Multiple devices: another signed-in device (old laptop) can hold the status — sign out everywhere.
5. Org-wide wrong presence → Service Health check.`,
  },
  {
    id: 'kb-duplicate-proxy-sync-error', title: 'Fix duplicate proxyAddress / UPN sync conflict', service: 'Entra ID', level: 'L2', tags: ['hybrid', 'sync', 'proxyaddresses', 'entra connect'], favorite: false,
    body: `Symptom: Entra Connect export error "AttributeValueMustBeUnique" / user gets onmicrosoft address.
1. Find the conflicting object: Entra admin center > Entra Connect > sync errors lists both objects.
2. Or search: Get-EntraUser/Get-MgUser -Filter on proxyAddresses, AND check deleted users (soft-deleted objects hold addresses!).
3. Decide the rightful owner; remove the address from the other object in its source (AD or cloud).
4. Run delta sync; verify the error clears and the address lands correctly.
Classic causes: re-created users, restored mailboxes, contacts holding the SMTP.`,
  },
  {
    id: 'kb-transport-rule-troubleshoot', title: 'Troubleshoot transport (mail flow) rules', service: 'Exchange Online', level: 'L2', tags: ['transport rules', 'mailflow'], favorite: false,
    body: `1. Message trace DETAIL shows which rule hit: ... | Get-MessageTraceDetail | where Event -eq 'Transport rule'.
2. List rules: Get-TransportRule | select Name,State,Priority,Mode.
3. Test safely: set rule Mode to 'Audit' (TestWithoutPolicyTips) instead of disabling in production.
4. Watch priority order — first matching rule with 'stop processing' wins.
5. Changes to transport rules = change management (affect ALL mail). Document before/after.`,
  },
  {
    id: 'kb-ca-troubleshooting', title: 'Conditional Access troubleshooting method', service: 'Entra ID', level: 'L2', tags: ['conditional access', 'signin', 'policy'], favorite: false,
    body: `1. Sign-in logs > failing entry > Conditional Access tab: per policy Success/Failure/Not applied.
2. Failure row > which condition matched and which control failed (MFA? compliant device? location?).
3. Use the What If tool (CA > Policies > What If) to simulate the user/app/location combination.
4. Device compliance failures: cross-check the device in Intune (compliant? enrolled? cert valid?).
5. Report-only policies show impact without blocking — read them too.
6. NEVER fix by exclusion at L1/L2; propose remediation to the policy owner with CorrelationId evidence.`,
  },
  {
    id: 'kb-compromise-response', title: 'Account compromise response runbook', service: 'Security', level: 'L3', tags: ['incident', 'compromise', 'bec', 'security'], favorite: false,
    body: `CONTAIN (minutes):
1. Reset password + Revoke-MgUserSignInSession (kill tokens).
2. Block sign-in temporarily if active attacker.
3. Disable suspicious inbox rules (do NOT delete — evidence) + remove unknown forwarding.
4. Check registered MFA methods for attacker-added devices; remove them.
INVESTIGATE:
5. Sign-in logs: first malicious login, IP, persistence window.
6. Audit log: rule creation, OAuth consents (revoke malicious app grants), mailbox exports.
7. Message trace: what was sent (internal phishing wave?).
8. Check OTHER users hit from the same IPs.
RECOVER & REPORT:
9. MFA re-register, user briefing, monitor 30 days.
10. Full timeline to security team/customer; consider breach notification obligations (customer decision).`,
  },
  {
    id: 'kb-hybrid-mailflow', title: 'Hybrid mail flow troubleshooting (on-prem ↔ EXO)', service: 'Migration', level: 'L3', tags: ['hybrid', 'mailflow', 'connectors', 'hcw'], favorite: false,
    body: `1. Determine direction failing: on-prem→cloud, cloud→on-prem, or external in/out.
2. Trace both sides: Get-MessageTrace (EXO) + Get-MessageTrackingLog (on-prem).
3. Check hybrid connectors: EAC inbound connector (cert name must match the on-prem cert EXACTLY — renewals break this silently) + on-prem send connector to EXO.
4. TLS test: Test-SmtpConnectivity / openssl s_client to verify the presented certificate.
5. Centralized mail transport on? Mail may route via on-prem by design.
6. Fixes via Hybrid Configuration Wizard re-run, NOT manual connector edits.
7. Escalation: HCW logs + traces both sides + cert chain output.`,
  },
  {
    id: 'kb-mrs-proxy-endpoint', title: 'MRS Proxy / hybrid migration endpoint errors', service: 'Migration', level: 'L3', tags: ['hybrid', 'mrs proxy', 'remote move', 'endpoint'], favorite: false,
    body: `Symptoms: "The connection to the server could not be completed" / Test-MigrationServerAvailability fails for ExchangeRemoteMove.
1. Enable MRS proxy on-prem: Set-WebServicesVirtualDirectory -Identity "EWS (Default Web Site)" -MRSProxyEnabled \$true (then iisreset).
2. Test: Test-MigrationServerAvailability -ExchangeRemoteMove -RemoteServer mail.contoso.com -Credentials (Get-Credential).
3. EWS must be published externally (443) with a valid 3rd-party certificate; no pre-auth on /EWS/mrsproxy.svc (bypass in proxy/WAF!).
4. 401: use on-prem DOMAIN\\user credentials with migration rights.
5. Timeouts mid-migration: MRS proxy throttling — raise limits in EWS web.config per Microsoft guidance.
6. Escalation evidence: full Test-MigrationServerAvailability output + IIS logs of /mrsproxy.svc hits.`,
  },
];
