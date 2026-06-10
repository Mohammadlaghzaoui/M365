import { MigrationError } from '../types';

export interface CTForm {
  sourceTenantId: string;
  targetTenantId: string;
  sourceOnMicrosoft: string;
  targetOnMicrosoft: string;
  appClientId: string;
  clientSecretName: string;
  endpointName: string;
  orgRelationshipName: string;
  scopeGroupName: string;
  targetDeliveryDomain: string;
  csvUsers: string;
}

export const CT_DEFAULTS: CTForm = {
  sourceTenantId: '',
  targetTenantId: '',
  sourceOnMicrosoft: 'source.onmicrosoft.com',
  targetOnMicrosoft: 'target.onmicrosoft.com',
  appClientId: '',
  clientSecretName: 'CrossTenantMigrationSecret',
  endpointName: 'CrossTenantEndpoint',
  orgRelationshipName: 'CrossTenantOrgRelationship',
  scopeGroupName: 'CrossTenantMigrationScope',
  targetDeliveryDomain: 'target.onmicrosoft.com',
  csvUsers: 'EmailAddress\nuser1@target.onmicrosoft.com\nuser2@target.onmicrosoft.com',
};

export const ctWarnings = [
  'Mailboxes on hold (litigation/eDiscovery) are BLOCKED from cross-tenant migration — remove or handle holds first.',
  'Only user-visible mailbox content migrates (no recoverable items beyond standard scope).',
  'The SOURCE mailbox is converted/DELETED after successful migration — there is no automatic rollback.',
  'Every migrating user needs the Cross Tenant User Data Migration license (assigned in BOTH source and target).',
  'The target MailUser MUST have the correct ExchangeGUID stamped before migration.',
  'Do NOT license the target MailUser into a mailbox too early — if a target mailbox exists before the move, the migration fails.',
  'Teams meeting URLs in migrated calendar items are NOT updated — meetings must be recreated or links break.',
  'Outbox items do not migrate.',
  'Teams chat folder content does not migrate.',
  'Mailbox signatures (cloud signatures/roaming) do not migrate.',
  'Source and target cannot use the same custom domain at the same time — plan domain cutover separately.',
  'Cross-cloud tenant-to-tenant migration (e.g. Commercial ↔ GCC) is NOT supported.',
];

export const ctTargetSteps = [
  'Verify licensing: every migrating user needs a Cross Tenant User Data Migration license; assign Exchange license only at the right moment.',
  'Register the migration application in the TARGET tenant: Entra admin center > App registrations > New. Single tenant, no redirect URI needed.',
  'Add the Office 365 Exchange Online application permission "Mailbox.Migration" (application type) and grant admin consent in the target tenant.',
  'Create a client secret on the app registration. Store it safely (it is needed in the SOURCE tenant setup).',
  'Send the admin consent URL to the SOURCE tenant admin (generated below) so the app gets consented in the source tenant.',
  'In target Exchange Online: create the migration endpoint (New-MigrationEndpoint -RemoteServer outlook.office.com with the app credentials).',
  'Create/update the organization relationship to the source tenant with -MailboxMoveEnabled and -MailboxMoveCapability Inbound.',
  'Prepare MailUser objects for every migrating mailbox: correct ExchangeGUID, ArchiveGUID (if archive exists), X500 proxy from source LegacyExchangeDN, and targetAddress pointing at the source.',
  'Validate with Test-MigrationServerAvailability before creating any batch.',
];

export const ctSourceSteps = [
  'Open the admin consent URL from the target tenant admin and grant consent as Global Admin of the SOURCE tenant.',
  'Wait for the service principal to appear, then verify it under Enterprise applications.',
  'Create a mail-enabled security group containing ONLY the mailboxes that are allowed to migrate (migration scope).',
  'Create/update the organization relationship to the target tenant: -MailboxMoveEnabled, -MailboxMoveCapability RemoteOutbound, -OAuthApplicationId <appId>, -MailboxMovePublishedScopes <scope group>.',
  'Confirm migrating mailboxes are members of the scope group, are NOT on hold, and have the Cross Tenant User Data Migration license.',
  'Collect per-user: ExchangeGUID, ArchiveGUID, LegacyExchangeDN, proxyAddresses — the target needs these to prepare MailUsers.',
];

export const ctValidationChecklist = [
  'Test-MigrationServerAvailability succeeds from the target tenant',
  'Every target MailUser has ExchangeGUID matching the source mailbox',
  'ArchiveGUID stamped for users with an Online Archive',
  'X500 proxy address (source LegacyExchangeDN) present on target MailUser',
  'ExternalEmailAddress / targetAddress points at source mailbox',
  'Desired primary SMTP set on target (target delivery domain routable)',
  'Cross Tenant User Data Migration license assigned (source + target)',
  'No holds on source mailboxes',
  'Users are in the source scope security group',
  'No target mailbox exists yet for migrating users (MailUser only)',
  'Migration batch CSV uses header "EmailAddress" with TARGET MailUser identities',
];

export const ctPostMigration = [
  'Verify move requests completed: Get-MoveRequest -Flags CrossTenant / Get-MigrationUser statistics',
  'Source mailbox converted to MailUser pointing at target — verify mail routing both ways',
  'User rebuilds the Outlook profile (old profile points at the source mailbox)',
  'Mobile devices: remove and re-add the account',
  'Inform users: Teams meeting links in old invites are broken — recreate recurring meetings',
  'Re-set signatures, delegates and shared mailbox permissions in the target (permissions do NOT migrate cross-tenant)',
  'Re-assign licenses in target as planned; remove migration licenses when done',
  'Cleanup after the project: remove migration endpoint, organization relationships, scope group, and the app registration + secret',
];

export const ctRollbackNotes = [
  'There is NO automatic rollback after a completed cross-tenant move — the source mailbox is converted.',
  'Before cutover, rollback = simply remove the migration batch before completion (mailbox stays in source).',
  'After completion, "rollback" means migrating BACK with a new project in the opposite direction — plan accordingly.',
  'Always pilot with test mailboxes first and keep the pilot evidence in the project file.',
];

export function ctAdminConsentUrl(sourceOnMicrosoft: string, appClientId: string): string {
  return `https://login.microsoftonline.com/${sourceOnMicrosoft || '<source-tenant>'}/adminconsent?client_id=${appClientId || '<application-client-id>'}&redirect_uri=https://office.com`;
}

export function ctScripts(f: CTForm) {
  const vars = `# ===== Variables =====
$SourceTenantId   = "${f.sourceTenantId || '<source-tenant-id>'}"
$TargetTenantId   = "${f.targetTenantId || '<target-tenant-id>'}"
$SourceDomain     = "${f.sourceOnMicrosoft}"
$TargetDomain     = "${f.targetOnMicrosoft}"
$AppId            = "${f.appClientId || '<application-client-id>'}"
$AppSecret        = Read-Host "Client secret (${f.clientSecretName})" -AsSecureString
$EndpointName     = "${f.endpointName}"
$OrgRelName       = "${f.orgRelationshipName}"
$ScopeGroup       = "${f.scopeGroupName}"
$TargetDelivery   = "${f.targetDeliveryDomain}"`;

  const endpoint = `# ===== TARGET tenant: migration endpoint =====
Connect-ExchangeOnline # as target admin
New-MigrationEndpoint -RemoteServer outlook.office.com \`
  -RemoteTenant $SourceDomain \`
  -Credentials (New-Object System.Management.Automation.PSCredential($AppId, $AppSecret)) \`
  -ExchangeRemoteMove:$true -Name $EndpointName -ApplicationId $AppId`;

  const orgRelTarget = `# ===== TARGET tenant: organization relationship (inbound) =====
$sourceTenantGuid = $SourceTenantId
$existing = Get-OrganizationRelationship | Where-Object { $_.DomainNames -contains $sourceTenantGuid }
if ($existing) {
  Set-OrganizationRelationship $existing.Name -Enabled:$true -MailboxMoveEnabled:$true -MailboxMoveCapability Inbound
} else {
  New-OrganizationRelationship $OrgRelName -Enabled:$true -MailboxMoveEnabled:$true \`
    -MailboxMoveCapability Inbound -DomainNames $sourceTenantGuid
}`;

  const orgRelSource = `# ===== SOURCE tenant: organization relationship (outbound) =====
Connect-ExchangeOnline # as source admin
$targetTenantGuid = "${f.targetTenantId || '<target-tenant-id>'}"
$appId = "${f.appClientId || '<application-client-id>'}"
$scope = Get-DistributionGroup "${f.scopeGroupName}" -ErrorAction Stop
$existing = Get-OrganizationRelationship | Where-Object { $_.DomainNames -contains $targetTenantGuid }
if ($existing) {
  Set-OrganizationRelationship $existing.Name -Enabled:$true -MailboxMoveEnabled:$true \`
    -MailboxMoveCapability RemoteOutbound -OAuthApplicationId $appId \`
    -MailboxMovePublishedScopes $scope.Name
} else {
  New-OrganizationRelationship "${f.orgRelationshipName}" -Enabled:$true -MailboxMoveEnabled:$true \`
    -MailboxMoveCapability RemoteOutbound -DomainNames $targetTenantGuid \`
    -OAuthApplicationId $appId -MailboxMovePublishedScopes $scope.Name
}`;

  const mailUserPrep = `# ===== TARGET tenant: MailUser preparation (per user) =====
# Values from SOURCE: Get-Mailbox user | fl ExchangeGuid,ArchiveGuid,LegacyExchangeDN,EmailAddresses
$upn = "user@$TargetDomain"
Set-MailUser $upn -ExchangeGuid "<source-ExchangeGuid>"
Set-MailUser $upn -ArchiveGuid "<source-ArchiveGuid>"   # only if source has an archive
Set-MailUser $upn -EmailAddresses @{Add="X500:<source-LegacyExchangeDN>"}
Set-MailUser $upn -ExternalEmailAddress "user@$SourceDomain"  # targetAddress -> source
# Verify:
Get-MailUser $upn | Format-List ExchangeGuid,ArchiveGuid,EmailAddresses,ExternalEmailAddress,PrimarySmtpAddress`;

  const test = `# ===== TARGET tenant: validate before batch =====
Test-MigrationServerAvailability -Endpoint $EndpointName \`
  -TestMailbox "user@$TargetDomain"`;

  const batch = `# ===== TARGET tenant: create migration batch =====
# CSV: header "EmailAddress", one TARGET MailUser per line
New-MigrationBatch -Name "CT-Batch-01" \`
  -SourceEndpoint $EndpointName \`
  -CSVData ([System.IO.File]::ReadAllBytes("C:\\Migrations\\batch01.csv")) \`
  -TargetDeliveryDomain $TargetDelivery \`
  -AutoStart
# Monitor:
Get-MigrationBatch "CT-Batch-01" | Format-List Status,TotalCount,SyncedCount,FailedCount
Get-MigrationUser -BatchId "CT-Batch-01" | Get-MigrationUserStatistics | Select Identity,Status,Error
Get-MoveRequest -Flags CrossTenant | Get-MoveRequestStatistics | Select DisplayName,StatusDetail,PercentComplete
# Complete (cutover):
Complete-MigrationBatch "CT-Batch-01"`;

  const csv = `EmailAddress\n${(f.csvUsers || '').split('\n').filter((l) => l && !/^EmailAddress$/i.test(l.trim())).join('\n') || 'user1@' + f.targetOnMicrosoft}`;

  return { vars, endpoint, orgRelTarget, orgRelSource, mailUserPrep, test, batch, csv };
}

export const ctErrors: MigrationError[] = [
  {
    id: 'ct-license',
    name: 'CrossTenantMigrationWithoutLicensePermanentException',
    meaning: 'The user being migrated does not have the required Cross Tenant User Data Migration license.',
    cause: 'License not assigned (or not assigned in BOTH tenants where required) before the move was attempted.',
    fix: ['Assign the "Cross Tenant User Data Migration" add-on license to the user (source and target as required).', 'Wait for license provisioning (up to 24h, usually faster).', 'Restart the migration user: Remove-MigrationUser then re-add, or Resume after fix.'],
    powershell: 'Get-MgUserLicenseDetail -UserId user@contoso.com | Select-Object SkuPartNumber',
    prevention: 'Add license assignment to the pre-migration checklist and verify per batch before starting.',
    escalationNote: 'License purchasing/assignment blocked — needs customer licensing owner.',
  },
  {
    id: 'ct-scope',
    name: 'MailboxNotInCrossTenantMigrationScopeException',
    meaning: 'The source mailbox is not inside the published migration scope of the organization relationship.',
    cause: 'User is not a member of the mail-enabled security group set in MailboxMovePublishedScopes, or the scope group is wrong on the org relationship.',
    fix: ['Add the user to the scope security group in the SOURCE tenant.', 'Verify: Get-OrganizationRelationship | fl MailboxMovePublishedScopes.', 'Allow group membership to propagate (can take a few hours), then retry the migration user.'],
    powershell: "Add-DistributionGroupMember 'CrossTenantMigrationScope' -Member user@source.com\nGet-OrganizationRelationship | Format-List Name,MailboxMovePublishedScopes,MailboxMoveEnabled",
    prevention: 'Populate and verify the scope group as part of source preparation; export membership and diff against the batch CSV.',
    escalationNote: 'Scope group propagation exceeding 24h — raise Microsoft case.',
  },
  {
    id: 'ct-auxarchive-target',
    name: 'AuxArchiveNotFoundInTargetRecipientException',
    meaning: 'The source mailbox has auxiliary (expanded) archives, but the target recipient is missing the matching aux archive configuration.',
    cause: 'Source archive expanded into auxiliary archives (auto-expanding archive) which the target MailUser is not provisioned for.',
    fix: ['Verify source archive state: Get-Mailbox user | fl ArchiveGuid,AutoExpandingArchiveEnabled.', 'Aux/auto-expanded archives are a known hard blocker for cross-tenant moves — check current Microsoft guidance.', 'Open a Microsoft support case; aux archive scenarios often need backend handling.'],
    prevention: 'Identify auto-expanding archives during discovery and plan those users separately (e.g. third-party tool for archive content).',
    escalationNote: 'Aux archive blocker — Microsoft case required; consider alternative migration path for this user.',
  },
  {
    id: 'ct-notexpecteddb',
    name: 'MailboxIsNotInExpectedDBException',
    meaning: 'Exchange backend found the mailbox in a different database than expected — a transient service-side state mismatch.',
    cause: 'Recent mailbox move/maintenance inside the source service, stale directory data.',
    fix: ['Wait several hours and retry (Resume the migration user or remove and re-add).', 'Verify no other move request is active on the mailbox.', 'If persistent over 24-48h, open a Microsoft case.'],
    powershell: 'Get-MoveRequest -Identity user@source.com | Get-MoveRequestStatistics | Format-List StatusDetail,Message',
    prevention: 'Avoid scheduling migrations right after large source-side changes; retry logic in runbook.',
    escalationNote: 'Persistent backend DB mismatch — Microsoft support with move request statistics output.',
  },
  {
    id: 'ct-notaccepteddomain',
    name: 'NotAcceptedDomainException',
    meaning: 'An SMTP address on the object uses a domain that is not an accepted domain in the target tenant.',
    cause: 'Target MailUser still carries source-only domains in proxyAddresses, or TargetDeliveryDomain is not an accepted domain.',
    fix: ['List the target MailUser proxyAddresses and remove addresses with non-accepted domains.', 'Confirm TargetDeliveryDomain is an accepted domain in target: Get-AcceptedDomain.', 'Retry the migration user after cleanup.'],
    powershell: "Get-MailUser user@target.com | Select-Object -ExpandProperty EmailAddresses\nSet-MailUser user@target.com -EmailAddresses @{Remove='smtp:user@notaccepted.com'}",
    prevention: 'During MailUser prep, only stamp addresses with target-accepted domains plus the required X500.',
    escalationNote: 'Domain strategy conflict (domain still bound to source) — needs project-level domain cutover decision.',
  },
  {
    id: 'ct-sourceaux',
    name: 'SourceAuxArchiveIsProvisionedDuringCrossTenantMovePermanentException',
    meaning: 'During the move, the source archive expanded (aux archive got provisioned), invalidating the move.',
    cause: 'Auto-expanding archive grew past threshold while migration was in progress.',
    fix: ['Remove the failed migration user.', 'Re-evaluate the user: archive now has aux archives — see AuxArchive guidance (often Microsoft case / alternate tooling).', 'For other users: complete their moves before long delays allow expansion.'],
    prevention: 'Migrate large-archive users early; avoid long-running synced-but-not-completed states for archive-heavy users.',
    escalationNote: 'Archive expanded mid-flight — replan this user with project lead.',
  },
  {
    id: 'ct-duplicate-batch',
    name: 'UserDuplicateInOtherBatchException',
    meaning: 'The user already exists in another migration batch.',
    cause: 'User listed in two CSVs, or an old batch was never cleaned up.',
    fix: ['Find the existing entry: Get-MigrationUser user@target.com | fl BatchId,Status.', 'Remove the user from the obsolete batch: Remove-MigrationUser user@target.com.', 'Re-add to the correct batch.'],
    powershell: 'Get-MigrationUser -Identity user@target.com | Format-List Identity,BatchId,Status\nRemove-MigrationUser -Identity user@target.com',
    prevention: 'Deduplicate CSVs before batch creation; always clean completed/failed batches.',
    escalationNote: 'None usually needed — operational cleanup.',
  },
  {
    id: 'ct-missing-guid',
    name: 'MissingExchangeGuidException',
    meaning: 'The target MailUser has no (or an empty) ExchangeGUID, so Exchange cannot match it to the source mailbox.',
    cause: 'MailUser preparation step skipped or GUID stamped on the wrong object.',
    fix: ['Get the source value: Get-Mailbox user (source) | fl ExchangeGuid.', 'Stamp it on the target: Set-MailUser user@target.com -ExchangeGuid <guid>.', 'Verify, then restart the migration user.'],
    powershell: 'Set-MailUser user@target.com -ExchangeGuid "<source-ExchangeGuid>"\nGet-MailUser user@target.com | Select-Object ExchangeGuid',
    prevention: 'Make ExchangeGUID verification a hard gate in the validation checklist (script the comparison).',
    escalationNote: 'GUID cannot be set (object type conflict) — check for prematurely licensed mailbox; may need object rebuild.',
  },
  {
    id: 'ct-already-moving',
    name: 'SourceMailboxAlreadyBeingMovedPermanentException',
    meaning: 'A move request already exists for the source mailbox.',
    cause: 'Previous failed/abandoned batch left an active move request, or two batches target the same user.',
    fix: ['Inspect: Get-MoveRequest -Identity user (in the relevant tenant) | fl Status.', 'Remove the stale request: Remove-MoveRequest user (after confirming it is obsolete!).', 'Retry the new migration user after cleanup.'],
    powershell: 'Get-MoveRequest -Identity user@source.com | Format-List Status,TargetDeliveryDomain\nRemove-MoveRequest -Identity user@source.com -Confirm:$false',
    prevention: 'Always remove failed batches completely before recreating; one batch per user at a time.',
    escalationNote: 'Move request cannot be removed — Microsoft case.',
  },
  {
    id: 'ct-demoted-archive',
    name: 'UserAlreadyHasDemotedArchiveException',
    meaning: 'The target object carries remnants of a previous archive (demoted archive state), conflicting with archive provisioning during the move.',
    cause: 'Target user previously had an archive that was disabled/demoted; stale archive state on the object.',
    fix: ['Check target archive state: Get-MailUser user | fl ArchiveGuid,ArchiveState,DisabledArchiveGuid.', 'If DisabledArchiveGuid is set: this stale state must be cleared — commonly requires Microsoft support for MailUsers.', 'Alternative: recreate the target object cleanly (coordinate identity impact) and re-prep.'],
    powershell: 'Get-MailUser user@target.com | Format-List ArchiveGuid,DisabledArchiveGuid,ArchiveState',
    prevention: 'Use clean target objects; avoid enabling/disabling archives on MailUsers before migration.',
    escalationNote: 'Stale demoted archive on target — Microsoft case to clear backend state.',
  },
];
