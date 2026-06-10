import { MigrationError } from '../types';

export interface BTStep {
  id: string;
  title: string;
  details: string[];
}

export const btGuide: BTStep[] = [
  {
    id: 'bt-type',
    title: '1. Choose migration type',
    details: [
      'M365 to M365: project type "Mailbox" (plus separate projects for OneDrive/SharePoint/Teams via MigrationWiz Collaboration).',
      'Google Workspace to M365: mailbox project with G Suite (IMAP/API) source; plan Drive→OneDrive as separate document project.',
      'Exchange On-Premises to M365: mailbox project with Exchange source (Outlook Anywhere/EWS), or consider native hybrid instead for full coexistence.',
      'License per item: MigrationWiz licenses are per mailbox/user — size the purchase on the discovery numbers.',
    ],
  },
  {
    id: 'bt-perms',
    title: '2. Required admin permissions',
    details: [
      'M365 source/destination: use a dedicated migration admin account; modern auth app registration (MigrationWiz guides you) with ApplicationImpersonation (EWS) or Graph-based access depending on project type.',
      'Google Workspace: service account with domain-wide delegation, scopes for Gmail/Calendar/Contacts; admin user for the project.',
      'Exchange on-prem: account with ApplicationImpersonation role and EWS reachable from internet (or Outlook Anywhere).',
      'Document every account/app created — they must be removed at project closure.',
    ],
  },
  {
    id: 'bt-endpoints',
    title: '3. Create source & destination endpoints',
    details: [
      'MigrationWiz > Endpoints > Add: choose the correct endpoint type per environment.',
      'M365 endpoints: prefer "Office 365" with modern authentication (app registration + client secret/cert).',
      'Validate the endpoint immediately after creation — do not discover users with broken credentials.',
      'Use separate endpoints per tenant; never reuse a source endpoint as destination.',
    ],
  },
  {
    id: 'bt-import',
    title: '4. Import users (line items)',
    details: [
      'Autodiscover items via the endpoint, or import a CSV (Source Email, Destination Email columns).',
      'Verify the source→destination address mapping carefully — wrong mapping migrates data into the wrong mailbox.',
      'Remove out-of-scope items (service accounts, rooms handled natively) before licensing items.',
    ],
  },
  {
    id: 'bt-verify',
    title: '5. Verify credentials',
    details: [
      'Run "Verify Credentials" on a few items first, then all.',
      'Fix endpoint/permission errors NOW — a failed verification will fail the migration the same way.',
      'Typical failures: missing impersonation rights, MFA on basic-auth path, wrong service account scopes (Google).',
    ],
  },
  {
    id: 'bt-assessment',
    title: '6. Run assessment',
    details: [
      'Use the assessment/statistics run where available to size mailboxes and detect problem items.',
      'Flag mailboxes >50GB, heavy calendars and shared mailboxes for special attention.',
      'Output feeds the pilot selection and the cutover schedule.',
    ],
  },
  {
    id: 'bt-prestage',
    title: '7. Pre-stage migration',
    details: [
      'Start a Pre-Stage pass (e.g. items older than 90/30 days) days or weeks before cutover.',
      'Pre-stage moves the bulk without touching recent mail — users keep working in source.',
      'Monitor failures per item and resolve them before cutover week.',
    ],
  },
  {
    id: 'bt-full',
    title: '8. Full migration (cutover pass)',
    details: [
      'At cutover (after MX/identity switch per plan): run a Full Migration pass to copy remaining/recent items.',
      'Coordinate with the cutover checklist: MX, autodiscover, license/identity readiness in destination.',
      'Communicate the freeze window to users (changes in source after the pass may not arrive).',
    ],
  },
  {
    id: 'bt-delta',
    title: '9. Final delta pass',
    details: [
      'Within 72h after cutover, run a final delta pass to catch stragglers.',
      'Verify item counts source vs destination on sampled mailboxes.',
    ],
  },
  {
    id: 'bt-failed',
    title: '10. Review & retry failed items',
    details: [
      'Open each item\'s migration statistics: failed item count, error categories.',
      'Retry errors with "Retry Errors" — many EWS throttling errors clear on retry.',
      'Persistent failures: check the error helper below; document irrecoverable items (corrupt/oversized) for the customer.',
    ],
  },
  {
    id: 'bt-reports',
    title: '11. Export reports',
    details: [
      'Export per-item statistics and error reports from MigrationWiz for the project file.',
      'Customer-facing summary: migrated counts, failed item summary with explanation, open actions.',
    ],
  },
  {
    id: 'bt-validate',
    title: '12. Post-migration validation',
    details: [
      'Sample-check mailboxes: inbox, sent items, calendar (recurring meetings!), contacts.',
      'Verify delegates/shared mailbox permissions recreated in destination (MigrationWiz does not move all permissions).',
      'Confirm mobile devices and Outlook profiles reconfigured.',
      'Close: remove migration app registrations, admin accounts, and revoke secrets.',
    ],
  },
];

export const btChecklists = {
  permissions: [
    'Dedicated migration admin in source and destination',
    'Modern auth app registration completed (M365 endpoints)',
    'ApplicationImpersonation / service account scopes verified',
    'Test mailbox migrated successfully with these credentials',
  ],
  cutover: [
    'Pre-stage completed, failure list at acceptable level',
    'Destination licenses assigned and mailboxes provisioned',
    'MX record TTL lowered in advance',
    'MX + Autodiscover + SPF/DKIM switched per DNS plan',
    'Full pass started after MX switch',
    'Helpdesk briefed and staffed for cutover day',
    'User communication sent (what changes, new login, Outlook profile)',
  ],
  postMigration: [
    'Final delta pass completed',
    'Failed items reviewed, retried, residual documented',
    'Sampled mailbox content verification done',
    'Outlook/mobile profiles rebuilt',
    'Shared mailbox & delegate permissions recreated',
    'Reports exported to project file',
    'Migration accounts/apps/secrets removed',
  ],
};

export const btErrors: MigrationError[] = [
  {
    id: 'bt-insufficient-permissions',
    name: 'ErrorNonExistentMailbox / insufficient access rights',
    meaning: 'MigrationWiz cannot open the mailbox with the admin credentials.',
    cause: 'Missing ApplicationImpersonation (EWS) rights, wrong app permissions, or the mailbox address in the line item is wrong.',
    fix: ['Verify the item\'s source/destination address really exists.', 'Re-check impersonation role assignment / app permissions and re-verify credentials.', 'Retry the item after fixing.'],
    prevention: 'Verify credentials on ALL items before starting passes.',
    escalationNote: 'Permissions verified but access still failing — BitTitan support ticket with item ID.',
  },
  {
    id: 'bt-throttling',
    name: 'Connection did not succeed / throttling (HTTP 503, ErrorServerBusy)',
    meaning: 'Source or destination service is throttling the migration traffic.',
    cause: 'EWS/Graph throttling on tenant level, too many concurrent connections.',
    fix: ['Lower concurrent migrations in project Advanced Options.', 'For M365: request EWS throttling elevation via Microsoft (SfMC guidance) for the migration window.', 'Simply retry — throttling errors are transient.'],
    prevention: 'Schedule heavy passes off-hours; request throttling relaxation before big batches.',
    escalationNote: 'Persistent throttling despite mitigation — BitTitan support + Microsoft throttling request.',
  },
  {
    id: 'bt-oversized',
    name: 'Item exceeds maximum allowed size',
    meaning: 'An item (mail with attachments) is larger than the destination accepts (35/150MB limits).',
    cause: 'Destination message size limit below the item size.',
    fix: ['Raise destination max message size where possible (Set-Mailbox -MaxReceiveSize).', 'Otherwise document skipped items and deliver them separately (PST extract).'],
    powershell: 'Set-Mailbox user@contoso.com -MaxReceiveSize 150MB -MaxSendSize 150MB',
    prevention: 'Set destination size limits to maximum before migration passes.',
    escalationNote: 'Customer decision needed on irrecoverable oversized items.',
  },
  {
    id: 'bt-corrupt',
    name: 'Corrupted or unreadable item (ErrorCorruptData)',
    meaning: 'The source item cannot be read/converted.',
    cause: 'Corrupt MAPI properties, ancient items, malformed calendar exceptions.',
    fix: ['Retry once; if persistent, the item is genuinely corrupt.', 'Document item details (folder/subject/date) for the user; they can recreate manually if needed.'],
    prevention: 'Expectation management: a small residue of corrupt items is normal in old mailboxes.',
    escalationNote: 'High corrupt counts on one mailbox — investigate source mailbox health instead.',
  },
  {
    id: 'bt-auth-mfa',
    name: 'Authentication failed / 401 (basic auth blocked, MFA)',
    meaning: 'The endpoint credentials cannot authenticate.',
    cause: 'Basic authentication disabled (default in M365), MFA challenge on the admin account, expired secret.',
    fix: ['Switch endpoint to modern authentication (app registration).', 'Renew expired client secrets.', 'Use an account exempt from interactive MFA per policy ONLY if modern auth app flow is impossible (avoid).'],
    prevention: 'Always build endpoints with modern auth; calendar a reminder for secret expiry.',
    escalationNote: 'CA policy blocks the app — security team to scope a policy for the migration app.',
  },
  {
    id: 'bt-quota',
    name: 'Destination mailbox quota exceeded',
    meaning: 'Destination mailbox is full during migration.',
    cause: 'Destination quota smaller than source data; archive not enabled.',
    fix: ['Check destination quota and usage.', 'Enable Online Archive / assign bigger license, or filter what migrates (date filters).', 'Resume/retry the item afterwards.'],
    powershell: 'Get-MailboxStatistics user@contoso.com | Select-Object TotalItemSize,StorageLimitStatus',
    prevention: 'Compare source sizes vs destination quotas in the assessment stage.',
    escalationNote: 'License upgrade decision — customer.',
  },
  {
    id: 'bt-google-scope',
    name: 'Google: Insufficient OAuth scopes / 403 forbidden',
    meaning: 'The Google service account lacks the required API scopes for Gmail/Calendar/Contacts.',
    cause: 'Domain-wide delegation scopes not (correctly) added in Google Admin, API not enabled in the Google Cloud project.',
    fix: ['Re-add the exact scope list from BitTitan documentation under domain-wide delegation (client ID of the service account).', 'Enable Gmail API/Calendar API in the Google Cloud project.', 'Re-verify credentials.'],
    prevention: 'Copy-paste the scope list from BitTitan docs — typos are the #1 cause.',
    escalationNote: 'Google-side admin unavailable — customer Google admin must perform delegation.',
  },
  {
    id: 'bt-folder-mapping',
    name: 'Folder mapping / language folder issues',
    meaning: 'Items land in duplicated or wrongly named folders (e.g. "Inbox1", language mismatches).',
    cause: 'Localized default folder names in source, or repeated passes after folder renames.',
    fix: ['Use Advanced Options folder mapping (e.g. "FolderMapping=^INBOX->Inbox").', 'Avoid renaming default folders between passes.'],
    prevention: 'Pilot mailboxes with each source language before bulk passes.',
    escalationNote: 'Complex mapping needs — BitTitan support can advise project settings.',
  },
];
