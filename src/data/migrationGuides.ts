// Manual migration runbooks shown in the Migration module per project type.

export interface GuideSection {
  title: string;
  steps: string[];
}

export const t2tManualGuide: GuideSection[] = [
  {
    title: 'Phase 1 — Preparation (both tenants)',
    steps: [
      'Inventory the source tenant: users, mailbox sizes, shared mailboxes, groups, domains, DNS records (export everything).',
      'Buy/verify enough licenses in the TARGET tenant for all migrating users.',
      'Create admin accounts in both tenants for the migration (dedicated, MFA-protected).',
      'Lower the TTL of the MX and Autodiscover DNS records to 5 minutes at least 24-48h before cutover.',
      'Communicate the migration plan and downtime window to end users.',
      'Choose the data method: native cross-tenant mailbox migration (license required), third-party tool (BitTitan), or manual PST as last resort.',
    ],
  },
  {
    title: 'Phase 2 — Target tenant preparation',
    steps: [
      'Create the user accounts in the target tenant (initially on target.onmicrosoft.com UPNs).',
      'Assign licenses so mailboxes provision (for cross-tenant native: NOT before ExchangeGUID prep — see Cross-Tenant module).',
      'Recreate groups, shared mailboxes and resource mailboxes.',
      'Pre-stage data with your chosen tool while users still work in source (pre-stage pass).',
    ],
  },
  {
    title: 'Phase 3 — Domain move (the critical path)',
    steps: [
      'A custom domain can only exist in ONE tenant at a time.',
      'In the SOURCE tenant: change all UPNs and email addresses (primary + aliases) from the custom domain to source.onmicrosoft.com.',
      'Remove every reference to the domain: UPNs, proxyAddresses on mailboxes/groups/public folders, application IDs.',
      'Remove the domain from the source tenant (M365 admin center > Settings > Domains > Remove).',
      'Add and verify the domain in the TARGET tenant (TXT verification record).',
      'Switch UPNs and primary SMTP addresses in the target to the custom domain.',
      'Expect a propagation window — plan this inside the cutover weekend.',
    ],
  },
  {
    title: 'Phase 4 — DNS cutover',
    steps: [
      'Point the MX record to the target tenant value (from target admin center domain setup).',
      'Update Autodiscover CNAME to autodiscover.outlook.com (target).',
      'Update SPF to the target tenant include; configure DKIM CNAMEs in target and enable signing; keep DMARC consistent.',
      'Verify inbound and outbound mail flow in the target with test messages.',
    ],
  },
  {
    title: 'Phase 5 — Final data sync & post-migration',
    steps: [
      'Run the final delta pass of your migration tool after MX switch.',
      'Users: rebuild Outlook profiles, re-add mobile accounts, re-register MFA in target.',
      'Recreate what does not migrate: permissions (Full Access/Send As), signatures, inbox rules (verify), Teams meetings (links break).',
      'Validate: sampled content checks, shared mailbox access, calendars, mail flow.',
      'Keep the source tenant read-only/licensed for the agreed retention window, then decommission.',
    ],
  },
];

export const gmailToM365Guide: GuideSection[] = [
  {
    title: 'Phase 1 — Google Workspace preparation',
    steps: [
      'Sign in to the Google Admin console with a Super Admin account.',
      'Create a Google Cloud project and a service account; note the service account client ID; create a JSON key.',
      'Enable the required APIs in the project: Gmail API, Google Calendar API, People (Contacts) API.',
      'In Google Admin console > Security > API controls > Domain-wide delegation: add the service account client ID with the scopes required by Microsoft (mail, calendar, contacts readonly scopes from the official M365 Google migration doc).',
      'Create a migration "super user" admin in Google for the connection.',
    ],
  },
  {
    title: 'Phase 2 — Microsoft 365 preparation',
    steps: [
      'Verify your domain in Microsoft 365 (if the same domain is moving, plan the MX switch for cutover).',
      'Create users in M365 (sync or manual/CSV) and assign Exchange Online licenses — mailboxes must exist before migration.',
      'Optionally route mail during coexistence via a subdomain (e.g. m365.contoso.com) provisioned as accepted domain.',
      'In the new Exchange admin center: Migration > Add migration batch > "Migration to Exchange Online" > Google Workspace migration.',
      'Create the migration endpoint: upload the service account JSON key, enter the Google super admin email; EAC validates the connection.',
    ],
  },
  {
    title: 'Phase 3 — Run the migration batches',
    steps: [
      'Prepare a CSV: header "EmailAddress" with the M365 mailbox addresses (one per line).',
      'Create the batch with the CSV, select the endpoint and target delivery domain.',
      'Start with a pilot batch (IT + a few users); validate mail, calendar, contacts arrive correctly.',
      'Run production batches; Gmail labels become folders (an item with multiple labels lands in multiple folders — communicate this).',
      'Monitor: EAC migration dashboard, per-user sync status, failure reports.',
    ],
  },
  {
    title: 'Phase 4 — Cutover',
    steps: [
      'Lower MX TTL in advance; at cutover switch MX from Google (aspmx.l.google.com) to the M365 value.',
      'Update SPF to v=spf1 include:spf.protection.outlook.com -all (remove Google include after coexistence ends); enable DKIM in Defender portal; review DMARC.',
      'Complete the migration batches (final incremental sync runs automatically until completion).',
      'Decommission Google routing only after confirming all mail lands in M365.',
    ],
  },
  {
    title: 'Phase 5 — Beyond mail: Drive, Chat, and devices',
    steps: [
      'Google Drive → OneDrive/SharePoint: use Microsoft Migration Manager (SharePoint admin center > Migration > Google Workspace) or BitTitan/ShareGate.',
      'Recreate Google Groups as M365 Groups/distribution lists.',
      'Users set up Outlook profiles and mobile accounts; provide a "where is my stuff" guide (labels→folders mapping).',
      'Known gaps: Google Chat history, Sites, Forms do not migrate natively — set expectations.',
      'Keep Google licenses for the agreed fallback window, then downgrade/cancel.',
    ],
  },
];

export const exchangeOnPremGuide: GuideSection[] = [
  {
    title: 'Phase 1 — Assess & choose method',
    steps: [
      'Check Exchange Server version: 2016/2019 → hybrid (full/minimal); 2010/2013 → upgrade path or third-party tooling.',
      'Choose: Hybrid migration (best coexistence), Cutover (<150 mailboxes, Outlook Anywhere), Staged (legacy), or third-party (BitTitan).',
      'Verify prerequisites: Entra Connect for identity, valid third-party certificate, EWS/OWA published externally.',
    ],
  },
  {
    title: 'Phase 2 — Hybrid path (recommended)',
    steps: [
      'Install/verify Entra Connect (password hash sync recommended) and verify domains in M365.',
      'Run the Hybrid Configuration Wizard (HCW) on the Exchange server with Global Admin + on-prem Org Admin.',
      'HCW creates connectors, org relationship, OAuth and the migration endpoint automatically.',
      'Test free/busy and mail flow between on-prem and cloud test mailboxes.',
    ],
  },
  {
    title: 'Phase 3 — Migrate mailboxes',
    steps: [
      'EAC > Migration > Add batch > Remote move migration using the hybrid endpoint.',
      'Pilot batch first; then production waves; AutoSuspend before completion for controlled cutovers.',
      'Licenses must be assigned before completion (mailbox converts in cloud).',
      'Monitor with Get-MigrationBatch / Get-MoveRequestStatistics.',
    ],
  },
  {
    title: 'Phase 4 — Cutover & decommission',
    steps: [
      'Switch MX/SPF/Autodiscover to Exchange Online when all mailboxes are migrated (or keep centralized mail flow if required).',
      'Move mail-relay dependencies (apps/printers) to a connector-based or authenticated solution.',
      'Keep one Exchange server (or use modern management tools) for recipient management while Entra Connect syncs.',
      'Decommission remaining servers per Microsoft guidance after validation.',
    ],
  },
];
