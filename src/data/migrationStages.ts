import { MigrationStage } from '../types';
import { uid } from '../store/useLocalStorage';

const stageTemplates: { name: string; tasks: string[] }[] = [
  {
    name: 'Discovery',
    tasks: [
      'Inventory users, mailboxes (sizes!), shared mailboxes, groups, Teams, SharePoint sites',
      'Identify archives, holds, journaling, public folders',
      'List domains, DNS hosting, MX/SPF/DKIM/DMARC current state',
      'Identify third-party integrations (CRM, scan-to-mail, apps using SMTP/Graph)',
      'Identify compliance requirements (retention, holds, data residency)',
    ],
  },
  {
    name: 'Planning',
    tasks: [
      'Choose migration method/tool and write the migration plan',
      'Define batches/waves and pilot group',
      'Set cutover date and communication plan',
      'Define rollback strategy and acceptance criteria',
      'Get customer sign-off on the plan',
    ],
  },
  {
    name: 'Source preparation',
    tasks: [
      'Create migration admin/app credentials in source',
      'Remove/document holds and journaling where they block migration',
      'Clean up: disable leavers, empty obsolete mailboxes out of scope',
      'Export permissions (delegates, send-as, shared mailbox access)',
    ],
  },
  {
    name: 'Destination preparation',
    tasks: [
      'Tenant baseline: org settings, security defaults/CA plan',
      'Create migration admin/app in destination',
      'Set mailbox size limits and quotas to maximum where applicable',
      'Prepare Teams/SharePoint structure if in scope',
    ],
  },
  {
    name: 'Identity preparation',
    tasks: [
      'Decide identity model (cloud-only, Entra Connect, cutover of sync)',
      'Provision users/MailUsers in destination',
      'Map UPNs and primary SMTP addresses',
      'Plan MFA registration for all users in destination',
    ],
  },
  {
    name: 'Domain and DNS preparation',
    tasks: [
      'Lower TTL on MX and autodiscover records ahead of cutover',
      'Pre-create destination DNS records (verification, SPF draft, DKIM CNAMEs)',
      'Plan domain release/verification timing (esp. tenant-to-tenant: domain can exist in one tenant only)',
      'Document full DNS cutover runbook with rollback values',
    ],
  },
  {
    name: 'License preparation',
    tasks: [
      'Procure destination licenses (incl. Cross Tenant User Data Migration licenses if applicable)',
      'Map license SKUs per user group',
      'Assign licenses per plan timing (careful with cross-tenant: not too early on MailUsers)',
    ],
  },
  {
    name: 'BitTitan setup',
    tasks: [
      'Create MigrationWiz project(s) and endpoints',
      'Import and map users; verify credentials on all items',
      'Purchase/allocate MigrationWiz licenses',
      'Configure advanced options (folder mapping, filters, concurrency)',
    ],
  },
  {
    name: 'Pilot migration',
    tasks: [
      'Migrate pilot mailboxes (IT + representative users)',
      'Validate content, calendar, delegates, mobile, Outlook profiles',
      'Capture issues and update the runbook',
      'Pilot sign-off by customer',
    ],
  },
  {
    name: 'Pre-stage migration',
    tasks: [
      'Run pre-stage passes for all batches (older items)',
      'Track failed items and resolve per batch',
      'Confirm batch completion status before cutover week',
    ],
  },
  {
    name: 'Cutover',
    tasks: [
      'Execute DNS cutover runbook (MX, autodiscover, SPF/DKIM/DMARC)',
      'Switch identity/licenses per plan',
      'Run full/cutover migration pass',
      'Verify inbound/outbound mail flow in destination',
      'Helpdesk on standby; monitor closely',
    ],
  },
  {
    name: 'Final delta',
    tasks: [
      'Run final delta pass within 72h',
      'Compare item counts on sample mailboxes',
      'Close out residual failed items with documentation',
    ],
  },
  {
    name: 'Post-migration validation',
    tasks: [
      'Validate permissions recreated (shared mailboxes, delegates, send-as)',
      'Validate Teams/SharePoint content if in scope',
      'Mail flow + transport rules + connectors validated',
      'Security baseline applied in destination (MFA, CA)',
    ],
  },
  {
    name: 'User support',
    tasks: [
      'Hypercare period with prioritized migration tickets',
      'Outlook profile/mobile reconfiguration support',
      'Communicate known limitations (Teams meeting links, signatures)',
    ],
  },
  {
    name: 'Closure',
    tasks: [
      'Remove migration accounts, app registrations, secrets and endpoints',
      'Export final reports; deliver project file to customer',
      'Lessons learned + handover to regular support',
      'Decommission source per plan (after agreed retention window)',
    ],
  },
];

export function buildStages(): MigrationStage[] {
  return stageTemplates.map((s) => ({
    id: uid(),
    name: s.name,
    status: 'not-started',
    owner: '',
    risk: 'medium',
    notes: '',
    dueDate: '',
    tasks: s.tasks.map((t) => ({ id: uid(), label: t, done: false })),
  }));
}

export const migrationTypeLabels: Record<string, string> = {
  'm365-to-m365': 'Microsoft 365 → Microsoft 365',
  'google-to-m365': 'Google Workspace → Microsoft 365',
  'exchange-onprem-to-m365': 'Exchange On-Premises → Microsoft 365',
};
