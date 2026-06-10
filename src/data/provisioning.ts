import { AccountType, GroupCatalogueEntry, ProvisioningRequest } from '../types';

/**
 * Approved group catalogue (runbook section 7): the GUI never shows the full
 * directory — only catalogued groups with metadata that controls routing
 * (Graph vs AD), allowed account types, approval and expiry rules.
 * Manage entries in the module; stored in localStorage.
 */
export const seedCatalogue: GroupCatalogueEntry[] = [
  { id: '00000000-0000-0000-0000-000000000001', displayName: 'M365-E3-License', source: 'Cloud', allowedAccountTypes: ['Internal', 'ExternalMember'], approvalRequired: false, expiryRequired: false, description: 'License group — assigns Microsoft 365 E3 via group-based licensing.' },
  { id: '00000000-0000-0000-0000-000000000002', displayName: 'M365-F3-License', source: 'Cloud', allowedAccountTypes: ['Internal', 'ExternalMember'], approvalRequired: false, expiryRequired: false, description: 'License group — frontline F3.' },
  { id: 'CN=Application-Users,OU=Groups,DC=domain,DC=local', displayName: 'Application-Users', source: 'OnPrem', allowedAccountTypes: ['Internal', 'ExternalMember'], approvalRequired: false, expiryRequired: false, description: 'Core business application access (managed on-prem).' },
  { id: 'CN=VPN-Access,OU=Groups,DC=domain,DC=local', displayName: 'VPN-Access', source: 'Synced', allowedAccountTypes: ['Internal'], approvalRequired: true, expiryRequired: false, description: 'Remote access — approval required; membership managed on-prem.' },
  { id: '00000000-0000-0000-0000-000000000005', displayName: 'Teams-Standard-Users', source: 'Cloud', allowedAccountTypes: ['Internal', 'ExternalMember', 'Guest'], approvalRequired: false, expiryRequired: false, description: 'Standard Teams policies and channels.' },
  { id: '00000000-0000-0000-0000-000000000006', displayName: 'B2B-Partner-Collaboration', source: 'Cloud', allowedAccountTypes: ['Guest'], approvalRequired: false, expiryRequired: true, description: 'Approved guest collaboration spaces — expiry mandatory.' },
  { id: '00000000-0000-0000-0000-000000000007', displayName: 'Finance-Application-Access', source: 'Cloud', allowedAccountTypes: ['Internal'], approvalRequired: true, expiryRequired: true, description: 'Financial data application — approval + expiry enforced.' },
];

export interface ValidationIssue {
  field: string;
  message: string;
  blocking: boolean;
}

/** Field validation per runbook section 4 — strict before anything is sent. */
export function validateRequest(req: ProvisioningRequest, catalogue: GroupCatalogueEntry[]): ValidationIssue[] {
  const issues: ValidationIssue[] = [];
  const need = (cond: boolean, field: string, message: string, blocking = true) => {
    if (cond) issues.push({ field, message, blocking });
  };

  need(!req.requestedBy.trim(), 'Requester', 'Requester is required (audit trail).');
  need(!req.ticketNumber.trim(), 'Ticket number', 'Ticket/request number is required (requestId traceability).', false);
  need(!req.firstName.trim(), 'First name', 'First name is required.');
  need(!req.lastName.trim(), 'Last name', 'Last name is required.');
  need(!req.upnPrefix.trim() || /[^a-z0-9._-]/i.test(req.upnPrefix), 'UPN prefix', 'UPN prefix is required and may only contain letters, digits, dots, hyphens.');
  need(!/^[^@\s]+\.[^@\s]+$/.test(req.upnDomain), 'UPN domain', 'A valid UPN domain is required.');
  need(!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(req.mail), 'Mail', 'A valid mail address is required.');
  need((req.accountType === 'Internal' || req.accountType === 'ExternalMember') && !req.managerUpn.trim(),
    'Manager', 'Manager is mandatory for Internal and External member accounts (runbook rule).');
  need((req.accountType === 'Guest' || req.accountType === 'ExternalMember') && !req.endDate,
    'End date', 'End date is required for Guest and External member accounts (lifecycle policy).');
  if (req.startDate && req.endDate) {
    need(req.endDate <= req.startDate, 'End date', 'End date must be after the start date.');
  }
  need(req.accountType === 'Guest' && req.mail.split('@')[1]?.toLowerCase() === req.upnDomain.toLowerCase(),
    'Mail', 'A guest must be invited on their EXTERNAL email address, not a corporate one.');

  for (const gid of req.groups) {
    const g = catalogue.find((c) => c.id === gid);
    if (!g) {
      issues.push({ field: 'Groups', message: `Group ${gid} is not in the approved catalogue — free groupId input is not allowed.`, blocking: true });
      continue;
    }
    if (!g.allowedAccountTypes.includes(req.accountType)) {
      issues.push({ field: 'Groups', message: `"${g.displayName}" is not allowed for account type ${req.accountType}.`, blocking: true });
    }
    if (g.expiryRequired && !req.endDate) {
      issues.push({ field: 'Groups', message: `"${g.displayName}" enforces an expiry — set an end date.`, blocking: true });
    }
  }
  return issues;
}

export const accountTypeInfo: Record<AccountType, { label: string; route: string; rule: string }> = {
  Internal: {
    label: 'Internal employee',
    route: 'On-Premises AD first → Entra Connect sync → cloud groups after sync',
    rule: 'NEVER created cloud-only in a hybrid tenant — AD remains the source of authority.',
  },
  ExternalMember: {
    label: 'External member',
    route: 'Microsoft Graph POST /users (userType: Member) — or AD-first per policy',
    rule: 'Only when a managed member identity is wanted; groups via approved catalogue, end date mandatory.',
  },
  Guest: {
    label: 'Guest (B2B)',
    route: 'Microsoft Graph Invitation API (POST /invitations)',
    rule: 'Never created via POST /users — invite on the external address, then assign approved B2B groups.',
  },
};

export function buildRequestJson(req: ProvisioningRequest): string {
  return JSON.stringify({
    requestId: req.ticketNumber || req.requestId,
    requestedBy: req.requestedBy,
    accountType: req.accountType,
    firstName: req.firstName,
    lastName: req.lastName,
    displayName: req.displayName,
    userPrincipalName: `${req.upnPrefix}@${req.upnDomain}`,
    mail: req.mail,
    managerUpn: req.managerUpn,
    department: req.department,
    jobTitle: req.jobTitle,
    startDate: req.startDate,
    endDate: req.endDate,
    groups: req.groups,
  }, null, 2);
}
