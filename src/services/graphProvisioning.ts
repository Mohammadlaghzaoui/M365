import { getGraphToken } from './sso';
import { AccountType, GroupCatalogueEntry, ProvisioningRequest } from '../types';

/**
 * Microsoft Graph user provisioning (per the enterprise provisioning runbook).
 *
 * Security model: this module uses DELEGATED permissions via the signed-in
 * operator's Microsoft 365 SSO session (MSAL popup) — the frontend never holds
 * a client secret or application credential. For full production scale the
 * runbook architecture (API Management + Azure Functions + automation account
 * with managed identity) applies; this module implements the same contract so
 * the GUI can later point at that backend without changes.
 *
 * Required delegated scopes (admin consent in the app registration):
 *  - User.ReadWrite.All   (create member users, duplicate check)
 *  - User.Invite.All      (guest invitations)
 *  - Group.ReadWrite.All  (cloud group assignment)
 */

const GRAPH = 'https://graph.microsoft.com/v1.0';
export const PROVISIONING_SCOPES = ['User.ReadWrite.All', 'User.Invite.All', 'Group.ReadWrite.All'];

async function graph(path: string, init?: RequestInit & { scopes?: string[] }) {
  const token = await getGraphToken(init?.scopes ?? PROVISIONING_SCOPES);
  const res = await fetch(`${GRAPH}${path}`, {
    ...init,
    headers: { 'content-type': 'application/json', authorization: `Bearer ${token}`, ...(init?.headers ?? {}) },
  });
  if (!res.ok) {
    const body = await res.text();
    let message = body;
    try { message = JSON.parse(body)?.error?.message ?? body; } catch { /* keep raw */ }
    throw new Error(`Graph ${res.status}: ${message}`);
  }
  return res.status === 204 ? null : res.json();
}

const esc = (v: string) => v.replace(/'/g, "''");

export interface DuplicateMatch {
  matchedAttribute: string;
  matchedValue: string;
  objectId: string;
  displayName: string;
  userType?: string;
}

/**
 * Duplicate check per runbook section 5: UPN, mail, proxyAddresses and
 * otherMails in Entra ID (covers synced on-prem objects too), plus existing
 * guests on the invited address.
 */
export async function duplicateCheck(upn: string, mail: string): Promise<DuplicateMatch | null> {
  const filters = [
    `userPrincipalName eq '${esc(upn)}'`,
    `mail eq '${esc(mail)}'`,
    `proxyAddresses/any(p:p eq 'SMTP:${esc(mail)}')`,
    `otherMails/any(m:m eq '${esc(mail)}')`,
  ];
  for (const f of filters) {
    const data = await graph(`/users?$filter=${encodeURIComponent(f)}&$select=id,displayName,userPrincipalName,mail,userType&$top=1`);
    const hit = data?.value?.[0];
    if (hit) {
      return {
        matchedAttribute: f.split(' ')[0].split('/')[0],
        matchedValue: f.includes('userPrincipalName') ? upn : mail,
        objectId: hit.id,
        displayName: hit.displayName,
        userType: hit.userType,
      };
    }
  }
  return null;
}

export async function verifyManager(managerUpn: string): Promise<{ id: string; displayName: string }> {
  const data = await graph(`/users/${encodeURIComponent(managerUpn)}?$select=id,displayName`);
  return { id: data.id, displayName: data.displayName };
}

function tempPassword(): string {
  const bytes = crypto.getRandomValues(new Uint8Array(12));
  return 'Wp!' + Array.from(bytes).map((b) => 'abcdefghjkmnpqrstuvwxyzABCDEFGHJKMNPQRSTUVWXYZ23456789'[b % 54]).join('');
}

/** External member: Graph POST /users with userType Member (runbook 6.2). */
export async function createMemberUser(req: ProvisioningRequest): Promise<{ id: string; upn: string; initialPassword: string }> {
  const password = tempPassword();
  const upn = `${req.upnPrefix}@${req.upnDomain}`;
  const user = await graph('/users', {
    method: 'POST',
    body: JSON.stringify({
      accountEnabled: true,
      displayName: req.displayName,
      givenName: req.firstName,
      surname: req.lastName,
      mailNickname: req.upnPrefix,
      userPrincipalName: upn,
      mail: req.mail || undefined,
      userType: 'Member',
      department: req.department || undefined,
      jobTitle: req.jobTitle || undefined,
      employeeType: 'External member',
      passwordProfile: { password, forceChangePasswordNextSignIn: true },
    }),
  });
  if (req.managerUpn) {
    const mgr = await verifyManager(req.managerUpn);
    await graph(`/users/${user.id}/manager/$ref`, {
      method: 'PUT',
      body: JSON.stringify({ '@odata.id': `${GRAPH}/users/${mgr.id}` }),
    });
  }
  return { id: user.id, upn, initialPassword: password };
}

/** Guest: Graph Invitation API — never POST /users for guests (runbook 6.3). */
export async function inviteGuest(req: ProvisioningRequest, redirectUrl: string): Promise<{ id: string; inviteRedeemUrl: string }> {
  const data = await graph('/invitations', {
    method: 'POST',
    body: JSON.stringify({
      invitedUserEmailAddress: req.mail,
      invitedUserDisplayName: req.displayName,
      inviteRedirectUrl: redirectUrl || 'https://myapps.microsoft.com',
      sendInvitationMessage: true,
      invitedUserMessageInfo: { customizedMessageBody: `You have been invited by ${req.requestedBy} (request ${req.requestId}).` },
    }),
  });
  return { id: data.invitedUser?.id, inviteRedeemUrl: data.inviteRedeemUrl };
}

/** Cloud group assignment via the approved catalogue (runbook 7). */
export async function addToCloudGroup(userObjectId: string, groupId: string): Promise<void> {
  await graph(`/groups/${groupId}/members/$ref`, {
    method: 'POST',
    body: JSON.stringify({ '@odata.id': `${GRAPH}/directoryObjects/${userObjectId}` }),
  });
}

/** AD-first PowerShell for INTERNAL accounts — hybrid rule: never cloud-only (runbook 6.1). */
export function buildOnPremScript(req: ProvisioningRequest, groups: GroupCatalogueEntry[]): string {
  const upn = `${req.upnPrefix}@${req.upnDomain}`;
  const onPremGroups = groups.filter((g) => g.source !== 'Cloud');
  const cloudGroups = groups.filter((g) => g.source === 'Cloud');
  return `# ===== INTERNAL ACCOUNT — hybrid AD-first provisioning (request ${req.requestId}) =====
# Rule: internal identities are created in On-Premises AD and synced to Entra ID.
# Run in the on-prem AD PowerShell context with a delegated service account (no Domain Admin).

$ou = "OU=Internal Users,DC=yourdomain,DC=local"   # adjust to the internal OU per naming policy
$password = Read-Host "Initial password" -AsSecureString

New-ADUser \`
  -GivenName "${req.firstName}" \`
  -Surname "${req.lastName}" \`
  -Name "${req.displayName}" \`
  -DisplayName "${req.displayName}" \`
  -SamAccountName "${req.upnPrefix}" \`
  -UserPrincipalName "${upn}" \`
  -EmailAddress "${req.mail}" \`
  -Department "${req.department}" \`
  -Title "${req.jobTitle}" \`
  -Path $ou \`
  -AccountPassword $password \`
  -ChangePasswordAtLogon $true \`
  -Enabled $true

# Mail attributes per mail policy (hybrid Exchange):
Set-ADUser "${req.upnPrefix}" -Add @{proxyAddresses="SMTP:${req.mail}"; mailNickname="${req.upnPrefix}"}

# Manager:
${req.managerUpn ? `Set-ADUser "${req.upnPrefix}" -Manager (Get-ADUser -Filter "UserPrincipalName -eq '${req.managerUpn}'")` : '# (no manager supplied)'}

# On-prem / synced group memberships (managed on-prem per catalogue):
${onPremGroups.length ? onPremGroups.map((g) => `Add-ADGroupMember -Identity "${g.id}" -Members "${req.upnPrefix}"`).join('\n') : '# (none selected)'}

# Trigger sync (on the Entra Connect server):
Start-ADSyncSyncCycle -PolicyType Delta

# AFTER sync completes — cloud-only group/license assignment (run in Graph PowerShell):
${cloudGroups.length ? cloudGroups.map((g) => `# New-MgGroupMember -GroupId "${g.id}" -DirectoryObjectId (Get-MgUser -UserId "${upn}").Id  # ${g.displayName}`).join('\n') : '# (none selected)'}

# Status for the request log: PendingSync -> verify object in Entra ID, then mark Created.`;
}
