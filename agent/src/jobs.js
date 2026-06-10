import { randomUUID } from 'node:crypto';
import { runPowerShell } from './powershell.js';
import { graphRequest } from './graph.js';
import { config } from './config.js';

export const jobs = new Map();

export function createJob(type, payload) {
  const job = {
    id: randomUUID(),
    type,
    payload,
    status: 'queued',
    progress: 0,
    log: [],
    result: null,
    error: null,
    startedAt: new Date().toISOString(),
    finishedAt: null,
  };
  jobs.set(job.id, job);
  // Cap memory: keep last 200 jobs.
  if (jobs.size > 200) jobs.delete(jobs.keys().next().value);
  return job;
}

function log(job, text, level = 'info') {
  job.log.push({ t: new Date().toISOString(), level, text });
}

export async function runJob(job) {
  job.status = 'running';
  try {
    if (job.type === 'provision-user') await provisionUser(job);
    else if (job.type === 'migrate-batch') await migrateBatch(job);
    else if (job.type === 'powershell') await rawPowerShell(job);
    else throw new Error(`Unknown job type ${job.type}`);
    job.status = 'completed';
    job.progress = 100;
    log(job, 'Job completed.', 'ok');
  } catch (e) {
    job.status = 'failed';
    job.error = String(e.message ?? e);
    log(job, `Job failed: ${job.error}`, 'err');
  } finally {
    job.finishedAt = new Date().toISOString();
  }
}

// ---- Provision user: Internal=on-prem AD (PowerShell), else cloud (Graph) ----
async function provisionUser(job) {
  const p = job.payload ?? {};
  if (p.accountType === 'Internal') {
    log(job, `Creating on-prem AD user ${p.upn} in ${p.ou ?? '(default OU)'} ...`, 'info');
    const script = p.script || buildAdScript(p);
    await runPowerShell(script, { onData: (l) => log(job, l) });
    log(job, 'AD object created. Triggering Entra Connect delta sync ...', 'info');
    if (p.entraConnectServer) {
      await runPowerShell(`Invoke-Command -ComputerName ${p.entraConnectServer} -ScriptBlock { Start-ADSyncSyncCycle -PolicyType Delta }`, { onData: (l) => log(job, l) }).catch((e) => log(job, `Sync trigger warning: ${e.message}`, 'warn'));
    }
    job.result = { status: 'PendingSync', upn: p.upn };
  } else if (p.accountType === 'Guest') {
    log(job, `Sending B2B invitation to ${p.mail} via Graph ...`, 'info');
    const inv = await graphRequest('POST', '/invitations', {
      invitedUserEmailAddress: p.mail,
      invitedUserDisplayName: p.displayName,
      inviteRedirectUrl: p.redirectUrl || 'https://myapps.microsoft.com',
      sendInvitationMessage: true,
    });
    job.result = { status: 'Invited', objectId: inv.invitedUser?.id, redeemUrl: inv.inviteRedeemUrl };
    log(job, `Invitation sent (objectId ${inv.invitedUser?.id}).`, 'ok');
  } else {
    log(job, `Creating cloud member ${p.upn} via Graph ...`, 'info');
    const user = await graphRequest('POST', '/users', {
      accountEnabled: true,
      displayName: p.displayName,
      givenName: p.firstName,
      surname: p.lastName,
      mailNickname: p.upnPrefix,
      userPrincipalName: p.upn,
      userType: 'Member',
      passwordProfile: { password: p.password, forceChangePasswordNextSignIn: true },
    });
    job.result = { status: 'Created', objectId: user.id, upn: user.userPrincipalName };
    log(job, `User created (objectId ${user.id}).`, 'ok');
  }
}

function buildAdScript(p) {
  const ou = p.ou || 'OU=Users,DC=domain,DC=local';
  return [
    `$pw = ConvertTo-SecureString '${(p.password || 'TempP@ss!' + Math.random().toString(36).slice(2)).replace(/'/g, "''")}' -AsPlainText -Force`,
    `New-ADUser -GivenName '${p.firstName}' -Surname '${p.lastName}' -Name '${p.displayName}' -DisplayName '${p.displayName}' -SamAccountName '${p.upnPrefix}' -UserPrincipalName '${p.upn}' -EmailAddress '${p.mail}' -Path '${ou}' -AccountPassword $pw -ChangePasswordAtLogon $true -Enabled $true`,
    `Write-Output "Created ${p.upn}"`,
  ].join('\n');
}

// ---- Mailbox migration batch (Exchange Online remote/cross-tenant move) ----
async function migrateBatch(job) {
  const p = job.payload ?? {};
  log(job, `Preparing migration batch "${p.batchName}" (${(p.users || []).length} users) ...`, 'info');
  // The real command runs in an Exchange Online PowerShell session on the agent host.
  // We pass through the operator-validated New-MigrationBatch invocation.
  const script = p.script;
  if (!script) {
    log(job, 'No batch script supplied — agent expects the portal to pass the validated New-MigrationBatch command.', 'warn');
    throw new Error('Missing migration batch script.');
  }
  log(job, 'Connecting to Exchange Online and submitting the batch ...', 'info');
  const out = await runPowerShell(script, { onData: (l) => log(job, l) });
  job.result = { status: 'Submitted', output: out.slice(-2000) };
  log(job, 'Migration batch submitted. Poll Get-MigrationUserStatistics for progress.', 'ok');
}

// ---- Raw operator PowerShell (guarded by ALLOW_RAW_POWERSHELL) ----
async function rawPowerShell(job) {
  const script = job.payload?.script;
  if (!script) throw new Error('No script provided.');
  log(job, 'Executing operator PowerShell ...', 'info');
  const out = await runPowerShell(script, { onData: (l) => log(job, l) });
  job.result = { output: out.slice(-4000) };
}
