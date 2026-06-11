import { randomUUID } from 'node:crypto';
import { runPowerShell } from './powershell.js';
import { graphRequest } from './graph.js';
import { config } from './config.js';
import { buildMigrationBatchScript, getBatchStatus, testEndpoint, exoConfigured, buildCrossTenantCheckScript } from './exchange.js';

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
    else if (job.type === 'migrate-test') await migrateTest(job);
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
  log(job, `Endpoint: ${p.endpointName} | Target delivery domain: ${p.targetDeliveryDomain}`, 'info');
  if (!exoConfigured()) {
    log(job, 'Exchange Online app-only auth is not configured (EXO_APP_ID/EXO_ORG/EXO_CERT_THUMBPRINT). Falling back to an interactive Connect-ExchangeOnline on the agent host.', 'warn');
  }
  // Prefer an operator-supplied validated script; otherwise build it from params.
  const script = p.script || buildMigrationBatchScript(p);
  log(job, 'Connecting to Exchange Online and submitting the batch ...', 'info');
  const out = await runPowerShell(script, { onData: (l) => log(job, l) });
  job.result = { status: 'Submitted', batchName: p.batchName, output: out.slice(-4000) };
  log(job, 'Migration batch submitted. Use migrate/status to poll Get-MigrationUserStatistics.', 'ok');
}

// ---- Test migration: validate endpoint + dry-run a small batch, surface errors ----
async function migrateTest(job) {
  const p = job.payload ?? {};
  if (p.crossTenant) {
    log(job, `CROSS-TENANT live validation for "${p.batchName}" — org relationship, endpoint, MailUser/ExchangeGuid per user ...`, 'info');
    const out = await runPowerShell(buildCrossTenantCheckScript(p), {
      onData: (l) => log(job, l, l.includes('[ FAIL ]') ? 'err' : l.includes('[ WARN ]') ? 'warn' : l.includes('[ PASS ]') ? 'ok' : 'info'),
    });
    const failed = (out.match(/\[ FAIL \]/g) || []).length;
    job.result = { status: failed ? 'TestFailed' : 'TestPassed', failures: failed };
    if (failed) throw new Error(`${failed} blocking issue(s) found in live cross-tenant validation.`);
    log(job, 'Live cross-tenant validation PASSED.', 'ok');
    return;
  }
  log(job, `TEST migration for "${p.batchName}" — validating before any real move ...`, 'info');
  log(job, `Step 1/2: Test-MigrationServerAvailability on endpoint "${p.endpointName}" ...`, 'info');
  const testMailbox = (p.users || [])[0]?.source;
  try {
    const avail = await testEndpoint(p.endpointName, testMailbox, (l) => log(job, l));
    const r = avail[0] ?? {};
    if (String(r.Result ?? r.result).toLowerCase().includes('success')) {
      log(job, `Endpoint reachable: ${r.Message ?? r.message ?? 'OK'}`, 'ok');
    } else {
      log(job, `Endpoint check returned: ${r.Result ?? ''} ${r.Message ?? r.message ?? ''}`, 'warn');
    }
  } catch (e) {
    log(job, `Endpoint validation FAILED: ${e.message}`, 'err');
    job.result = { status: 'TestFailed', stage: 'endpoint', error: String(e.message) };
    throw e;
  }
  log(job, `Step 2/2: validating ${(p.users || []).length} CSV row(s) and recipient readiness ...`, 'info');
  // Per-user recipient validation (cheap, read-only) to surface the common errors.
  const checkScript = buildRecipientCheckScript(p);
  const out = await runPowerShell(checkScript, { onData: (l) => log(job, l) });
  job.result = { status: 'TestPassed', detail: out.slice(-2000) };
  log(job, 'TEST migration validation finished — review any per-user warnings above.', 'ok');
}

function buildRecipientCheckScript(p) {
  const esc = (s = '') => String(s).replace(/'/g, "''");
  const lines = (p.users || []).map((u) => `Try {
  $r = Get-Recipient -Identity '${esc(u.destination || u.source)}' -ErrorAction Stop
  Write-Output ("OK  ${esc(u.source)} -> " + $r.RecipientTypeDetails)
} Catch {
  Write-Output ("ERR ${esc(u.source)} -> " + $_.Exception.Message)
}`).join('\n');
  return `Import-Module ExchangeOnlineManagement -ErrorAction Stop
if (-not (Get-ConnectionInformation -ErrorAction SilentlyContinue)) {
  ${exoConfigured() ? `Connect-ExchangeOnline -AppId '${esc(config.exoAppId)}' -Organization '${esc(config.exoOrg)}' -CertificateThumbprint '${esc(config.exoCertThumbprint)}' -ShowBanner:$false -ErrorAction Stop` : `Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop`}
}
${lines}`;
}

// ---- Raw operator PowerShell (guarded by ALLOW_RAW_POWERSHELL) ----
async function rawPowerShell(job) {
  const script = job.payload?.script;
  if (!script) throw new Error('No script provided.');
  log(job, 'Executing operator PowerShell ...', 'info');
  const out = await runPowerShell(script, { onData: (l) => log(job, l) });
  job.result = { output: out.slice(-4000) };
}
