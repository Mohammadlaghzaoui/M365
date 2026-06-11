import { runPowerShell } from './powershell.js';
import { config } from './config.js';

/**
 * Exchange Online execution for the Migration Console.
 *
 * Auth: app-only via certificate (Connect-ExchangeOnline -AppId -Organization
 * -CertificateThumbprint) is the supported unattended method. Configure on the
 * agent host:
 *   EXO_APP_ID, EXO_ORG (tenant.onmicrosoft.com), EXO_CERT_THUMBPRINT
 * The certificate must be installed in the host's certificate store and the
 * app granted Exchange.ManageAsApp + the Exchange Administrator role.
 *
 * Falls back to an interactive Connect-ExchangeOnline if no app config is set
 * (useful when an engineer runs the agent locally and signs in once).
 */

export function exoConfigured() {
  return !!(config.exoAppId && config.exoOrg && config.exoCertThumbprint);
}

function connectBlock() {
  if (exoConfigured()) {
    return `Import-Module ExchangeOnlineManagement -ErrorAction Stop
Connect-ExchangeOnline -AppId '${config.exoAppId}' -Organization '${config.exoOrg}' -CertificateThumbprint '${config.exoCertThumbprint}' -ShowBanner:$false -ErrorAction Stop`;
  }
  return `Import-Module ExchangeOnlineManagement -ErrorAction Stop
if (-not (Get-ConnectionInformation -ErrorAction SilentlyContinue)) { Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop }`;
}

const esc = (s = '') => String(s).replace(/'/g, "''");

/** Build the New-MigrationBatch command from validated console parameters. */
export function buildMigrationBatchScript(p) {
  const csv = (p.users || []).map((u) => `${u.source}`).join('\n');
  const csvData = `EmailAddress\n${csv}`;
  return `${connectBlock()}
$ErrorActionPreference = 'Stop'
$csv = @"
${csvData}
"@
$bytes = [System.Text.Encoding]::UTF8.GetBytes($csv)
$existing = Get-MigrationBatch -Identity '${esc(p.batchName)}' -ErrorAction SilentlyContinue
if ($existing) { Write-Output 'Batch already exists — resuming.'; Start-MigrationBatch -Identity '${esc(p.batchName)}' -ErrorAction SilentlyContinue }
else {
  New-MigrationBatch -Name '${esc(p.batchName)}' -SourceEndpoint '${esc(p.endpointName)}' -CSVData $bytes -TargetDeliveryDomain '${esc(p.targetDeliveryDomain)}' -AutoStart -ErrorAction Stop | Out-Null
  Write-Output 'Batch created and started.'
}
Get-MigrationUser -BatchId '${esc(p.batchName)}' | Get-MigrationUserStatistics | Select-Object Identity,Status,PercentageComplete,SyncedItemCount,SkippedItemCount,Error | ConvertTo-Json -Depth 3 -Compress`;
}

/** Status poll for a running batch — returns parsed per-user statistics. */
export async function getBatchStatus(batchName, onData) {
  const script = `${connectBlock()}
Get-MigrationUser -BatchId '${esc(batchName)}' -ErrorAction Stop | Get-MigrationUserStatistics |
  Select-Object @{n='identity';e={$_.Identity.ToString()}},@{n='status';e={$_.Status.ToString()}},@{n='percent';e={$_.PercentageComplete}},@{n='synced';e={$_.SyncedItemCount}},@{n='skipped';e={$_.SkippedItemCount}},@{n='error';e={if($_.Error){$_.Error.ToString()}else{''}}} |
  ConvertTo-Json -Depth 3 -Compress`;
  const out = await runPowerShell(script, { onData });
  return parseJsonOutput(out);
}

/** Verify the cross-tenant/remote endpoint is reachable for a test mailbox. */
export async function testEndpoint(endpointName, testMailbox, onData) {
  const script = `${connectBlock()}
Test-MigrationServerAvailability -Endpoint '${esc(endpointName)}' ${testMailbox ? `-TestMailbox '${esc(testMailbox)}'` : ''} -ErrorAction Stop |
  Select-Object Result,Message | ConvertTo-Json -Compress`;
  const out = await runPowerShell(script, { onData });
  return parseJsonOutput(out);
}

function parseJsonOutput(out) {
  // PowerShell may emit log lines before the JSON; grab the last JSON token.
  const match = out.trim().match(/(\[.*\]|\{.*\})\s*$/s);
  if (!match) return [];
  try {
    const parsed = JSON.parse(match[1]);
    return Array.isArray(parsed) ? parsed : [parsed];
  } catch {
    return [];
  }
}
