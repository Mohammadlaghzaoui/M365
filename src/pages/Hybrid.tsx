import { useRef, useState } from 'react';
import { Server, PlayCircle, CheckCircle2, XCircle, AlertCircle, AlertTriangle } from 'lucide-react';
import { Button, Card, CodeBlock, Field, PageHeader, Section, TextArea } from '../components/ui';
import { useLocalStorage } from '../store/useLocalStorage';
import { AIHelper } from '../components/AIHelper';
import { PSConsole, buildConsoleScript, ConsoleLine } from '../components/PSConsole';

/**
 * Hybrid / On-Premises migration connector assistant.
 * Builds and validates the Exchange remote-move (hybrid) endpoint configuration:
 * MRS proxy on-prem, migration endpoint in EXO, and remote move batches.
 */

interface HybridForm {
  onPremFqdn: string;       // externally published EWS host, e.g. mail.contoso.com
  endpointName: string;
  adminAccount: string;     // DOMAIN\user or UPN with migration rights
  targetDeliveryDomain: string; // tenant.mail.onmicrosoft.com
  batchName: string;
  csvUsers: string;
}

const DEFAULTS: HybridForm = {
  onPremFqdn: 'mail.contoso.com',
  endpointName: 'Hybrid-OnPrem-Endpoint',
  adminAccount: 'CONTOSO\\migadmin',
  targetDeliveryDomain: 'contoso.mail.onmicrosoft.com',
  batchName: 'Hybrid-Wave-01',
  csvUsers: 'EmailAddress\nuser1@contoso.com\nuser2@contoso.com',
};

interface Check { name: string; status: 'passed' | 'failed' | 'warning'; detail: string; fix?: string }

function validate(f: HybridForm): Check[] {
  const checks: Check[] = [];
  const fqdnOk = /^[a-z0-9.-]+\.[a-z]{2,}$/i.test(f.onPremFqdn) && !/contoso\.com$/i.test(f.onPremFqdn);
  checks.push(fqdnOk
    ? { name: 'On-prem EWS FQDN', status: 'passed', detail: `${f.onPremFqdn} looks like a valid external hostname.` }
    : { name: 'On-prem EWS FQDN', status: 'failed', detail: `"${f.onPremFqdn}" is a placeholder or invalid — this must be the externally published EWS hostname (the name on the certificate).`, fix: 'Use the public DNS name that serves /EWS/mrsproxy.svc over 443.' });
  const tddOk = /\.mail\.onmicrosoft\.com$/i.test(f.targetDeliveryDomain);
  checks.push(tddOk
    ? { name: 'Target delivery domain', status: 'passed', detail: `${f.targetDeliveryDomain} is a hybrid routing domain.` }
    : { name: 'Target delivery domain', status: 'warning', detail: `"${f.targetDeliveryDomain}" is normally the tenant's <tenant>.mail.onmicrosoft.com hybrid routing domain.`, fix: 'Check accepted domains in EXO for the .mail.onmicrosoft.com entry.' });
  checks.push(f.adminAccount.includes('\\') || f.adminAccount.includes('@')
    ? { name: 'On-prem admin account', status: 'passed', detail: `${f.adminAccount} format OK (DOMAIN\\user or UPN).` }
    : { name: 'On-prem admin account', status: 'failed', detail: 'Provide on-prem credentials as DOMAIN\\user or UPN.', fix: 'The account needs migration rights (Recipient Management / Org Management).' });
  const users = f.csvUsers.split('\n').map((l) => l.trim()).filter((l) => l && !/^EmailAddress$/i.test(l));
  const dupes = users.filter((u, i) => users.findIndex((x) => x.toLowerCase() === u.toLowerCase()) !== i);
  checks.push(users.length
    ? { name: `Batch users (${users.length})`, status: dupes.length ? 'failed' : 'passed', detail: dupes.length ? `Duplicates found: ${[...new Set(dupes)].join(', ')}` : 'CSV rows valid (primary SMTP of on-prem mailboxes).', fix: dupes.length ? 'Remove duplicate rows.' : undefined }
    : { name: 'Batch users', status: 'failed', detail: 'CSV is empty.', fix: 'Add the on-prem mailbox addresses to migrate.' });
  return checks;
}

function scripts(f: HybridForm) {
  return [
    {
      title: '1. ON-PREM — enable MRS Proxy (one-time, Exchange Management Shell)',
      code: `Set-WebServicesVirtualDirectory -Identity "EWS (Default Web Site)" -MRSProxyEnabled $true\niisreset\n# Verify externally reachable (must return 401/200, NOT 403/404 and NOT a pre-auth page):\n# https://${f.onPremFqdn}/EWS/mrsproxy.svc`,
    },
    {
      title: '2. EXO — test connectivity to on-prem (run BEFORE creating anything)',
      code: `Connect-ExchangeOnline\n$cred = Get-Credential # ${f.adminAccount}\nTest-MigrationServerAvailability -ExchangeRemoteMove -RemoteServer ${f.onPremFqdn} -Credentials $cred`,
    },
    {
      title: '3. EXO — create the hybrid migration endpoint',
      code: `New-MigrationEndpoint -ExchangeRemoteMove -Name "${f.endpointName}" \`\n  -RemoteServer ${f.onPremFqdn} -Credentials $cred\n# Note: the Hybrid Configuration Wizard usually creates this — check first:\nGet-MigrationEndpoint | Format-Table Identity,EndpointType,RemoteServer`,
    },
    {
      title: '4. EXO — create the remote move batch (onboarding)',
      code: `New-MigrationBatch -Name "${f.batchName}" \`\n  -SourceEndpoint "${f.endpointName}" \`\n  -CSVData ([System.IO.File]::ReadAllBytes("C:\\Migrations\\${f.batchName}.csv")) \`\n  -TargetDeliveryDomain "${f.targetDeliveryDomain}" \`\n  -AutoStart\n# Controlled cutover: omit -AutoComplete; complete manually in the cutover window:\nComplete-MigrationBatch "${f.batchName}"`,
    },
    {
      title: '5. Monitor',
      code: `Get-MigrationBatch "${f.batchName}" | Format-List Status,TotalCount,SyncedCount,FailedCount\nGet-MigrationUser -BatchId "${f.batchName}" | Get-MigrationUserStatistics | Select-Object Identity,Status,PercentComplete,Error\nGet-MoveRequest | Get-MoveRequestStatistics | Select-Object DisplayName,StatusDetail,PercentComplete`,
    },
  ];
}

const prereqs = [
  'Entra Connect installed and healthy — users synced (hybrid identity is mandatory for remote moves).',
  'Hybrid Configuration Wizard (HCW) completed: connectors, org relationship, OAuth.',
  'Valid third-party certificate on the EWS-published hostname (no self-signed).',
  'EWS /mrsproxy.svc published externally on 443 WITHOUT pre-authentication (reverse proxy/WAF must pass it through).',
  'Exchange Online licenses assigned to migrating users BEFORE batch completion.',
  'Target delivery domain (<tenant>.mail.onmicrosoft.com) present as accepted domain & in on-prem email address policy.',
  'MX/Autodiscover cutover planned separately (or keep centralized mail flow).',
];

export default function Hybrid() {
  const [form, setForm] = useLocalStorage<HybridForm>('hybrid-form', DEFAULTS);
  const [checks, setChecks] = useState<Check[] | null>(null);
  const [consoleLines, setConsoleLines] = useState<ConsoleLine[] | null>(null);
  const pendingChecks = useRef<Check[] | null>(null);
  const set = <K extends keyof HybridForm>(k: K, v: string) => setForm((f) => ({ ...f, [k]: v }));
  const passed = checks?.every((c) => c.status !== 'failed') ?? false;

  const runValidation = () => {
    const result = validate(form);
    pendingChecks.current = result;
    setChecks(null);
    const failed = result.filter((c) => c.status === 'failed').length;
    setConsoleLines(buildConsoleScript({
      command: `Test-WorkPilotHybridEndpoint -RemoteServer ${form.onPremFqdn} -Endpoint "${form.endpointName}" -Mode Test`,
      connectLines: [
        `Resolving ${form.onPremFqdn} ...`,
        `Probing https://${form.onPremFqdn}/EWS/mrsproxy.svc (simulated — run Test-MigrationServerAvailability for the live check) ...`,
        `Loading endpoint configuration "${form.endpointName}" ...`,
        `Verifying target delivery domain ${form.targetDeliveryDomain} ...`,
      ],
      checks: result,
      summary: failed === 0
        ? `VALIDATION PASSED — configuration consistent. Runbook below is ready for the live Test-MigrationServerAvailability.`
        : `VALIDATION FAILED — ${failed} blocking issue(s). Fix and re-run.`,
      passed: failed === 0,
    }));
  };

  const icon = (s: string) =>
    s === 'passed' ? <CheckCircle2 size={16} className="shrink-0 text-emerald-500" />
    : s === 'warning' ? <AlertCircle size={16} className="shrink-0 text-amber-500" />
    : <XCircle size={16} className="shrink-0 text-red-500" />;

  return (
    <div>
      <PageHeader title="Hybrid / On-Prem Migration Connector" subtitle="Exchange on-premises → Exchange Online remote moves: MRS proxy setup, endpoint validation, and remote move batches for hybrid environments." icon={<Server size={20} />} />

      <Card className="p-5 mb-5">
        <Section title="Connector configuration">
          <div className="grid gap-3 md:grid-cols-3">
            <Field label="On-prem EWS FQDN (external)" value={form.onPremFqdn} onChange={(v) => set('onPremFqdn', v)} placeholder="mail.customer.com" />
            <Field label="Migration endpoint name" value={form.endpointName} onChange={(v) => set('endpointName', v)} />
            <Field label="On-prem admin (DOMAIN\\user)" value={form.adminAccount} onChange={(v) => set('adminAccount', v)} />
            <Field label="Target delivery domain" value={form.targetDeliveryDomain} onChange={(v) => set('targetDeliveryDomain', v)} placeholder="tenant.mail.onmicrosoft.com" />
            <Field label="Batch name" value={form.batchName} onChange={(v) => set('batchName', v)} />
          </div>
          <div className="mt-3">
            <TextArea label="User CSV (on-prem primary SMTP — header EmailAddress)" value={form.csvUsers} onChange={(v) => set('csvUsers', v)} rows={4} />
          </div>
          <div className="mt-4">
            <Button variant="ai" onClick={runValidation}><PlayCircle size={16} /> Validate configuration (test mode)</Button>
          </div>
        </Section>
      </Card>

      {consoleLines && (
        <div className="mb-5">
          <PSConsole lines={consoleLines} onDone={() => setChecks(pendingChecks.current)} title="Windows PowerShell — Hybrid Endpoint Validator" />
        </div>
      )}

      {checks && (
        <Card className={`p-5 mb-5 ${passed ? 'border-emerald-300 dark:border-emerald-700' : 'border-red-300 dark:border-red-700'}`}>
          <Section title={passed ? 'Validation passed — runbook below is ready' : 'Validation failed — fix the items below'}>
            <div className="space-y-2.5">
              {checks.map((c) => (
                <div key={c.name} className="flex items-start gap-2.5 text-sm">
                  {icon(c.status)}
                  <div>
                    <span className="font-medium text-slate-700 dark:text-slate-200">{c.name}</span>
                    <span className="text-slate-500 dark:text-slate-400"> — {c.detail}</span>
                    {c.fix && <div className="text-blue-600 dark:text-blue-400">Fix: {c.fix}</div>}
                  </div>
                </div>
              ))}
            </div>
          </Section>
        </Card>
      )}

      <Card className="p-5 mb-5 border-amber-200 dark:border-amber-800">
        <Section title="Prerequisites (hybrid environment)">
          <ul className="space-y-2">
            {prereqs.map((p, i) => (
              <li key={i} className="flex items-start gap-2 text-sm text-amber-700 dark:text-amber-300">
                <AlertTriangle size={15} className="mt-0.5 shrink-0" />{p}
              </li>
            ))}
          </ul>
        </Section>
      </Card>

      <div className="space-y-4">
        {scripts(form).map((s) => (
          <div key={s.title}>
            <div className="mb-1 text-sm font-semibold text-slate-600 dark:text-slate-300">{s.title}</div>
            <CodeBlock code={s.code} />
          </div>
        ))}
      </div>

      <div className="mt-6">
        <AIHelper
          context={`Hybrid Exchange remote move setup. On-prem EWS: ${form.onPremFqdn}, endpoint: ${form.endpointName}, TDD: ${form.targetDeliveryDomain}.`}
          defaultPrompt="Test-MigrationServerAvailability fails for my hybrid endpoint — walk me through diagnosing MRS proxy, certificate and publishing issues step by step."
        />
      </div>
    </div>
  );
}
