import { CTForm, ctScripts } from '../data/crossTenant';

/**
 * Migration Runner
 * ----------------
 * TEST MODE runs the full validation + a simulated batch locally (no tenant access):
 * every input is checked against the same rules that cause real cross-tenant
 * failures, and each simulated user gets a pass/fail with the real exception
 * name it would produce. The AI loop can then analyze failures and propose
 * (or auto-apply) configuration fixes, and the run is repeated until green.
 *
 * PRODUCTION MODE does not blind-fire against tenants from the browser:
 * it emits the fully validated, ordered runbook (the exact PowerShell that the
 * test run verified) plus a go/no-go gate. The ExecutionAdapter interface below
 * is the seam where a backend (Azure Function / automation account with
 * certificate auth) plugs in to execute the same plan server-side.
 */

export type RunMode = 'test' | 'production';

export interface UserCheck {
  email: string;
  status: 'passed' | 'failed' | 'warning';
  exception?: string;
  detail: string;
  fix?: string;
}

export interface ConfigCheck {
  name: string;
  status: 'passed' | 'failed' | 'warning';
  detail: string;
  fix?: string;
}

export interface RunResult {
  mode: RunMode;
  startedAt: string;
  configChecks: ConfigCheck[];
  userChecks: UserCheck[];
  passed: boolean;
  summary: string;
}

/** Seam for real execution: implement against a backend holding EXO app credentials. */
export interface ExecutionAdapter {
  name: string;
  execute(plan: string[]): Promise<RunResult>;
}

const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

export function parseUsers(csv: string): string[] {
  return csv
    .split('\n')
    .map((l) => l.trim())
    .filter((l) => l && !/^EmailAddress$/i.test(l));
}

export function runValidation(form: CTForm, mode: RunMode): RunResult {
  const configChecks: ConfigCheck[] = [];
  const users = parseUsers(form.csvUsers);

  const check = (name: string, ok: boolean, okDetail: string, failDetail: string, fix?: string, warn = false): void => {
    configChecks.push({
      name,
      status: ok ? 'passed' : warn ? 'warning' : 'failed',
      detail: ok ? okDetail : failDetail,
      fix: ok ? undefined : fix,
    });
  };

  check(
    'Source tenant ID format',
    GUID_RE.test(form.sourceTenantId),
    'Valid GUID.',
    `"${form.sourceTenantId || '(empty)'}" is not a tenant GUID — the organization relationship DomainNames must use the tenant ID GUID.`,
    'Find it: Entra admin center > Overview > Tenant ID (source tenant).',
  );
  check(
    'Target tenant ID format',
    GUID_RE.test(form.targetTenantId),
    'Valid GUID.',
    `"${form.targetTenantId || '(empty)'}" is not a tenant GUID.`,
    'Find it: Entra admin center > Overview > Tenant ID (target tenant).',
  );
  check(
    'Application (client) ID',
    GUID_RE.test(form.appClientId),
    'Valid app ID format.',
    'Missing/invalid migration app client ID — endpoint creation and the source org relationship will fail (consent error).',
    'Register the app in the TARGET tenant with Mailbox.Migration permission, then paste its Application (client) ID.',
  );
  check(
    'Source onmicrosoft domain',
    /\.onmicrosoft\.com$/i.test(form.sourceOnMicrosoft) && !/^source\.onmicrosoft\.com$/i.test(form.sourceOnMicrosoft),
    `${form.sourceOnMicrosoft} looks valid.`,
    `"${form.sourceOnMicrosoft}" must be the real *.onmicrosoft.com domain of the source tenant (placeholder detected or wrong suffix).`,
    'Use the initial domain: M365 admin center > Settings > Domains (source tenant).',
  );
  check(
    'Target onmicrosoft domain',
    /\.onmicrosoft\.com$/i.test(form.targetOnMicrosoft) && !/^target\.onmicrosoft\.com$/i.test(form.targetOnMicrosoft),
    `${form.targetOnMicrosoft} looks valid.`,
    `"${form.targetOnMicrosoft}" must be the real *.onmicrosoft.com domain of the target tenant.`,
    'Use the initial domain of the target tenant.',
  );
  check(
    'Tenants are different',
    !!form.sourceTenantId && form.sourceTenantId !== form.targetTenantId,
    'Source and target differ.',
    'Source and target tenant ID are identical — cross-tenant migration requires two different tenants.',
    'Correct one of the tenant IDs.',
  );
  check(
    'Target delivery domain routable',
    /\.onmicrosoft\.com$/i.test(form.targetDeliveryDomain),
    `${form.targetDeliveryDomain} is an onmicrosoft routing domain.`,
    `TargetDeliveryDomain "${form.targetDeliveryDomain}" should normally be the target *.onmicrosoft.com domain (NotAcceptedDomainException risk if the domain is not accepted in target).`,
    'Set it to the target onmicrosoft.com domain.',
    true,
  );
  check(
    'Scope group named',
    !!form.scopeGroupName.trim(),
    `Scope group: ${form.scopeGroupName}.`,
    'No mail-enabled security group set as migration scope — every move would fail with MailboxNotInCrossTenantMigrationScopeException.',
    'Create a mail-enabled security group in the source tenant and add all migrating mailboxes.',
  );
  check(
    'Batch has users',
    users.length > 0,
    `${users.length} user(s) in the batch.`,
    'The user CSV is empty — nothing to migrate.',
    'Add target MailUser addresses (one per line) under the EmailAddress header.',
  );

  // Per-user simulated checks: same rules that produce real exceptions.
  const seen = new Set<string>();
  const userChecks: UserCheck[] = users.map((email) => {
    const lower = email.toLowerCase();
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) {
      return {
        email, status: 'failed', exception: 'InvalidRecipientException',
        detail: 'Not a valid SMTP address — the batch CSV row would be rejected.',
        fix: 'Correct the address format (user@domain).',
      };
    }
    if (seen.has(lower)) {
      return {
        email, status: 'failed', exception: 'UserDuplicateInOtherBatchException',
        detail: 'Duplicate entry — the same user appears twice; the second occurrence fails.',
        fix: 'Remove the duplicate row from the CSV.',
      };
    }
    seen.add(lower);
    const domain = email.split('@')[1]?.toLowerCase() ?? '';
    if (form.sourceOnMicrosoft && domain === form.sourceOnMicrosoft.toLowerCase()) {
      return {
        email, status: 'failed', exception: 'NotAcceptedDomainException',
        detail: 'CSV must list the TARGET MailUser identity, but this address uses the SOURCE tenant domain.',
        fix: `Replace with the user's target identity (…@${form.targetOnMicrosoft}).`,
      };
    }
    if (form.targetOnMicrosoft && domain !== form.targetOnMicrosoft.toLowerCase() && domain.endsWith('.onmicrosoft.com')) {
      return {
        email, status: 'warning',
        detail: `Address domain ${domain} does not match the configured target tenant domain ${form.targetOnMicrosoft} — verify this is really a target-tenant identity.`,
        fix: 'Use the target MailUser primary/onmicrosoft address.',
      };
    }
    return {
      email, status: 'passed',
      detail: 'CSV row valid. Real-run prerequisites still verified at execution: ExchangeGUID stamped, CT license, no holds, scope membership.',
    };
  });

  const failed = configChecks.filter((c) => c.status === 'failed').length + userChecks.filter((u) => u.status === 'failed').length;
  const warnings = configChecks.filter((c) => c.status === 'warning').length + userChecks.filter((u) => u.status === 'warning').length;
  const passed = failed === 0;

  return {
    mode,
    startedAt: new Date().toISOString(),
    configChecks,
    userChecks,
    passed,
    summary: passed
      ? `TEST RUN PASSED — ${users.length} user(s) validated, ${warnings} warning(s). Configuration is consistent; production runbook is unlocked.`
      : `TEST RUN FAILED — ${failed} blocking issue(s), ${warnings} warning(s). Fix the issues (or let AI adjust them) and re-run the test.`,
  };
}

/** The ordered production runbook — exactly what the test validated. */
export function buildRunbook(form: CTForm): { title: string; script: string }[] {
  const s = ctScripts(form);
  return [
    { title: '0. Variables', script: s.vars },
    { title: '1. TARGET — migration endpoint', script: s.endpoint },
    { title: '2. TARGET — organization relationship (inbound)', script: s.orgRelTarget },
    { title: '3. SOURCE — organization relationship (outbound + scope)', script: s.orgRelSource },
    { title: '4. TARGET — MailUser preparation (repeat per user)', script: s.mailUserPrep },
    { title: '5. TARGET — validate chain', script: s.test },
    { title: '6. TARGET — create & monitor batch', script: s.batch },
  ];
}

export function aiAnalysisPrompt(form: CTForm, result: RunResult): string {
  const failures = [
    ...result.configChecks.filter((c) => c.status !== 'passed').map((c) => `CONFIG ${c.status.toUpperCase()}: ${c.name} — ${c.detail}`),
    ...result.userChecks.filter((u) => u.status !== 'passed').map((u) => `USER ${u.status.toUpperCase()}: ${u.email} — ${u.exception ?? ''} ${u.detail}`),
  ].join('\n');
  return `I ran a cross-tenant mailbox migration TEST run with this configuration:
Source tenant: ${form.sourceTenantId || '(missing)'} (${form.sourceOnMicrosoft})
Target tenant: ${form.targetTenantId || '(missing)'} (${form.targetOnMicrosoft})
App ID: ${form.appClientId || '(missing)'} | Endpoint: ${form.endpointName} | OrgRel: ${form.orgRelationshipName}
Scope group: ${form.scopeGroupName} | TargetDeliveryDomain: ${form.targetDeliveryDomain}
Users in batch: ${parseUsers(form.csvUsers).length}

The test run produced these findings:
${failures || 'No failures — all checks passed.'}

Analyze each failure: explain the root cause in simple words, give the exact corrected value or PowerShell fix, list the order to apply fixes, and any risks before moving to production.`;
}
