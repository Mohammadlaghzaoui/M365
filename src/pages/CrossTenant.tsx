import { useEffect, useRef, useState } from 'react';
import { ArrowLeftRight, AlertTriangle, CheckCircle2, XCircle, AlertCircle, PlayCircle, Rocket, Sparkles, Loader2, Cpu, Radio } from 'lucide-react';
import { PSConsole, buildConsoleScript, ConsoleLine } from '../components/PSConsole';
import { agentConfigured, agentHealth, testMigration, trackJob } from '../services/agent';
import { parseUsers } from '../services/migrationRunner';
import { Badge, Button, Card, CodeBlock, Field, PageHeader, Section, TextArea, TextOutput } from '../components/ui';
import { CT_DEFAULTS, CTForm, ctAdminConsentUrl, ctErrors, ctPostMigration, ctRollbackNotes, ctScripts, ctSourceSteps, ctTargetSteps, ctValidationChecklist, ctWarnings } from '../data/crossTenant';
import { runValidation, RunResult, buildRunbook, aiAnalysisPrompt } from '../services/migrationRunner';
import { useLocalStorage } from '../store/useLocalStorage';
import { chat } from '../services/ai';
import { aiEnabled } from '../store/settings';
import { AIHelper } from '../components/AIHelper';

type Tab = 'runner' | 'setup' | 'generator' | 'errors';

export default function CrossTenant() {
  const [form, setForm] = useLocalStorage<CTForm>('ct-form', CT_DEFAULTS);
  const [tab, setTab] = useState<Tab>('runner');
  const [result, setResult] = useState<RunResult | null>(null);
  const [aiReply, setAiReply] = useState('');
  const [aiLoading, setAiLoading] = useState(false);
  const [prodUnlocked, setProdUnlocked] = useState(false);
  const [prodConfirmed, setProdConfirmed] = useState(false);
  const [selectedError, setSelectedError] = useState(ctErrors[0].id);
  const [consoleLines, setConsoleLines] = useState<ConsoleLine[] | null>(null);
  const pendingResult = useRef<RunResult | null>(null);
  const hasAgent = agentConfigured();
  const [agentOk, setAgentOk] = useState<boolean | null>(null);
  const [liveLines, setLiveLines] = useState<{ text: string; type: 'info' | 'ok' | 'warn' | 'err' }[]>([]);
  const [liveRunning, setLiveRunning] = useState(false);
  const [livePassed, setLivePassed] = useState<boolean | null>(null);
  const liveLogRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!hasAgent) return;
    agentHealth().then(() => setAgentOk(true)).catch(() => setAgentOk(false));
  }, [hasAgent]);

  useEffect(() => { liveLogRef.current?.scrollTo({ top: liveLogRef.current.scrollHeight }); }, [liveLines]);

  const addLive = (text: string, type: 'info' | 'ok' | 'warn' | 'err' = 'info') =>
    setLiveLines((prev) => [...prev.slice(-400), { text: `[${new Date().toLocaleTimeString('en-GB')}] ${text}`, type }]);

  // LIVE validation: real Exchange Online checks via the agent on this machine.
  const runLiveValidation = async () => {
    if (liveRunning) return;
    setLiveLines([]);
    setLivePassed(null);
    setLiveRunning(true);
    addLive(`Submitting LIVE cross-tenant validation to the agent (endpoint "${form.endpointName}") ...`, 'info');
    try {
      const users = parseUsers(form.csvUsers).map((u) => ({ source: u, destination: u }));
      const { jobId } = await testMigration({
        batchName: `ct-live-${Date.now().toString(36)}`,
        endpointName: form.endpointName,
        targetDeliveryDomain: form.targetDeliveryDomain,
        users,
        crossTenant: true,
      });
      addLive(`Agent job ${jobId} accepted — streaming real Exchange Online output ...`, 'info');
      const job = await trackJob(jobId, (l) =>
        addLive(l.text, l.level === 'err' ? 'err' : l.level === 'warn' ? 'warn' : l.level === 'ok' ? 'ok' : 'info'));
      const ok = job.status === 'completed';
      setLivePassed(ok);
      addLive(ok ? 'LIVE VALIDATION PASSED — tenants are genuinely ready for this batch.' : `LIVE VALIDATION FAILED — ${job.error ?? 'see lines above'}.`, ok ? 'ok' : 'err');
    } catch (e) {
      setLivePassed(false);
      addLive(`Agent error: ${e instanceof Error ? e.message : e}`, 'err');
    } finally {
      setLiveRunning(false);
    }
  };

  const set = (k: keyof CTForm, v: string) => setForm((f) => ({ ...f, [k]: v }));
  const scripts = ctScripts(form);
  const err = ctErrors.find((e) => e.id === selectedError)!;

  const runTest = () => {
    const r = runValidation(form, 'test');
    pendingResult.current = r;
    setResult(null);
    setProdUnlocked(false);
    setProdConfirmed(false);
    setAiReply('');
    setConsoleLines(buildConsoleScript({
      command: `Start-WorkPilotMigrationTest -Type CrossTenant -TargetDeliveryDomain "${form.targetDeliveryDomain}" -Mode Test`,
      connectLines: [
        `Connecting to Exchange Online — target tenant ${form.targetOnMicrosoft} ...`,
        'Authentication successful. Session established in TEST MODE — no changes will be made.',
        `Loading migration endpoint "${form.endpointName}" ...`,
        `Resolving organization relationship "${form.orgRelationshipName}" ...`,
        `Reading migration scope group "${form.scopeGroupName}" from source tenant ...`,
        `Parsing batch CSV (${form.csvUsers.split('\n').filter((l) => l.trim() && !/^EmailAddress$/i.test(l.trim())).length} users) ...`,
      ],
      checks: r.configChecks,
      users: r.userChecks,
      summary: r.summary,
      passed: r.passed,
    }));
  };

  const onConsoleDone = () => {
    const r = pendingResult.current;
    if (!r) return;
    setResult(r);
    setProdUnlocked(r.passed);
  };

  const askAI = async () => {
    if (!result) return;
    setAiLoading(true);
    try {
      setAiReply(await chat([{ role: 'user', content: aiAnalysisPrompt(form, result) }]));
    } catch (e) {
      setAiReply(`AI error: ${e instanceof Error ? e.message : e}`);
    } finally {
      setAiLoading(false);
    }
  };

  const statusIcon = (s: string) =>
    s === 'passed' ? <CheckCircle2 size={16} className="shrink-0 text-emerald-500" />
    : s === 'warning' ? <AlertCircle size={16} className="shrink-0 text-amber-500" />
    : <XCircle size={16} className="shrink-0 text-red-500" />;

  const tabs: { id: Tab; label: string }[] = [
    { id: 'runner', label: 'Migration Runner (Test → Production)' },
    { id: 'setup', label: 'Step-by-step setup' },
    { id: 'generator', label: 'Script generator' },
    { id: 'errors', label: 'Error helper' },
  ];

  return (
    <div>
      <PageHeader title="Cross-Tenant Mailbox Migration" subtitle="Exchange Online tenant-to-tenant mailbox moves: guided setup, validated test runs, AI error analysis and a gated production runbook." icon={<ArrowLeftRight size={20} />} />

      {/* Config form — shared by all tabs */}
      <Card className="p-5 mb-5">
        <Section title="Migration configuration">
          <div className="grid gap-3 md:grid-cols-3">
            <Field label="Source tenant ID (GUID)" value={form.sourceTenantId} onChange={(v) => set('sourceTenantId', v)} placeholder="xxxxxxxx-xxxx-..." />
            <Field label="Target tenant ID (GUID)" value={form.targetTenantId} onChange={(v) => set('targetTenantId', v)} placeholder="xxxxxxxx-xxxx-..." />
            <Field label="Application (client) ID" value={form.appClientId} onChange={(v) => set('appClientId', v)} placeholder="app registration in TARGET" />
            <Field label="Source onmicrosoft domain" value={form.sourceOnMicrosoft} onChange={(v) => set('sourceOnMicrosoft', v)} />
            <Field label="Target onmicrosoft domain" value={form.targetOnMicrosoft} onChange={(v) => set('targetOnMicrosoft', v)} />
            <Field label="Client secret name" value={form.clientSecretName} onChange={(v) => set('clientSecretName', v)} />
            <Field label="Migration endpoint name" value={form.endpointName} onChange={(v) => set('endpointName', v)} />
            <Field label="Organization relationship name" value={form.orgRelationshipName} onChange={(v) => set('orgRelationshipName', v)} />
            <Field label="Security group scope name" value={form.scopeGroupName} onChange={(v) => set('scopeGroupName', v)} />
            <Field label="Target delivery domain" value={form.targetDeliveryDomain} onChange={(v) => set('targetDeliveryDomain', v)} />
          </div>
          <div className="mt-3">
            <TextArea label="User CSV list (target MailUser addresses — header EmailAddress)" value={form.csvUsers} onChange={(v) => set('csvUsers', v)} rows={4} />
          </div>
        </Section>
      </Card>

      <div className="mb-5 flex flex-wrap gap-2">
        {tabs.map((t) => (
          <button key={t.id} onClick={() => setTab(t.id)}
            className={`rounded-lg px-4 py-2 text-sm font-semibold ${tab === t.id ? 'bg-blue-600 text-white' : 'border border-slate-300 dark:border-slate-600 text-slate-600 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-700'}`}>
            {t.label}
          </button>
        ))}
      </div>

      {/* ============ RUNNER ============ */}
      {tab === 'runner' && (
        <div className="space-y-5">
          <Card className="p-5 border-violet-200 dark:border-violet-800">
            <div className="flex flex-wrap items-center gap-3">
              <div className="flex-1 min-w-64">
                <h3 className="font-bold text-slate-800 dark:text-slate-100 flex items-center gap-2"><PlayCircle size={18} className="text-violet-500" /> Test mode first — always</h3>
                <p className="mt-1 text-sm text-slate-500 dark:text-slate-400">
                  The test run validates your whole configuration locally (tenant IDs, app, scope, delivery domain, every CSV user) against the rules that cause real cross-tenant failures — and shows the exact exception each problem would produce. Production stays locked until the test passes.
                </p>
              </div>
              <Button variant="ai" onClick={runTest}><PlayCircle size={16} /> Run test migration</Button>
            </div>
          </Card>

          {/* LIVE validation via agent — real Exchange Online checks */}
          <Card className={`p-5 ${agentOk ? (livePassed === true ? 'border-emerald-300 dark:border-emerald-700' : livePassed === false ? 'border-red-300 dark:border-red-700' : 'border-blue-300 dark:border-blue-700') : 'opacity-80'}`}>
            <div className="flex flex-wrap items-center gap-3">
              <span className={`rounded-lg p-2 ${agentOk ? 'bg-emerald-100 dark:bg-emerald-900/40 text-emerald-600' : 'bg-slate-100 dark:bg-slate-700 text-slate-400'}`}><Radio size={18} /></span>
              <div className="flex-1 min-w-64">
                <h3 className="font-bold text-slate-800 dark:text-slate-100 flex items-center gap-2">
                  LIVE validation — real Exchange Online
                  {livePassed === true && <Badge color="green">PASSED</Badge>}
                  {livePassed === false && <Badge color="red">FAILED</Badge>}
                </h3>
                <p className="text-sm text-slate-500 dark:text-slate-400">
                  {agentOk
                    ? 'Runs the REAL checks via your agent: organization relationship, Get-MigrationEndpoint + Test-MigrationServerAvailability, and per user Get-MailUser ExchangeGuid/targetAddress in the target tenant. Read-only — nothing is migrated.'
                    : hasAgent
                      ? 'Agent configured but unreachable — start it on your PC (start.bat) and refresh.'
                      : <>Connect the <a href="#/agent" className="text-blue-500 hover:underline">Migration Agent</a> to run this validation for real against your tenants.</>}
                </p>
              </div>
              <Button onClick={runLiveValidation} disabled={!agentOk || liveRunning}>
                {liveRunning ? <Loader2 size={16} className="animate-spin" /> : <Cpu size={16} />} Run LIVE validation
              </Button>
            </div>
            {liveLines.length > 0 && (
              <div className="mt-4 overflow-hidden rounded-xl border border-slate-700 bg-[#012456]">
                <div className="flex items-center gap-2 border-b border-slate-700 bg-slate-900 px-3 py-2">
                  <span className="h-3 w-3 rounded-full bg-red-500" /><span className="h-3 w-3 rounded-full bg-amber-400" /><span className="h-3 w-3 rounded-full bg-emerald-500" />
                  <span className="ml-2 text-xs font-medium text-slate-300">Exchange Online — live cross-tenant validation (agent)</span>
                </div>
                <div ref={liveLogRef} className="h-64 overflow-y-auto p-3 font-mono text-xs leading-relaxed">
                  {liveLines.map((l, i) => (
                    <div key={i} className={`whitespace-pre-wrap ${{ info: 'text-cyan-300', ok: 'text-emerald-400', warn: 'text-amber-300', err: 'text-red-400' }[l.type]}`}>{l.text}</div>
                  ))}
                  {liveRunning && <span className="inline-block h-3.5 w-2 animate-pulse bg-slate-200 align-middle" />}
                </div>
              </div>
            )}
          </Card>

          {consoleLines && (
            <PSConsole lines={consoleLines} onDone={onConsoleDone} title={`Windows PowerShell — Cross-Tenant Migration Runner (${result ? 'completed' : 'running'})`} />
          )}

          {result && (
            <>
              <Card className={`p-5 ${result.passed ? 'border-emerald-300 dark:border-emerald-700' : 'border-red-300 dark:border-red-700'}`}>
                <div className="flex items-center gap-3">
                  {result.passed ? <CheckCircle2 size={24} className="text-emerald-500" /> : <XCircle size={24} className="text-red-500" />}
                  <div>
                    <h3 className="font-bold text-slate-800 dark:text-slate-100">{result.passed ? 'Test run PASSED' : 'Test run FAILED'}</h3>
                    <p className="text-sm text-slate-500 dark:text-slate-400">{result.summary}</p>
                  </div>
                  {aiEnabled() && !result.passed && (
                    <Button variant="ai" className="ml-auto" onClick={askAI} disabled={aiLoading}>
                      {aiLoading ? <Loader2 size={16} className="animate-spin" /> : <Sparkles size={16} />} AI: analyze & fix
                    </Button>
                  )}
                </div>
              </Card>

              {aiReply && (
                <Card className="p-5 border-violet-200 dark:border-violet-800">
                  <h4 className="mb-2 text-sm font-semibold text-violet-700 dark:text-violet-300 flex items-center gap-2"><Sparkles size={15} /> AI analysis & fixes</h4>
                  <pre className="whitespace-pre-wrap text-sm text-slate-700 dark:text-slate-200 font-sans leading-relaxed">{aiReply}</pre>
                  <p className="mt-3 text-xs text-slate-400">Apply the fixes in the configuration above, then re-run the test until it passes.</p>
                </Card>
              )}

              <div className="grid gap-5 lg:grid-cols-2">
                <Card className="p-5">
                  <Section title={`Configuration checks (${result.configChecks.length})`}>
                    <div className="space-y-2.5">
                      {result.configChecks.map((c) => (
                        <div key={c.name} className="flex items-start gap-2.5">
                          {statusIcon(c.status)}
                          <div className="text-sm">
                            <div className="font-medium text-slate-700 dark:text-slate-200">{c.name}</div>
                            <div className="text-slate-500 dark:text-slate-400">{c.detail}</div>
                            {c.fix && <div className="mt-0.5 text-blue-600 dark:text-blue-400">Fix: {c.fix}</div>}
                          </div>
                        </div>
                      ))}
                    </div>
                  </Section>
                </Card>
                <Card className="p-5">
                  <Section title={`Simulated batch users (${result.userChecks.length})`}>
                    <div className="space-y-2.5">
                      {result.userChecks.map((u, i) => (
                        <div key={i} className="flex items-start gap-2.5">
                          {statusIcon(u.status)}
                          <div className="text-sm">
                            <div className="font-medium text-slate-700 dark:text-slate-200">{u.email} {u.exception && <Badge color="red">{u.exception}</Badge>}</div>
                            <div className="text-slate-500 dark:text-slate-400">{u.detail}</div>
                            {u.fix && <div className="mt-0.5 text-blue-600 dark:text-blue-400">Fix: {u.fix}</div>}
                          </div>
                        </div>
                      ))}
                    </div>
                  </Section>
                </Card>
              </div>
            </>
          )}

          {/* Production */}
          <Card className={`p-5 ${prodUnlocked ? 'border-emerald-300 dark:border-emerald-700' : 'opacity-70'}`}>
            <h3 className="font-bold text-slate-800 dark:text-slate-100 flex items-center gap-2"><Rocket size={18} className="text-emerald-500" /> Production mode</h3>
            {!prodUnlocked ? (
              <p className="mt-1 text-sm text-slate-500 dark:text-slate-400">Locked — run the test until it passes. Production emits the exact validated runbook in execution order.</p>
            ) : (
              <>
                <div className="mt-3 rounded-lg bg-amber-50 dark:bg-amber-900/20 border border-amber-200 dark:border-amber-800 p-3 text-sm text-amber-700 dark:text-amber-300 flex gap-2">
                  <AlertTriangle size={16} className="mt-0.5 shrink-0" />
                  <span>Production gate: completing a cross-tenant batch CONVERTS the source mailboxes — there is no automatic rollback. Confirm you have pilot results, customer approval and a final-delta window before executing.</span>
                </div>
                <label className="mt-3 flex items-center gap-2 text-sm font-medium text-slate-700 dark:text-slate-200">
                  <input type="checkbox" checked={prodConfirmed} onChange={(e) => setProdConfirmed(e.target.checked)} className="h-4 w-4 rounded border-slate-300 text-emerald-600" />
                  I have approval and pilot evidence — show the production runbook
                </label>
                {prodConfirmed && (
                  <div className="mt-4 space-y-4">
                    {buildRunbook(form).map((step) => (
                      <div key={step.title}>
                        <div className="mb-1 text-sm font-semibold text-slate-600 dark:text-slate-300">{step.title}</div>
                        <CodeBlock code={step.script} />
                      </div>
                    ))}
                    <p className="text-xs text-slate-400">Execute these blocks in order from an Exchange Online PowerShell session with the indicated tenant context. Hooking this runbook to automated server-side execution is prepared via the ExecutionAdapter interface (see services/migrationRunner.ts) — a backend with app credentials is required for that; browser-side execution of tenant admin commands is deliberately not done.</p>
                  </div>
                )}
              </>
            )}
          </Card>
        </div>
      )}

      {/* ============ SETUP ============ */}
      {tab === 'setup' && (
        <div className="space-y-5">
          <Card className="p-5 border-amber-200 dark:border-amber-800">
            <Section title="Critical warnings — read before starting">
              <ul className="space-y-2">
                {ctWarnings.map((w, i) => (
                  <li key={i} className="flex items-start gap-2 text-sm text-amber-700 dark:text-amber-300">
                    <AlertTriangle size={15} className="mt-0.5 shrink-0" />{w}
                  </li>
                ))}
              </ul>
            </Section>
          </Card>
          <div className="grid gap-5 lg:grid-cols-2">
            <Card className="p-5">
              <Section title="Target tenant setup steps">
                <ol className="space-y-2.5">
                  {ctTargetSteps.map((s, i) => (
                    <li key={i} className="flex items-start gap-3 text-sm text-slate-700 dark:text-slate-200">
                      <span className="flex h-6 w-6 shrink-0 items-center justify-center rounded-full bg-blue-100 dark:bg-blue-900/50 text-xs font-bold text-blue-700 dark:text-blue-300">{i + 1}</span>
                      <span className="pt-0.5">{s}</span>
                    </li>
                  ))}
                </ol>
              </Section>
            </Card>
            <Card className="p-5">
              <Section title="Source tenant setup steps">
                <ol className="space-y-2.5">
                  {ctSourceSteps.map((s, i) => (
                    <li key={i} className="flex items-start gap-3 text-sm text-slate-700 dark:text-slate-200">
                      <span className="flex h-6 w-6 shrink-0 items-center justify-center rounded-full bg-violet-100 dark:bg-violet-900/50 text-xs font-bold text-violet-700 dark:text-violet-300">{i + 1}</span>
                      <span className="pt-0.5">{s}</span>
                    </li>
                  ))}
                </ol>
              </Section>
            </Card>
          </div>
          <Card className="p-5">
            <Section title="Admin consent URL (send to SOURCE tenant Global Admin)">
              <CodeBlock code={ctAdminConsentUrl(form.sourceOnMicrosoft, form.appClientId)} />
            </Section>
          </Card>
          <div className="grid gap-5 lg:grid-cols-3">
            <TextOutput title="Validation checklist" text={ctValidationChecklist.map((c) => `[ ] ${c}`).join('\n')} />
            <TextOutput title="Post-migration / user instructions" text={ctPostMigration.map((c) => `- ${c}`).join('\n')} />
            <TextOutput title="Rollback notes" text={ctRollbackNotes.map((c) => `- ${c}`).join('\n')} />
          </div>
        </div>
      )}

      {/* ============ GENERATOR ============ */}
      {tab === 'generator' && (
        <div className="space-y-4">
          {[
            { title: 'PowerShell variables', code: scripts.vars },
            { title: 'Migration endpoint (TARGET)', code: scripts.endpoint },
            { title: 'Organization relationship — TARGET (inbound)', code: scripts.orgRelTarget },
            { title: 'Organization relationship — SOURCE (outbound + scope)', code: scripts.orgRelSource },
            { title: 'MailUser preparation (per user, TARGET)', code: scripts.mailUserPrep },
            { title: 'Test migration (validate chain)', code: scripts.test },
            { title: 'Migration batch (create / monitor / complete)', code: scripts.batch },
            { title: 'CSV template', code: scripts.csv },
          ].map((b) => (
            <div key={b.title}>
              <div className="mb-1 text-sm font-semibold text-slate-600 dark:text-slate-300">{b.title}</div>
              <CodeBlock code={b.code} />
            </div>
          ))}
        </div>
      )}

      {/* ============ ERRORS ============ */}
      {tab === 'errors' && (
        <div className="grid gap-6 lg:grid-cols-[340px_1fr]">
          <div className="space-y-1.5 lg:sticky lg:top-20 lg:self-start">
            {ctErrors.map((e) => (
              <button key={e.id} onClick={() => setSelectedError(e.id)}
                className={`w-full rounded-lg px-3 py-2.5 text-left text-xs font-mono font-medium break-all ${selectedError === e.id ? 'bg-red-600 text-white' : 'border border-slate-200 dark:border-slate-700 text-slate-600 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-700'}`}>
                {e.name}
              </button>
            ))}
          </div>
          <div className="space-y-4">
            <Card className="p-5">
              <h3 className="font-mono text-sm font-bold text-red-600 dark:text-red-400 break-all">{err.name}</h3>
              <Section title="Meaning in simple words"><p className="text-sm text-slate-700 dark:text-slate-200">{err.meaning}</p></Section>
              <Section title="Likely cause"><p className="text-sm text-slate-700 dark:text-slate-200">{err.cause}</p></Section>
              <Section title="Fix steps">
                <ol className="space-y-2">
                  {err.fix.map((f, i) => (
                    <li key={i} className="flex items-start gap-2.5 text-sm text-slate-700 dark:text-slate-200">
                      <span className="flex h-5 w-5 shrink-0 items-center justify-center rounded-full bg-emerald-100 dark:bg-emerald-900/50 text-xs font-bold text-emerald-700 dark:text-emerald-300">{i + 1}</span>{f}
                    </li>
                  ))}
                </ol>
              </Section>
              {err.powershell && <Section title="PowerShell"><CodeBlock code={err.powershell} /></Section>}
              <Section title="Prevention"><p className="text-sm text-slate-700 dark:text-slate-200">{err.prevention}</p></Section>
              <Section title="Escalation note"><p className="text-sm text-amber-700 dark:text-amber-300">{err.escalationNote}</p></Section>
            </Card>
            <AIHelper context={`Cross-tenant migration error: ${err.name}\nMeaning: ${err.meaning}\nCause: ${err.cause}`} defaultPrompt="I hit this error in a live migration. Walk me through diagnosing it step by step with exact PowerShell." />
          </div>
        </div>
      )}
    </div>
  );
}
