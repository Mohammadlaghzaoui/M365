import { useEffect, useMemo, useRef, useState } from 'react';
import { MonitorPlay, PlayCircle, ShieldCheck, Gauge, Layers, Zap, RotateCcw, Download, Upload, Trash2, Wrench } from 'lucide-react';
import { Badge, Button, Card, Field, PageHeader, ProgressBar, Section, Select, TextArea } from '../components/ui';
import { uid, useLocalStorage } from '../store/useLocalStorage';
import { ConsoleProjectState, MigrationLineItem, MigrationPass, MigrationProject } from '../types';
import { buildStages, migrationTypeLabels } from '../data/migrationStages';
import { getIntegrations } from '../store/settings';
import { agentConfigured, verifyEndpoint, startMigration, trackJob, agentHealth, AgentHealth } from '../services/agent';
import { Link } from 'react-router-dom';
import { Cpu, CheckCircle2, XCircle } from 'lucide-react';

/**
 * Migration Console — MigrationWiz-style operational console on top of the
 * migration projects: endpoints, line items, verify / assessment / pre-stage /
 * full / delta / retry passes with a live execution log, per-item statistics
 * and CSV reporting.
 *
 * The pass engine runs in SIMULATION/PLANNING mode in this browser version:
 * outcomes are deterministic per mailbox (same mailbox = same result) and
 * mirror the real MigrationWiz failure categories, so the console doubles as
 * cutover rehearsal and training. The BitTitan API key from Settings is the
 * hook for a backend that drives the real MigrationWiz REST API later.
 */

const hash = (s: string) => [...s].reduce((a, c) => (a * 31 + c.charCodeAt(0)) >>> 0, 7);

const endpointTypes = ['Microsoft 365', 'Google Workspace', 'Exchange On-Premises (EWS)', 'IMAP'];

const statusColor: Record<MigrationLineItem['status'], string> = {
  NotStarted: 'gray', Verifying: 'blue', Verified: 'green', VerifyFailed: 'red',
  Assessing: 'blue', Assessed: 'green', PreStaging: 'blue', PreStaged: 'purple',
  Migrating: 'blue', DeltaSync: 'blue', Completed: 'green', Failed: 'red',
};

interface LogLine { text: string; type: 'info' | 'ok' | 'warn' | 'err' | 'cmd' }

function LiveLog({ lines }: { lines: LogLine[] }) {
  const ref = useRef<HTMLDivElement>(null);
  useEffect(() => { ref.current?.scrollTo({ top: ref.current.scrollHeight }); }, [lines]);
  const color = { info: 'text-cyan-300', ok: 'text-emerald-400', warn: 'text-amber-300', err: 'text-red-400', cmd: 'text-yellow-300' };
  return (
    <div className="overflow-hidden rounded-xl border border-slate-700 bg-[#012456]">
      <div className="flex items-center gap-2 border-b border-slate-700 bg-slate-900 px-3 py-2">
        <span className="h-3 w-3 rounded-full bg-red-500" /><span className="h-3 w-3 rounded-full bg-amber-400" /><span className="h-3 w-3 rounded-full bg-emerald-500" />
        <span className="ml-2 text-xs font-medium text-slate-300">Migration Console — live execution log</span>
      </div>
      <div ref={ref} className="h-56 overflow-y-auto p-3 font-mono text-xs leading-relaxed">
        {lines.length === 0 && <span className="text-slate-500">Waiting for a pass to start…</span>}
        {lines.map((l, i) => <div key={i} className={`whitespace-pre-wrap ${color[l.type]}`}>{l.text}</div>)}
        <span className="inline-block h-3.5 w-2 animate-pulse bg-slate-200 align-middle" />
      </div>
    </div>
  );
}

const emptyState = (): ConsoleProjectState => ({
  sourceEndpoint: { type: 'Microsoft 365', verified: false },
  destEndpoint: { type: 'Microsoft 365', verified: false },
  items: [],
});

export default function MigrationConsole() {
  const [projects, setProjects] = useLocalStorage<MigrationProject[]>('migration-projects', []);
  const [states, setStates] = useLocalStorage<Record<string, ConsoleProjectState>>('mw-console', {});
  const [projectId, setProjectId] = useState<string>(projects[0]?.id ?? '');
  const [csv, setCsv] = useState('source@old-tenant.com,user@new-tenant.com');
  const [log, setLog] = useState<LogLine[]>([]);
  const [running, setRunning] = useState<MigrationPass | null>(null);
  const [quickName, setQuickName] = useState('');
  const timer = useRef<ReturnType<typeof setInterval> | null>(null);
  const bittitanKey = getIntegrations().bittitan;
  const hasAgent = agentConfigured();
  const [agentInfo, setAgentInfo] = useState<AgentHealth | null>(null);
  const [agentOk, setAgentOk] = useState<boolean | null>(null);

  useEffect(() => {
    if (!hasAgent) return;
    agentHealth().then((h) => { setAgentInfo(h); setAgentOk(true); }).catch(() => setAgentOk(false));
  }, [hasAgent]);

  const verifyViaAgent = async () => {
    addLog('Agent: verifying endpoints (real connectivity check) ...', 'cmd');
    for (const [key, type] of [['sourceEndpoint', 'source'], ['destEndpoint', 'destination']] as const) {
      const epType = state[key].type === 'Exchange On-Premises (EWS)' ? 'exchange-onprem' : state[key].type === 'Microsoft 365' ? 'graph' : 'graph';
      try {
        const r = await verifyEndpoint(epType);
        if (r.verified) { setState((s) => ({ ...s, [key]: { ...s[key], verified: true } })); addLog(`Agent: ${type} endpoint OK — ${r.detail ?? ''}`, 'ok'); }
        else { addLog(`Agent: ${type} endpoint failed — ${r.error ?? ''}`, 'err'); }
      } catch (e) {
        addLog(`Agent: ${type} verify error — ${e instanceof Error ? e.message : e}`, 'err');
      }
    }
  };

  const fullMigrateViaAgent = async () => {
    const queue = items.filter((i) => ['Verified', 'Assessed', 'PreStaged'].includes(i.status));
    if (!queue.length) { addLog('No eligible items for the agent migration.', 'warn'); return; }
    addLog(`Agent: submitting migration batch for ${queue.length} mailbox(es) ...`, 'cmd');
    try {
      const { jobId } = await startMigration({
        batchName: `${project?.name ?? 'batch'}-${Date.now().toString(36)}`,
        users: queue.map((i) => ({ source: i.sourceEmail, destination: i.destEmail })),
        // The validated PowerShell command is supplied by the operator workflow;
        // here we pass the intent and let the agent run its New-MigrationBatch.
        script: '',
      });
      addLog(`Agent: job ${jobId} accepted. Streaming log ...`, 'info');
      const job = await trackJob(jobId, (l) => addLog(`Agent» ${l.text}`, l.level === 'err' ? 'err' : l.level === 'ok' ? 'ok' : 'info'));
      addLog(`Agent: job ${job.status}.`, job.status === 'completed' ? 'ok' : 'err');
    } catch (e) {
      addLog(`Agent migration error: ${e instanceof Error ? e.message : e}`, 'err');
    }
  };

  const project = projects.find((p) => p.id === projectId) ?? null;
  const state = states[projectId] ?? emptyState();
  const items = state.items;

  const setState = (fn: (s: ConsoleProjectState) => ConsoleProjectState) =>
    setStates((prev) => ({ ...prev, [projectId]: fn(prev[projectId] ?? emptyState()) }));

  const addLog = (text: string, type: LogLine['type'] = 'info') =>
    setLog((prev) => [...prev.slice(-400), { text: `[${new Date().toLocaleTimeString('en-GB')}] ${text}`, type }]);

  useEffect(() => () => { if (timer.current) clearInterval(timer.current); }, []);

  const stats = useMemo(() => ({
    total: items.length,
    completed: items.filter((i) => i.status === 'Completed').length,
    failed: items.filter((i) => i.status === 'Failed' || i.status === 'VerifyFailed').length,
    mbMigrated: Math.round(items.reduce((a, i) => a + i.mbMigrated, 0)),
    mbTotal: Math.round(items.reduce((a, i) => a + i.mbTotal, 0)),
    itemsFailed: items.reduce((a, i) => a + i.itemsFailed, 0),
  }), [items]);

  const createQuickProject = () => {
    const p: MigrationProject = {
      id: uid(), createdAt: new Date().toISOString(), name: quickName || 'Console project',
      customer: '', type: 'm365-to-m365', sourceEnv: '', destTenant: '', domains: '',
      users: 0, sharedMailboxes: 0, groups: 0, teams: 0, mailboxSizeNotes: '',
      tool: 'BitTitan MigrationWiz', cutoverDate: '', status: 'active', notes: '', stages: buildStages(),
    };
    setProjects((prev) => [p, ...prev]);
    setProjectId(p.id);
  };

  const importItems = () => {
    const rows = csv.split('\n').map((l) => l.trim()).filter((l) => l && !/^source/i.test(l));
    const newItems: MigrationLineItem[] = rows.map((r) => {
      const [src, dst] = r.split(/[,;\t]/).map((x) => x?.trim() ?? '');
      const h = hash(src || r);
      return {
        id: uid(), sourceEmail: src || r, destEmail: dst || src || r,
        status: 'NotStarted', progress: 0,
        itemsTotal: 2000 + (h % 60000), itemsMigrated: 0, itemsFailed: 0,
        mbTotal: 250 + (h % 3800), mbMigrated: 0,
      };
    });
    setState((s) => ({ ...s, items: [...s.items, ...newItems] }));
    addLog(`Imported ${newItems.length} line item(s) from CSV.`, 'ok');
  };

  // Deterministic failure model mirroring real MigrationWiz error categories.
  const outcome = (item: MigrationLineItem, pass: MigrationPass): { fail?: string; transient?: boolean } => {
    const h = hash(item.sourceEmail);
    if (pass === 'verify' && h % 9 === 2 && item.lastPass !== 'verify') {
      return { fail: 'Authentication failed / 401 — endpoint credential rejected for this mailbox', transient: true };
    }
    if ((pass === 'prestage' || pass === 'full') && !item.error) {
      if (h % 11 === 4) return { fail: 'Connection did not succeed / throttling (HTTP 503, ErrorServerBusy)', transient: true };
      if (h % 13 === 6) return { fail: 'Item exceeds maximum allowed size — raise destination MaxReceiveSize', transient: false };
    }
    return {};
  };

  const eligible = (i: MigrationLineItem, pass: MigrationPass): boolean => {
    switch (pass) {
      case 'verify': return i.status === 'NotStarted' || i.status === 'VerifyFailed';
      case 'assessment': return i.status === 'Verified';
      case 'prestage': return i.status === 'Verified' || i.status === 'Assessed';
      case 'full': return ['Verified', 'Assessed', 'PreStaged'].includes(i.status);
      case 'delta': return i.status === 'Completed';
      case 'retry': return (i.status === 'Failed' && !!i.errorTransient) || i.status === 'VerifyFailed';
      default: return false;
    }
  };

  const runningStatus: Record<MigrationPass, MigrationLineItem['status']> = {
    verify: 'Verifying', assessment: 'Assessing', prestage: 'PreStaging', full: 'Migrating', delta: 'DeltaSync', retry: 'Migrating',
  };
  const doneStatus: Record<MigrationPass, MigrationLineItem['status']> = {
    verify: 'Verified', assessment: 'Assessed', prestage: 'PreStaged', full: 'Completed', delta: 'Completed', retry: 'Completed',
  };

  const startPass = (pass: MigrationPass) => {
    if (running) return;
    const queue = items.filter((i) => eligible(i, pass)).map((i) => i.id);
    if (!queue.length) { addLog(`No eligible line items for pass "${pass}".`, 'warn'); return; }
    if (pass !== 'verify' && (!state.sourceEndpoint.verified || !state.destEndpoint.verified)) {
      addLog('Endpoints not verified — run Verify Credentials first.', 'err'); return;
    }
    setRunning(pass);
    addLog(`Start-MigrationPass -Type ${pass} -Items ${queue.length} -Concurrency 3`, 'cmd');
    addLog(`${pass.toUpperCase()} pass started on ${queue.length} item(s)${pass === 'retry' ? ' (transient errors only — persistent need Mark fixed)' : ''}.`, 'info');

    const active = new Set<string>();
    timer.current = setInterval(() => {
      setStates((prev) => {
        const st = prev[projectId] ?? emptyState();
        let queueChanged = false;
        const next = st.items.map((i) => {
          if (!queue.includes(i.id)) return i;
          if (!active.has(i.id) && active.size < 3 && !['Completed', 'Failed', 'Verified', 'Assessed', 'PreStaged', 'VerifyFailed'].includes(i.status === runningStatus[pass] ? 'x' : i.status)) {
            if (eligible(i, pass)) {
              active.add(i.id);
              addLog(`Processing ${i.sourceEmail} → ${i.destEmail} ...`);
              return { ...i, status: runningStatus[pass], progress: pass === 'verify' || pass === 'assessment' ? 40 : 5, error: pass === 'retry' ? undefined : i.error, lastPass: pass };
            }
          }
          if (active.has(i.id)) {
            const speed = pass === 'verify' || pass === 'assessment' || pass === 'delta' ? 35 : 7 + (hash(i.id) % 9);
            const progress = Math.min(100, i.progress + speed);
            if (progress >= 100) {
              active.delete(i.id);
              queueChanged = true;
              queue.splice(queue.indexOf(i.id), 1);
              const res = outcome(i, pass === 'retry' ? (i.mbMigrated > 0 ? 'full' : 'prestage') : pass);
              if (res.fail && pass !== 'retry') {
                addLog(`${i.sourceEmail}: FAILED — ${res.fail}`, 'err');
                return { ...i, progress: 0, status: pass === 'verify' ? 'VerifyFailed' as const : 'Failed' as const, error: res.fail, errorTransient: res.transient, itemsFailed: i.itemsFailed + (pass === 'verify' ? 0 : 12 + hash(i.id) % 90) };
              }
              const mbDone = pass === 'prestage' ? i.mbTotal * 0.8 : pass === 'assessment' || pass === 'verify' ? i.mbMigrated : i.mbTotal;
              const itemsDone = pass === 'verify' || pass === 'assessment' ? i.itemsMigrated : pass === 'prestage' ? Math.round(i.itemsTotal * 0.8) : i.itemsTotal;
              addLog(`${i.sourceEmail}: ${doneStatus[pass]} (${Math.round(mbDone)} MB, ${itemsDone.toLocaleString()} items)`, 'ok');
              return { ...i, progress: 0, status: doneStatus[pass], error: undefined, errorTransient: undefined, mbMigrated: mbDone, itemsMigrated: itemsDone };
            }
            return { ...i, progress };
          }
          return i;
        });
        if (queueChanged && queue.length === 0 && active.size === 0) {
          if (timer.current) clearInterval(timer.current);
          setTimeout(() => {
            setRunning(null);
            addLog(`${pass.toUpperCase()} pass completed.`, 'ok');
          }, 50);
        }
        return { ...prev, [projectId]: { ...st, items: next } };
      });
    }, 350);
  };

  const verifyEndpoints = () => {
    addLog('Verifying source endpoint credentials ...', 'info');
    setTimeout(() => {
      setState((s) => ({ ...s, sourceEndpoint: { ...s.sourceEndpoint, verified: true } }));
      addLog(`Source endpoint (${state.sourceEndpoint.type}): credentials OK.`, 'ok');
    }, 900);
    setTimeout(() => {
      setState((s) => ({ ...s, destEndpoint: { ...s.destEndpoint, verified: true } }));
      addLog(`Destination endpoint (${state.destEndpoint.type}): credentials OK.`, 'ok');
    }, 1700);
  };

  const exportCsv = () => {
    const rows = [
      'Source,Destination,Status,ItemsTotal,ItemsMigrated,ItemsFailed,MBTotal,MBMigrated,Error',
      ...items.map((i) => `${i.sourceEmail},${i.destEmail},${i.status},${i.itemsTotal},${i.itemsMigrated},${i.itemsFailed},${Math.round(i.mbTotal)},${Math.round(i.mbMigrated)},"${i.error ?? ''}"`),
    ];
    const blob = new Blob([rows.join('\n')], { type: 'text/csv' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `migration-report-${project?.name ?? 'project'}-${new Date().toISOString().slice(0, 10)}.csv`;
    a.click();
    URL.revokeObjectURL(a.href);
    addLog('Statistics report exported (CSV).', 'ok');
  };

  const markFixed = (id: string) =>
    setState((s) => ({ ...s, items: s.items.map((i) => (i.id === id ? { ...i, errorTransient: true, error: `${i.error} — marked resolved by operator` } : i)) }));

  return (
    <div>
      <PageHeader title="Migration Console" subtitle="Endpoints, line items, verify → assessment → pre-stage → full → delta passes, retry handling, live log and statistics export. Connect the Migration Agent for real execution; without it the engine runs as a deterministic rehearsal." icon={<MonitorPlay size={20} />} />

      {/* Agent status banner */}
      <Card className={`mb-5 p-4 ${hasAgent && agentOk ? 'border-emerald-300 dark:border-emerald-700' : hasAgent && agentOk === false ? 'border-red-300 dark:border-red-700' : 'border-violet-200 dark:border-violet-800'}`}>
        <div className="flex flex-wrap items-center gap-3">
          <span className={`rounded-lg p-2 ${hasAgent && agentOk ? 'bg-emerald-100 dark:bg-emerald-900/40 text-emerald-600' : 'bg-violet-100 dark:bg-violet-900/40 text-violet-600'}`}><Cpu size={18} /></span>
          <div className="flex-1 min-w-64">
            {!hasAgent ? (
              <><div className="text-sm font-semibold text-slate-800 dark:text-slate-100">Migration Agent not connected — rehearsal mode</div>
              <div className="text-xs text-slate-500 dark:text-slate-400">Passes simulate deterministically. Connect the agent in <Link to="/settings" className="text-blue-500 hover:underline">Settings → Integrations</Link> for real execution against AD / Exchange / Graph.</div></>
            ) : agentOk === null ? (
              <div className="text-sm text-slate-500">Checking agent …</div>
            ) : agentOk ? (
              <><div className="flex items-center gap-2 text-sm font-semibold text-emerald-700 dark:text-emerald-300"><CheckCircle2 size={15} /> Agent connected — REAL execution enabled</div>
              <div className="text-xs text-slate-500 dark:text-slate-400">{agentInfo?.name} v{agentInfo?.version} on {agentInfo?.host} · PowerShell {agentInfo?.capabilities.powershell ? '✓' : '✗'} · Graph {agentInfo?.capabilities.graph ? '✓' : '✗'} · modules: {agentInfo?.capabilities.modules.join(', ') || '—'}</div></>
            ) : (
              <><div className="flex items-center gap-2 text-sm font-semibold text-red-600 dark:text-red-400"><XCircle size={15} /> Agent unreachable</div>
              <div className="text-xs text-slate-500 dark:text-slate-400">Check the agent URL/key in Settings and that the agent service is running.</div></>
            )}
          </div>
          <Badge color={hasAgent && agentOk ? 'green' : 'purple'}>{hasAgent && agentOk ? 'LIVE' : 'REHEARSAL'}</Badge>
        </div>
      </Card>

      {/* Project picker */}
      <Card className="mb-5 p-4">
        <div className="flex flex-wrap items-end gap-3">
          {projects.length > 0 && (
            <div className="w-72">
              <Select label="Migration project" value={projectId} onChange={setProjectId}
                options={projects.map((p) => ({ value: p.id, label: `${p.name} (${migrationTypeLabels[p.type]})` }))} />
            </div>
          )}
          <Field label="Quick project name" value={quickName} onChange={setQuickName} placeholder="Customer wave 1" />
          <Button variant="secondary" onClick={createQuickProject}>+ New console project</Button>
          {!bittitanKey.enabled && (
            <span className="text-xs text-slate-400">Tip: store your BitTitan API key in <Link className="text-blue-500 hover:underline" to="/settings">Settings → Integrations</Link> for future live execution.</span>
          )}
        </div>
      </Card>

      {!project ? (
        <Card className="p-8 text-center text-sm text-slate-400">Create or select a project to open the console.</Card>
      ) : (
        <div className="space-y-5">
          {/* Stats */}
          <div className="grid grid-cols-2 gap-4 lg:grid-cols-5">
            <Card className="p-4"><div className="text-xs font-semibold uppercase text-slate-400">Line items</div><div className="mt-1 text-2xl font-bold text-blue-600 dark:text-blue-400">{stats.total}</div></Card>
            <Card className="p-4"><div className="text-xs font-semibold uppercase text-slate-400">Completed</div><div className="mt-1 text-2xl font-bold text-emerald-500">{stats.completed}</div></Card>
            <Card className="p-4"><div className="text-xs font-semibold uppercase text-slate-400">Failed</div><div className={`mt-1 text-2xl font-bold ${stats.failed ? 'text-red-500' : 'text-emerald-500'}`}>{stats.failed}</div></Card>
            <Card className="p-4"><div className="text-xs font-semibold uppercase text-slate-400">Data migrated</div><div className="mt-1 text-2xl font-bold text-violet-500">{(stats.mbMigrated / 1024).toFixed(1)} GB</div><ProgressBar value={stats.mbTotal ? (stats.mbMigrated / stats.mbTotal) * 100 : 0} color="bg-violet-500" /></Card>
            <Card className="p-4"><div className="text-xs font-semibold uppercase text-slate-400">Failed items</div><div className="mt-1 text-2xl font-bold text-amber-500">{stats.itemsFailed.toLocaleString()}</div></Card>
          </div>

          {/* Endpoints */}
          <Card className="p-5">
            <Section title="Endpoints">
              <div className="grid gap-4 sm:grid-cols-2">
                {(['sourceEndpoint', 'destEndpoint'] as const).map((key) => (
                  <div key={key} className="rounded-lg border border-slate-200 dark:border-slate-700 p-3">
                    <div className="mb-2 flex items-center justify-between">
                      <span className="text-sm font-semibold text-slate-700 dark:text-slate-200">{key === 'sourceEndpoint' ? 'Source endpoint' : 'Destination endpoint'}</span>
                      <Badge color={state[key].verified ? 'green' : 'orange'}>{state[key].verified ? 'verified' : 'not verified'}</Badge>
                    </div>
                    <Select label="" value={state[key].type}
                      onChange={(v) => setState((s) => ({ ...s, [key]: { type: v, verified: false } }))}
                      options={endpointTypes.map((t) => ({ value: t, label: t }))} />
                  </div>
                ))}
              </div>
              <div className="mt-3"><Button variant="secondary" onClick={hasAgent && agentOk ? verifyViaAgent : verifyEndpoints}><ShieldCheck size={15} /> Verify endpoint credentials{hasAgent && agentOk ? ' (agent)' : ''}</Button></div>
            </Section>
          </Card>

          {/* Import + pass toolbar */}
          <Card className="p-5">
            <Section title="Import line items (CSV: source,destination)">
              <TextArea label="" value={csv} onChange={setCsv} rows={3} />
              <div className="mt-2"><Button variant="secondary" onClick={importItems}><Upload size={15} /> Import users</Button></div>
            </Section>
            <Section title="Migration passes">
              <div className="flex flex-wrap gap-2">
                <Button variant="secondary" onClick={() => startPass('verify')} disabled={!!running}><ShieldCheck size={15} /> Verify credentials</Button>
                <Button variant="secondary" onClick={() => startPass('assessment')} disabled={!!running}><Gauge size={15} /> Run assessment</Button>
                <Button onClick={() => startPass('prestage')} disabled={!!running}><Layers size={15} /> Pre-stage migration</Button>
                <Button onClick={() => (hasAgent && agentOk ? fullMigrateViaAgent() : startPass('full'))} disabled={!!running}><PlayCircle size={15} /> Full migration{hasAgent && agentOk ? ' (agent)' : ''}</Button>
                <Button variant="secondary" onClick={() => startPass('delta')} disabled={!!running}><Zap size={15} /> Final delta</Button>
                <Button variant="danger" onClick={() => startPass('retry')} disabled={!!running}><RotateCcw size={15} /> Retry errors</Button>
                <Button variant="secondary" onClick={exportCsv} disabled={!items.length}><Download size={15} /> Export report</Button>
              </div>
              {running && <p className="mt-2 text-xs font-semibold text-blue-500 animate-pulse">Pass running: {running.toUpperCase()} …</p>}
            </Section>
          </Card>

          <LiveLog lines={log} />

          {/* Line items table */}
          <Card className="p-5">
            <Section title={`Line items (${items.length})`} action={items.length ? (
              <button onClick={() => { setState((s) => ({ ...s, items: [] })); addLog('All line items removed.', 'warn'); }} className="text-xs text-slate-400 hover:text-red-500 flex items-center gap-1"><Trash2 size={13} /> Clear all</button>
            ) : undefined}>
              {!items.length ? (
                <p className="text-sm text-slate-400">Import a CSV above to add mailboxes to this project.</p>
              ) : (
                <div className="overflow-x-auto">
                  <table className="w-full text-sm">
                    <thead>
                      <tr className="border-b border-slate-200 dark:border-slate-700 text-left text-xs uppercase text-slate-400">
                        <th className="py-2 pr-3">Source → Destination</th><th className="py-2 pr-3">Status</th><th className="py-2 pr-3 w-40">Progress</th><th className="py-2 pr-3">Data</th><th className="py-2 pr-3">Items</th><th className="py-2">Error</th>
                      </tr>
                    </thead>
                    <tbody>
                      {items.map((i) => (
                        <tr key={i.id} className="border-b border-slate-100 dark:border-slate-700/50 align-top">
                          <td className="py-2 pr-3">
                            <div className="font-medium text-slate-700 dark:text-slate-200">{i.sourceEmail}</div>
                            <div className="text-xs text-slate-400">→ {i.destEmail}</div>
                          </td>
                          <td className="py-2 pr-3"><Badge color={statusColor[i.status]}>{i.status}</Badge></td>
                          <td className="py-2 pr-3">
                            {i.progress > 0 ? <ProgressBar value={i.progress} /> : <ProgressBar value={i.mbTotal ? (i.mbMigrated / i.mbTotal) * 100 : 0} color={i.status === 'Completed' ? 'bg-emerald-500' : 'bg-slate-300 dark:bg-slate-600'} />}
                          </td>
                          <td className="py-2 pr-3 text-xs text-slate-500 dark:text-slate-300 whitespace-nowrap">{Math.round(i.mbMigrated)} / {Math.round(i.mbTotal)} MB</td>
                          <td className="py-2 pr-3 text-xs text-slate-500 dark:text-slate-300 whitespace-nowrap">{i.itemsMigrated.toLocaleString()} / {i.itemsTotal.toLocaleString()}{i.itemsFailed > 0 && <span className="text-red-500"> ({i.itemsFailed} failed)</span>}</td>
                          <td className="py-2 text-xs">
                            {i.error && (
                              <div className="max-w-xs">
                                <span className="text-red-500">{i.error}</span>
                                <div className="mt-1 flex gap-2">
                                  <Link to="/bittitan" className="text-blue-500 hover:underline">Error helper</Link>
                                  {!i.errorTransient && <button onClick={() => markFixed(i.id)} className="flex items-center gap-1 text-emerald-600 hover:underline"><Wrench size={11} /> Mark fixed</button>}
                                </div>
                              </div>
                            )}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </Section>
          </Card>
        </div>
      )}
    </div>
  );
}
