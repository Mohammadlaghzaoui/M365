import { useEffect, useState } from 'react';
import { Cpu, Download, CheckCircle2, XCircle, Loader2, Plug, ShieldCheck, Terminal, Cloud, RefreshCw } from 'lucide-react';
import { Button, Card, Field, PageHeader, Section, Badge } from '../components/ui';
import { getIntegrations, saveIntegrations } from '../store/settings';
import { agentHealth, AgentHealth } from '../services/agent';

export default function Agent() {
  const [intg, setIntg] = useState(getIntegrations());
  const [state, setState] = useState<'idle' | 'testing' | 'ok' | 'fail'>('idle');
  const [health, setHealth] = useState<AgentHealth | null>(null);
  const [msg, setMsg] = useState('');

  const agent = intg.agent;
  const setAgent = (patch: Partial<typeof agent>) => {
    const next = { ...intg, agent: { ...intg.agent, ...patch } };
    setIntg(next);
    saveIntegrations(next);
  };

  const test = async () => {
    setState('testing'); setMsg('');
    try {
      const h = await agentHealth();
      setHealth(h);
      setState('ok');
    } catch (e) {
      setState('fail');
      setMsg(e instanceof Error ? e.message : String(e));
    }
  };

  useEffect(() => {
    if (agent.enabled && agent.url && agent.apiKey) test();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  return (
    <div className="max-w-4xl">
      <PageHeader title="Migration Agent" subtitle="Download the agent, run it on your PC, and the portal executes real migrations and provisioning through it." icon={<Cpu size={20} />} />

      {/* Status */}
      <Card className={`mb-5 p-5 ${state === 'ok' ? 'border-emerald-300 dark:border-emerald-700' : state === 'fail' ? 'border-red-300 dark:border-red-700' : ''}`}>
        <div className="flex flex-wrap items-center gap-3">
          <span className={`rounded-lg p-2.5 ${state === 'ok' ? 'bg-emerald-100 dark:bg-emerald-900/40 text-emerald-600' : 'bg-violet-100 dark:bg-violet-900/40 text-violet-600'}`}><Cpu size={20} /></span>
          <div className="flex-1 min-w-64">
            {state === 'ok' && health ? (
              <>
                <div className="flex items-center gap-2 text-sm font-semibold text-emerald-700 dark:text-emerald-300"><CheckCircle2 size={16} /> Agent connected — real execution enabled</div>
                <div className="text-xs text-slate-500 dark:text-slate-400">{health.name} v{health.version} on {health.host} · PowerShell {health.capabilities.powershell ? '✓' : '✗'} · Graph {health.capabilities.graph ? '✓' : '✗'} · modules: {health.capabilities.modules.join(', ') || '—'}</div>
              </>
            ) : state === 'fail' ? (
              <>
                <div className="flex items-center gap-2 text-sm font-semibold text-red-600 dark:text-red-400"><XCircle size={16} /> Not connected</div>
                <div className="text-xs text-slate-500 dark:text-slate-400">{msg || 'Start the agent on your PC and check the URL/key below.'}</div>
              </>
            ) : (
              <div className="text-sm text-slate-500 dark:text-slate-400">Agent not connected yet — follow the steps below.</div>
            )}
          </div>
          <Badge color={state === 'ok' ? 'green' : 'purple'}>{state === 'ok' ? 'LIVE' : 'OFFLINE'}</Badge>
          <Button variant="secondary" onClick={test} disabled={state === 'testing'}>
            {state === 'testing' ? <Loader2 size={15} className="animate-spin" /> : <RefreshCw size={15} />} Test connection
          </Button>
        </div>
      </Card>

      {/* Download + steps */}
      <Card className="p-5 mb-5">
        <Section title="1. Download & run the agent on your PC">
          <div className="flex flex-wrap items-center gap-3">
            <a href="./downloads/workpilot-agent.zip" download>
              <Button><Download size={16} /> Download agent (.zip)</Button>
            </a>
            <span className="text-xs text-slate-400">~26 KB · Windows · requires Node.js 18+</span>
          </div>
          <ol className="mt-4 space-y-2.5">
            {[
              <>Install <a className="text-blue-500 hover:underline" href="https://nodejs.org" target="_blank" rel="noreferrer">Node.js 18+</a> (the "LTS" download) — one time only.</>,
              <>Unzip the downloaded file somewhere on your PC.</>,
              <>Double-click <code className="rounded bg-slate-100 dark:bg-slate-700 px-1">start.bat</code>. The first run installs the components (wait a moment), then the agent runs on <code className="rounded bg-slate-100 dark:bg-slate-700 px-1">http://localhost:8787</code>.</>,
              <>Keep that black window open — closing it stops the agent.</>,
            ].map((t, i) => (
              <li key={i} className="flex items-start gap-3 text-sm text-slate-700 dark:text-slate-200">
                <span className="flex h-6 w-6 shrink-0 items-center justify-center rounded-full bg-blue-100 dark:bg-blue-900/50 text-xs font-bold text-blue-700 dark:text-blue-300">{i + 1}</span>
                <span className="pt-0.5">{t}</span>
              </li>
            ))}
          </ol>
        </Section>
      </Card>

      {/* Connect */}
      <Card className="p-5 mb-5">
        <Section title="2. Connect this portal to your agent">
          <label className="flex items-center gap-2 text-sm font-medium text-slate-700 dark:text-slate-300">
            <input type="checkbox" checked={agent.enabled} onChange={(e) => setAgent({ enabled: e.target.checked })} className="h-4 w-4 rounded border-slate-300 text-blue-600" />
            Enable the Migration Agent
          </label>
          {agent.enabled && (
            <div className="mt-3 space-y-3">
              <div className="grid gap-3 sm:grid-cols-2">
                <Field label="Agent URL" value={agent.url} onChange={(v) => setAgent({ url: v })} placeholder="http://localhost:8787" />
                <Field label="API key" type="password" value={agent.apiKey} onChange={(v) => setAgent({ apiKey: v })} placeholder="workpilot-local-key" />
              </div>
              <div className="flex flex-wrap items-center gap-3">
                <Button variant="ai" onClick={test} disabled={state === 'testing' || !agent.url}>
                  {state === 'testing' ? <Loader2 size={15} className="animate-spin" /> : <Plug size={15} />} Test agent connection
                </Button>
                {state === 'ok' && <span className="flex items-center gap-1.5 text-sm text-emerald-600"><CheckCircle2 size={15} /> Connected</span>}
                {state === 'fail' && <span className="text-sm text-red-500">{msg}</span>}
              </div>
              <p className="text-xs text-slate-400">
                Defaults match the agent: URL <code className="rounded bg-slate-100 dark:bg-slate-700 px-1">http://localhost:8787</code>, key <code className="rounded bg-slate-100 dark:bg-slate-700 px-1">workpilot-local-key</code> (change it in the agent's <code>.env</code> for production). Modern browsers allow this HTTPS portal to reach <code>localhost</code>.
              </p>
            </div>
          )}
        </Section>
      </Card>

      {/* What it does */}
      <Card className="p-5">
        <Section title="What the agent does once connected">
          <div className="grid gap-4 sm:grid-cols-3">
            <div className="rounded-lg border border-slate-200 dark:border-slate-700 p-3">
              <ShieldCheck size={18} className="mb-1.5 text-blue-500" />
              <div className="text-sm font-semibold text-slate-700 dark:text-slate-200">Endpoint verify</div>
              <div className="text-xs text-slate-500 dark:text-slate-400">Real connectivity checks for Graph, on-prem AD and Exchange.</div>
            </div>
            <div className="rounded-lg border border-slate-200 dark:border-slate-700 p-3">
              <Terminal size={18} className="mb-1.5 text-violet-500" />
              <div className="text-sm font-semibold text-slate-700 dark:text-slate-200">Provisioning</div>
              <div className="text-xs text-slate-500 dark:text-slate-400">New-ADUser on-prem + Entra Connect sync, or Graph cloud accounts & invitations.</div>
            </div>
            <div className="rounded-lg border border-slate-200 dark:border-slate-700 p-3">
              <Cloud size={18} className="mb-1.5 text-emerald-500" />
              <div className="text-sm font-semibold text-slate-700 dark:text-slate-200">Migration</div>
              <div className="text-xs text-slate-500 dark:text-slate-400">Submits real Exchange migration batches; the Console streams the live log.</div>
            </div>
          </div>
          <p className="mt-3 text-xs text-amber-600 dark:text-amber-400">
            On-prem AD/Exchange actions only work if the PC running the agent has access (domain-joined, the right PowerShell modules, a delegated account). Cloud actions work once you fill TENANT_ID/CLIENT_ID/CLIENT_SECRET in the agent's .env.
          </p>
        </Section>
      </Card>
    </div>
  );
}
