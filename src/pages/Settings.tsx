import { useEffect, useRef, useState } from 'react';
import {
  Settings as SettingsIcon, Plug, Loader2, CheckCircle2, XCircle, Trash2, LogIn, LogOut,
  Sparkles, KeyRound, Blocks, Palette, Database, Download, Upload, UserPlus, Users,
} from 'lucide-react';
import { Button, Card, Field, PageHeader, Section, Select, TextArea, Badge } from '../components/ui';
import { AISettings, AIProvider, SSOSettings, ServiceNowSettings } from '../types';
import {
  DEFAULT_SYSTEM_PROMPT, getAISettings, saveAISettings, getSSOSettings, saveSSOSettings,
  getServiceNowSettings, saveServiceNowSettings, getBranding, saveBranding, BrandingSettings,
  getIntegrations, saveIntegrations, IntegrationSettings,
} from '../store/settings';
import { testConnection } from '../services/ai';
import { signIn, signOut, currentAccount } from '../services/sso';
import { testServiceNow } from '../services/servicenow';
import { testJira, testZendesk, testTeamsWebhook, testSlackWebhook } from '../services/integrations';
import { addUser, changePassword, getSession, listUsers, removeUser } from '../services/auth';
import { load, save } from '../store/useLocalStorage';

type Tab = 'ai' | 'sso' | 'integrations' | 'branding' | 'account' | 'data';

const tabs: { id: Tab; label: string; icon: typeof Sparkles }[] = [
  { id: 'ai', label: 'AI Provider', icon: Sparkles },
  { id: 'sso', label: 'Microsoft SSO', icon: KeyRound },
  { id: 'integrations', label: 'Integrations', icon: Blocks },
  { id: 'branding', label: 'Branding', icon: Palette },
  { id: 'account', label: 'Account & Users', icon: Users },
  { id: 'data', label: 'Data', icon: Database },
];

function TestBadge({ state, msg }: { state: 'idle' | 'testing' | 'ok' | 'fail'; msg: string }) {
  return (
    <>
      {state === 'ok' && <span className="flex items-center gap-1.5 text-sm text-emerald-600"><CheckCircle2 size={15} /> Connected</span>}
      {state === 'fail' && <span className="flex items-center gap-1.5 text-sm text-red-500"><XCircle size={15} /> Failed</span>}
      {msg && <p className={`w-full text-xs ${state === 'ok' ? 'text-emerald-600' : 'text-red-500'}`}>{msg}</p>}
    </>
  );
}

function useTest(fn: () => Promise<string>) {
  const [state, setState] = useState<'idle' | 'testing' | 'ok' | 'fail'>('idle');
  const [msg, setMsg] = useState('');
  const run = async () => {
    setState('testing');
    setMsg('');
    try {
      setMsg(await fn());
      setState('ok');
    } catch (e) {
      setMsg(e instanceof Error ? e.message : String(e));
      setState('fail');
    }
  };
  return { state, msg, run };
}

export default function Settings() {
  const [tab, setTab] = useState<Tab>('ai');

  return (
    <div className="max-w-4xl">
      <PageHeader title="Settings" subtitle="AI provider, Microsoft SSO, PSA/ITSM integrations, branding, users and data management." icon={<SettingsIcon size={20} />} />
      <div className="mb-5 flex flex-wrap gap-2">
        {tabs.map((t) => (
          <button key={t.id} onClick={() => setTab(t.id)}
            className={`flex items-center gap-2 rounded-lg px-4 py-2 text-sm font-semibold ${tab === t.id ? 'bg-blue-600 text-white' : 'border border-slate-300 dark:border-slate-600 text-slate-600 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-700'}`}>
            <t.icon size={15} /> {t.label}
          </button>
        ))}
      </div>
      {tab === 'ai' && <AICard />}
      {tab === 'sso' && <SSOCard />}
      {tab === 'integrations' && <IntegrationsTab />}
      {tab === 'branding' && <BrandingCard />}
      {tab === 'account' && <AccountCard />}
      {tab === 'data' && <DataCard />}
    </div>
  );
}

// ---------------- AI ----------------
const providerInfo: Record<AIProvider, { label: string; modelHint: string; keyHint: string }> = {
  openrouter: { label: 'OpenRouter (recommended first)', modelHint: 'e.g. anthropic/claude-sonnet-4.5, openai/gpt-4o, or a :free model', keyHint: 'sk-or-v1-… (openrouter.ai/keys)' },
  openai: { label: 'OpenAI API', modelHint: 'e.g. gpt-4o, gpt-4o-mini', keyHint: 'sk-… (platform.openai.com)' },
  claude: { label: 'Claude API (Anthropic)', modelHint: 'e.g. claude-sonnet-4-5, claude-haiku-4-5', keyHint: 'sk-ant-… (console.anthropic.com)' },
  disabled: { label: 'Disabled (templates only)', modelHint: '', keyHint: '' },
};

function AICard() {
  const [s, setS] = useState<AISettings>(getAISettings());
  const [saved, setSaved] = useState(false);
  const t = useTest(async () => `Model replied: "${(await testConnection(s)).slice(0, 120)}"`);

  const set = <K extends keyof AISettings>(k: K, v: AISettings[K]) => { setS((p) => ({ ...p, [k]: v })); setSaved(false); };
  const doSave = () => { saveAISettings(s); setSaved(true); setTimeout(() => setSaved(false), 2000); };
  const info = providerInfo[s.provider];

  return (
    <Card className="p-5 space-y-4">
      <Section title="AI provider">
        <Select label="Provider" value={s.provider} onChange={(v) => set('provider', v as AIProvider)}
          options={(Object.keys(providerInfo) as AIProvider[]).map((p) => ({ value: p, label: providerInfo[p].label }))} />
        <p className="mt-2 text-xs text-slate-400">Switching provider only changes endpoint + auth header. Start with OpenRouter; move to OpenAI or Claude by selecting it and pasting that key.</p>
      </Section>
      {s.provider !== 'disabled' && (
        <>
          <Field label={`API key ${info.keyHint && `(${info.keyHint})`}`} type="password" value={s.apiKey} onChange={(v) => set('apiKey', v)} placeholder="Paste your API key" />
          <Field label={`Model name ${info.modelHint && `— ${info.modelHint}`}`} value={s.model} onChange={(v) => set('model', v)} />
          <label className="block">
            <span className="mb-1 block text-sm font-medium text-slate-700 dark:text-slate-300">Temperature: {s.temperature.toFixed(1)} (lower = more precise)</span>
            <input type="range" min={0} max={1} step={0.1} value={s.temperature} onChange={(e) => set('temperature', Number(e.target.value))} className="w-full accent-violet-600" />
          </label>
          <TextArea label="System prompt" value={s.systemPrompt} onChange={(v) => set('systemPrompt', v)} rows={6} />
          <button onClick={() => set('systemPrompt', DEFAULT_SYSTEM_PROMPT)} className="text-xs font-semibold text-blue-500 hover:underline">Reset to default system prompt</button>
        </>
      )}
      <div className="flex flex-wrap items-center gap-3 border-t border-slate-200 dark:border-slate-700 pt-4">
        <Button onClick={doSave}>{saved ? 'Saved ✓' : 'Save settings'}</Button>
        {s.provider !== 'disabled' && (
          <Button variant="ai" onClick={() => { saveAISettings(s); t.run(); }} disabled={t.state === 'testing' || !s.apiKey}>
            {t.state === 'testing' ? <Loader2 size={16} className="animate-spin" /> : <Plug size={16} />} Test AI connection
          </Button>
        )}
        <TestBadge state={t.state} msg={t.msg} />
      </div>
    </Card>
  );
}

// ---------------- SSO ----------------
function SSOCard() {
  const [s, setS] = useState<SSOSettings>(getSSOSettings());
  const [account, setAccount] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  const [msg, setMsg] = useState('');

  useEffect(() => { currentAccount().then((a) => setAccount(a ? `${a.name ?? a.username}` : null)); }, []);
  const set = <K extends keyof SSOSettings>(k: K, v: SSOSettings[K]) => { const next = { ...s, [k]: v }; setS(next); saveSSOSettings(next); };

  const doSignIn = async () => {
    setBusy(true); setMsg('');
    try { const a = await signIn(); setAccount(a.name ?? a.username); setMsg(`Signed in as ${a.username}`); }
    catch (e) { setMsg(e instanceof Error ? e.message : String(e)); }
    finally { setBusy(false); }
  };
  const doSignOut = async () => {
    setBusy(true);
    try { await signOut(); setAccount(null); setMsg('Signed out.'); }
    catch (e) { setMsg(e instanceof Error ? e.message : String(e)); }
    finally { setBusy(false); }
  };

  return (
    <Card className="p-5 space-y-3">
      <Section title="Microsoft 365 SSO (Entra ID)">
        <label className="flex items-center gap-2 text-sm font-medium text-slate-700 dark:text-slate-300">
          <input type="checkbox" checked={s.enabled} onChange={(e) => set('enabled', e.target.checked)} className="h-4 w-4 rounded border-slate-300 text-blue-600" />
          Enable Sign in with Microsoft (also shown on the login page)
        </label>
        {s.enabled && (
          <div className="mt-3 space-y-3">
            <Field label="Directory (tenant) ID" value={s.tenantId} onChange={(v) => set('tenantId', v)} placeholder="xxxxxxxx-xxxx-..." />
            <Field label="Application (client) ID" value={s.clientId} onChange={(v) => set('clientId', v)} placeholder="SPA app registration in your tenant" />
            <p className="text-xs text-slate-400">
              Setup: Entra admin center → App registrations → New → platform <strong>Single-page application</strong> with redirect URI = this portal's exact URL (e.g. https://sorrento.cloud/). Copy tenant + client ID here.
            </p>
            <div className="flex items-center gap-3">
              {account ? (
                <>
                  <span className="flex items-center gap-1.5 text-sm font-semibold text-emerald-600"><CheckCircle2 size={15} /> {account}</span>
                  <Button variant="secondary" onClick={doSignOut} disabled={busy}><LogOut size={15} /> Sign out</Button>
                </>
              ) : (
                <Button onClick={doSignIn} disabled={busy || !s.tenantId || !s.clientId}>
                  {busy ? <Loader2 size={15} className="animate-spin" /> : <LogIn size={15} />} Sign in with Microsoft
                </Button>
              )}
            </div>
            {msg && <p className="text-xs text-slate-500">{msg}</p>}
          </div>
        )}
      </Section>
    </Card>
  );
}

// ---------------- Integrations ----------------
function IntegrationsTab() {
  const [sn, setSn] = useState<ServiceNowSettings>(getServiceNowSettings());
  const [intg, setIntg] = useState<IntegrationSettings>(getIntegrations());
  const snTest = useTest(testServiceNow);
  const jiraTest = useTest(testJira);
  const zdTest = useTest(testZendesk);
  const teamsTest = useTest(testTeamsWebhook);
  const slackTest = useTest(testSlackWebhook);

  const setSnField = <K extends keyof ServiceNowSettings>(k: K, v: ServiceNowSettings[K]) => { const next = { ...sn, [k]: v }; setSn(next); saveServiceNowSettings(next); };
  const setI = <S extends keyof IntegrationSettings>(section: S, patch: Partial<IntegrationSettings[S]>) => {
    const next = { ...intg, [section]: { ...intg[section], ...patch } };
    setIntg(next);
    saveIntegrations(next);
  };

  return (
    <div className="space-y-5">
      {/* ServiceNow */}
      <Card className="p-5 space-y-3">
        <Section title="ServiceNow (ITSM)">
          <label className="flex items-center gap-2 text-sm font-medium text-slate-700 dark:text-slate-300">
            <input type="checkbox" checked={sn.enabled} onChange={(e) => setSnField('enabled', e.target.checked)} className="h-4 w-4 rounded border-slate-300 text-blue-600" />
            Enable — creates incidents straight from the Ticket Assistant
          </label>
          {sn.enabled && (
            <div className="mt-3 space-y-3">
              <Field label="Instance URL" value={sn.instanceUrl} onChange={(v) => setSnField('instanceUrl', v)} placeholder="https://yourinstance.service-now.com" />
              <div className="grid grid-cols-2 gap-3">
                <Field label="Username" value={sn.username} onChange={(v) => setSnField('username', v)} />
                <Field label="Password" type="password" value={sn.password} onChange={(v) => setSnField('password', v)} />
              </div>
              <div className="flex flex-wrap items-center gap-3">
                <Button variant="ai" onClick={snTest.run} disabled={snTest.state === 'testing' || !sn.instanceUrl}>
                  {snTest.state === 'testing' ? <Loader2 size={15} className="animate-spin" /> : <Plug size={15} />} Test connection
                </Button>
                <TestBadge state={snTest.state} msg={snTest.msg} />
              </div>
              <p className="text-xs text-amber-600 dark:text-amber-400">Instance needs a CORS rule for this portal's origin (System Web Services → REST → CORS Rules, Table API, GET/POST/PATCH).</p>
            </div>
          )}
        </Section>
      </Card>

      {/* Jira */}
      <Card className="p-5 space-y-3">
        <Section title="Jira Service Management">
          <label className="flex items-center gap-2 text-sm font-medium text-slate-700 dark:text-slate-300">
            <input type="checkbox" checked={intg.jira.enabled} onChange={(e) => setI('jira', { enabled: e.target.checked })} className="h-4 w-4 rounded border-slate-300 text-blue-600" />
            Enable Jira integration
          </label>
          {intg.jira.enabled && (
            <div className="mt-3 space-y-3">
              <Field label="Base URL" value={intg.jira.baseUrl} onChange={(v) => setI('jira', { baseUrl: v })} placeholder="https://yourcompany.atlassian.net" />
              <div className="grid grid-cols-3 gap-3">
                <Field label="Account email" value={intg.jira.email} onChange={(v) => setI('jira', { email: v })} />
                <Field label="API token" type="password" value={intg.jira.apiToken} onChange={(v) => setI('jira', { apiToken: v })} />
                <Field label="Project key" value={intg.jira.projectKey} onChange={(v) => setI('jira', { projectKey: v })} placeholder="ITSM" />
              </div>
              <div className="flex flex-wrap items-center gap-3">
                <Button variant="ai" onClick={jiraTest.run} disabled={jiraTest.state === 'testing'}>
                  {jiraTest.state === 'testing' ? <Loader2 size={15} className="animate-spin" /> : <Plug size={15} />} Test connection
                </Button>
                <TestBadge state={jiraTest.state} msg={jiraTest.msg} />
              </div>
            </div>
          )}
        </Section>
      </Card>

      {/* Zendesk */}
      <Card className="p-5 space-y-3">
        <Section title="Zendesk">
          <label className="flex items-center gap-2 text-sm font-medium text-slate-700 dark:text-slate-300">
            <input type="checkbox" checked={intg.zendesk.enabled} onChange={(e) => setI('zendesk', { enabled: e.target.checked })} className="h-4 w-4 rounded border-slate-300 text-blue-600" />
            Enable Zendesk integration
          </label>
          {intg.zendesk.enabled && (
            <div className="mt-3 space-y-3">
              <div className="grid grid-cols-3 gap-3">
                <Field label="Subdomain" value={intg.zendesk.subdomain} onChange={(v) => setI('zendesk', { subdomain: v })} placeholder="yourcompany" />
                <Field label="Agent email" value={intg.zendesk.email} onChange={(v) => setI('zendesk', { email: v })} />
                <Field label="API token" type="password" value={intg.zendesk.apiToken} onChange={(v) => setI('zendesk', { apiToken: v })} />
              </div>
              <div className="flex flex-wrap items-center gap-3">
                <Button variant="ai" onClick={zdTest.run} disabled={zdTest.state === 'testing'}>
                  {zdTest.state === 'testing' ? <Loader2 size={15} className="animate-spin" /> : <Plug size={15} />} Test connection
                </Button>
                <TestBadge state={zdTest.state} msg={zdTest.msg} />
              </div>
            </div>
          )}
        </Section>
      </Card>

      {/* PSA / ITSM connectors */}
      <Card className="p-5 space-y-4">
        <Section title="PSA / ITSM connectors (MSP tooling)">
          <p className="mb-2 text-xs text-slate-400">Credentials are stored locally and used by ticket sync features; most PSA APIs require a backend proxy for browser calls — the config here plugs straight into that backend later.</p>

          {([
            { key: 'topdesk', title: 'TOPdesk', fields: [['baseUrl', 'Base URL', 'https://yourcompany.topdesk.net'], ['username', 'Operator username', ''], ['appPassword', 'Application password', '']] },
            { key: 'halopsa', title: 'HaloPSA', fields: [['baseUrl', 'Base URL', 'https://yourcompany.halopsa.com'], ['clientId', 'Client ID', ''], ['clientSecret', 'Client secret', '']] },
            { key: 'connectwise', title: 'ConnectWise Manage', fields: [['baseUrl', 'API URL', 'https://api-eu.myconnectwise.net'], ['companyId', 'Company ID', ''], ['publicKey', 'Public key', ''], ['privateKey', 'Private key', '']] },
            { key: 'autotask', title: 'Datto Autotask', fields: [['apiUser', 'API user', ''], ['secret', 'Secret', ''], ['integrationCode', 'Integration code', '']] },
            { key: 'freshservice', title: 'Freshservice', fields: [['domain', 'Domain', 'yourcompany.freshservice.com'], ['apiKey', 'API key', '']] },
          ] as { key: keyof IntegrationSettings; title: string; fields: [string, string, string][] }[]).map((c) => {
            const section = intg[c.key] as unknown as Record<string, string | boolean>;
            return (
              <div key={c.key} className="border-t border-slate-100 dark:border-slate-700/60 pt-3 first:border-t-0 first:pt-0">
                <label className="flex items-center gap-2 text-sm font-medium text-slate-700 dark:text-slate-300">
                  <input type="checkbox" checked={Boolean(section.enabled)}
                    onChange={(e) => setI(c.key, { enabled: e.target.checked } as Partial<IntegrationSettings[typeof c.key]>)}
                    className="h-4 w-4 rounded border-slate-300 text-blue-600" />
                  {c.title}
                </label>
                {Boolean(section.enabled) && (
                  <div className="mt-2 grid gap-3 sm:grid-cols-3">
                    {c.fields.map(([fk, label, ph]) => (
                      <Field key={fk} label={label} type={/secret|password|key/i.test(fk) ? 'password' : 'text'}
                        value={String(section[fk] ?? '')} placeholder={ph}
                        onChange={(v) => setI(c.key, { [fk]: v } as Partial<IntegrationSettings[typeof c.key]>)} />
                    ))}
                  </div>
                )}
              </div>
            );
          })}
        </Section>
      </Card>

      {/* Teams & Slack webhooks */}
      <Card className="p-5 space-y-3">
        <Section title="Notifications — Microsoft Teams & Slack webhooks">
          <label className="flex items-center gap-2 text-sm font-medium text-slate-700 dark:text-slate-300">
            <input type="checkbox" checked={intg.teamsWebhook.enabled} onChange={(e) => setI('teamsWebhook', { enabled: e.target.checked })} className="h-4 w-4 rounded border-slate-300 text-blue-600" />
            Microsoft Teams incoming webhook
          </label>
          {intg.teamsWebhook.enabled && (
            <div className="flex flex-wrap items-end gap-3">
              <div className="min-w-72 flex-1"><Field label="Webhook URL (Teams channel → Workflows/Connectors)" value={intg.teamsWebhook.url} onChange={(v) => setI('teamsWebhook', { url: v })} placeholder="https://..." /></div>
              <Button variant="ai" onClick={teamsTest.run} disabled={teamsTest.state === 'testing'}>
                {teamsTest.state === 'testing' ? <Loader2 size={15} className="animate-spin" /> : <Plug size={15} />} Send test
              </Button>
              <TestBadge state={teamsTest.state} msg={teamsTest.msg} />
            </div>
          )}
          <label className="mt-3 flex items-center gap-2 text-sm font-medium text-slate-700 dark:text-slate-300">
            <input type="checkbox" checked={intg.slackWebhook.enabled} onChange={(e) => setI('slackWebhook', { enabled: e.target.checked })} className="h-4 w-4 rounded border-slate-300 text-blue-600" />
            Slack incoming webhook
          </label>
          {intg.slackWebhook.enabled && (
            <div className="flex flex-wrap items-end gap-3">
              <div className="min-w-72 flex-1"><Field label="Webhook URL" value={intg.slackWebhook.url} onChange={(v) => setI('slackWebhook', { url: v })} placeholder="https://hooks.slack.com/services/..." /></div>
              <Button variant="ai" onClick={slackTest.run} disabled={slackTest.state === 'testing'}>
                {slackTest.state === 'testing' ? <Loader2 size={15} className="animate-spin" /> : <Plug size={15} />} Send test
              </Button>
              <TestBadge state={slackTest.state} msg={slackTest.msg} />
            </div>
          )}
        </Section>
      </Card>

      {/* BitTitan & Syskit */}
      <Card className="p-5 space-y-3">
        <Section title="Migration & governance tooling">
          <label className="flex items-center gap-2 text-sm font-medium text-slate-700 dark:text-slate-300">
            <input type="checkbox" checked={intg.bittitan.enabled} onChange={(e) => setI('bittitan', { enabled: e.target.checked })} className="h-4 w-4 rounded border-slate-300 text-blue-600" />
            BitTitan MigrationWiz <Badge color="purple">API key stored for backend use</Badge>
          </label>
          {intg.bittitan.enabled && (
            <Field label="MigrationWiz API key" type="password" value={intg.bittitan.apiKey} onChange={(v) => setI('bittitan', { apiKey: v })} />
          )}
          <label className="mt-3 flex items-center gap-2 text-sm font-medium text-slate-700 dark:text-slate-300">
            <input type="checkbox" checked={intg.syskit.enabled} onChange={(e) => setI('syskit', { enabled: e.target.checked })} className="h-4 w-4 rounded border-slate-300 text-blue-600" />
            Syskit Points
          </label>
          {intg.syskit.enabled && (
            <Field label="Syskit base URL (used for deep links from the Syskit module)" value={intg.syskit.baseUrl} onChange={(v) => setI('syskit', { baseUrl: v })} placeholder="https://yourcompany.syskit365.com" />
          )}
        </Section>
      </Card>
    </div>
  );
}

// ---------------- Branding ----------------
function BrandingCard() {
  const [b, setB] = useState<BrandingSettings>(getBranding());
  const [saved, setSaved] = useState(false);
  const set = <K extends keyof BrandingSettings>(k: K, v: BrandingSettings[K]) => { setB((p) => ({ ...p, [k]: v })); setSaved(false); };

  return (
    <Card className="p-5 space-y-3">
      <Section title="Portal branding">
        <div className="grid gap-3 sm:grid-cols-2">
          <Field label="Company name" value={b.companyName} onChange={(v) => set('companyName', v)} placeholder="Sorrento Cloud" />
          <Field label="Portal name" value={b.portalName} onChange={(v) => set('portalName', v)} placeholder="M365 WorkPilot" />
          <Field label="Support email (signatures, login page)" value={b.supportEmail} onChange={(v) => set('supportEmail', v)} />
          <Select label="Default output language" value={b.defaultLanguage} onChange={(v) => set('defaultLanguage', v as 'en' | 'nl')}
            options={[{ value: 'en', label: 'English' }, { value: 'nl', label: 'Dutch (Nederlands)' }]} />
        </div>
        <div className="mt-4 flex items-center gap-3">
          <Button onClick={() => { saveBranding(b); setSaved(true); setTimeout(() => location.reload(), 600); }}>{saved ? 'Saved — reloading…' : 'Save branding'}</Button>
          <span className="text-xs text-slate-400">Shown in the sidebar, login page and generated emails.</span>
        </div>
      </Section>
    </Card>
  );
}

// ---------------- Account & users ----------------
function AccountCard() {
  const session = getSession();
  const [users, setUsers] = useState<string[]>(listUsers());
  const [cur, setCur] = useState('');
  const [next, setNext] = useState('');
  const [newEmail, setNewEmail] = useState('');
  const [newPass, setNewPass] = useState('');
  const [msg, setMsg] = useState('');

  const doChange = async () => {
    setMsg('');
    try {
      await changePassword(session?.email ?? '', cur, next);
      setMsg('Password changed.');
      setCur(''); setNext('');
    } catch (e) { setMsg(e instanceof Error ? e.message : String(e)); }
  };

  const doAdd = async () => {
    setMsg('');
    try {
      await addUser(newEmail, newPass);
      setUsers(listUsers());
      setMsg(`User ${newEmail} added.`);
      setNewEmail(''); setNewPass('');
    } catch (e) { setMsg(e instanceof Error ? e.message : String(e)); }
  };

  return (
    <div className="space-y-5">
      <Card className="p-5">
        <Section title={`Change password — ${session?.email ?? ''}`}>
          <div className="grid gap-3 sm:grid-cols-2">
            <Field label="Current password" type="password" value={cur} onChange={setCur} />
            <Field label="New password (min 6 chars)" type="password" value={next} onChange={setNext} />
          </div>
          <div className="mt-3"><Button onClick={doChange} disabled={!cur || !next}>Change password</Button></div>
        </Section>
      </Card>
      <Card className="p-5">
        <Section title="Local portal users">
          <div className="mb-3 space-y-1.5">
            {users.map((u) => (
              <div key={u} className="flex items-center gap-2 rounded-lg border border-slate-200 dark:border-slate-700 px-3 py-2 text-sm text-slate-700 dark:text-slate-200">
                {u}
                {u !== session?.email && users.length > 1 && (
                  <button onClick={() => { try { removeUser(u); setUsers(listUsers()); } catch (e) { setMsg(String(e)); } }} className="ml-auto text-slate-400 hover:text-red-500"><Trash2 size={14} /></button>
                )}
              </div>
            ))}
          </div>
          <div className="grid gap-3 sm:grid-cols-2">
            <Field label="New user email" value={newEmail} onChange={setNewEmail} placeholder="engineer@sorrento.cloud" />
            <Field label="Password" type="password" value={newPass} onChange={setNewPass} />
          </div>
          <div className="mt-3"><Button variant="secondary" onClick={doAdd} disabled={!newEmail || !newPass}><UserPlus size={15} /> Add user</Button></div>
          {msg && <p className="mt-2 text-xs text-slate-500">{msg}</p>}
          <p className="mt-3 text-xs text-amber-600 dark:text-amber-400">Local accounts are a usability gate in this browser-only version, not a hard security boundary — use Microsoft 365 SSO for enterprise authentication.</p>
        </Section>
      </Card>
    </div>
  );
}

// ---------------- Data ----------------
function DataCard() {
  const fileRef = useRef<HTMLInputElement>(null);
  const [msg, setMsg] = useState('');

  const exportData = () => {
    const data: Record<string, unknown> = {};
    Object.keys(localStorage).filter((k) => k.startsWith('workpilot:')).forEach((k) => {
      try { data[k] = JSON.parse(localStorage.getItem(k) ?? 'null'); } catch { data[k] = localStorage.getItem(k); }
    });
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `workpilot-backup-${new Date().toISOString().slice(0, 10)}.json`;
    a.click();
    URL.revokeObjectURL(a.href);
  };

  const importData = (file: File) => {
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const data = JSON.parse(String(reader.result));
        Object.entries(data).forEach(([k, v]) => {
          if (k.startsWith('workpilot:')) localStorage.setItem(k, JSON.stringify(v));
        });
        setMsg('Backup imported — reloading…');
        setTimeout(() => location.reload(), 800);
      } catch (e) {
        setMsg(`Import failed: ${e instanceof Error ? e.message : e}`);
      }
    };
    reader.readAsText(file);
  };

  const clearData = () => {
    if (confirm('Clear ALL local portal data (tickets, projects, notes, settings, users)? This cannot be undone.')) {
      Object.keys(localStorage).filter((k) => k.startsWith('workpilot:')).forEach((k) => localStorage.removeItem(k));
      location.reload();
    }
  };

  return (
    <Card className="p-5">
      <Section title="Local data">
        <p className="mb-4 text-sm text-slate-500 dark:text-slate-400">
          All portal data lives in this browser's localStorage under the <code className="rounded bg-slate-100 dark:bg-slate-700 px-1">workpilot:</code> prefix.
          Export a backup before switching machines or hosting; the storage layer (src/store) is a single module ready to swap for a real database.
        </p>
        <div className="flex flex-wrap gap-3">
          <Button variant="secondary" onClick={exportData}><Download size={15} /> Export backup (JSON)</Button>
          <Button variant="secondary" onClick={() => fileRef.current?.click()}><Upload size={15} /> Import backup</Button>
          <Button variant="danger" onClick={clearData}><Trash2 size={15} /> Clear all local data</Button>
          <input ref={fileRef} type="file" accept="application/json" className="hidden" onChange={(e) => e.target.files?.[0] && importData(e.target.files[0])} />
        </div>
        {msg && <p className="mt-3 text-xs text-slate-500">{msg}</p>}
      </Section>
    </Card>
  );
}
