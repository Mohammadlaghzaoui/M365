import { useEffect, useRef, useState } from 'react';
import { useNavigate, Link } from 'react-router-dom';
import { ArrowLeft, Building2, ShieldCheck, Loader2, Copy, Check, ExternalLink, Terminal, KeyRound } from 'lucide-react';
import { runDiscovery, useDiscoveryToken, LogLevel, DISCOVERY_SCOPES } from '../services/graphDiscovery';
import { requestDeviceCode, pollForToken, getOnecomToken, clearOnecomToken, DeviceCode } from '../services/onecomDeviceAuth';
import { connectTenant, getDiscoveryToken, discoveryAuthConfigured, getDiscoveryAuth, saveDiscoveryAuth } from '../services/discoveryAuth';
import { cloudAgentConfigured, startDeviceDiscovery, trackDeviceDiscovery, DeviceSession } from '../services/cloudDiscovery';
import { saveTenantResult } from '../services/tenantStore';
import { save } from '../store/useLocalStorage';

type Line = { text: string; level: LogLevel };

function CopyBtn({ text }: { text: string }) {
  const [done, setDone] = useState(false);
  return (
    <button onClick={() => { navigator.clipboard?.writeText(text); setDone(true); setTimeout(() => setDone(false), 1200); }}
      className="rounded border border-slate-300 p-1.5 text-slate-500 hover:bg-slate-100 dark:border-slate-600 dark:hover:bg-slate-800">
      {done ? <Check size={14} className="text-emerald-500" /> : <Copy size={14} />}
    </button>
  );
}

export default function SourceTenantConnect() {
  const nav = useNavigate();
  const [phase, setPhase] = useState<'idle' | 'awaiting' | 'collecting' | 'error'>('idle');
  const [code, setCode] = useState<DeviceCode | null>(null);
  const [lines, setLines] = useState<Line[]>([]);
  const [error, setError] = useState('');
  const [showCode, setShowCode] = useState(false);
  const [showAdvanced, setShowAdvanced] = useState(false);
  const [clientId, setClientId] = useState(() => getDiscoveryAuth().clientId);
  const [configured, setConfigured] = useState(() => discoveryAuthConfigured());
  const [agentSession, setAgentSession] = useState<DeviceSession | null>(null);
  const agentReady = cloudAgentConfigured();
  const logRef = useRef<HTMLDivElement>(null);
  useEffect(() => { logRef.current?.scrollTo({ top: logRef.current.scrollHeight }); }, [lines]);

  const redirectUri = window.location.origin + window.location.pathname;
  const addLine = (text: string, level: LogLevel = 'info') =>
    setLines((p) => [...p, { text: level === 'cmd' ? text : `[${new Date().toLocaleTimeString('en-GB')}] ${text}`, level }]);

  const finish = (r: Awaited<ReturnType<typeof runDiscovery>>) => {
    addLine('Read-only assessment complete. No tenant changes were made.', 'ok');
    save('discovery-result', r); saveTenantResult(r);
    addLine('Opening the Source Tenant dashboard …', 'ok');
    setTimeout(() => nav(`/tenants/${r.org.tenantId}`), 600);
  };

  // ---- Local agent: device code requested from YOUR PC (your location, not Denmark) ----
  const agentConnect = async () => {
    setError(''); setLines([]); setAgentSession(null);
    try {
      addLine('Asking your local agent (on this PC) for a sign-in code …', 'info');
      const s = await startDeviceDiscovery();
      setAgentSession(s); setPhase('awaiting');
      const done = await trackDeviceDiscovery(s.id, (text, level) => addLine(text, level));
      setAgentSession(done);
      if (done.phase === 'done' && done.result) { setPhase('collecting'); finish(done.result); }
      else { setError(done.error || 'The agent could not complete the sign-in.'); setPhase('error'); }
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      const friendly = /failed to fetch|networkerror|load failed|typeerror/i.test(msg)
        ? 'The local agent is not running on this PC (or its URL in Settings is wrong). Start the agent first, or use "Sign in with Microsoft" above — that needs no agent.'
        : msg;
      setError(friendly);
      addLine(`ERROR: ${friendly}`, 'err');
      setPhase('error');
    }
  };

  // ---- Recommended: full browser sign-in (your location, no code) ----
  const signInPopup = async () => {
    setError(''); setLines([]); setPhase('collecting');
    try {
      addLine('Opening the Microsoft sign-in window (in your browser) …', 'info');
      const { tenantId, username } = await connectTenant(DISCOVERY_SCOPES);
      addLine(`Signed in as ${username} (tenant ${tenantId}). Starting read-only discovery …`, 'ok');
      useDiscoveryToken(getDiscoveryToken);
      finish(await runDiscovery(addLine));
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
      addLine(`ERROR: ${e instanceof Error ? e.message : e}`, 'err');
      setPhase('error');
    }
  };

  const saveClient = () => {
    const id = clientId.trim();
    saveDiscoveryAuth({ clientId: id });
    setConfigured(!!id);
  };

  // ---- Fallback: device code via the one.com broker ----
  const startCode = async () => {
    setError(''); setLines([]); setCode(null);
    try {
      addLine('Requesting a device sign-in code from Microsoft …', 'info');
      const dc = await requestDeviceCode();
      setCode(dc); setPhase('awaiting');
      const { tenantId } = await pollForToken(dc);
      setPhase('collecting');
      addLine(`Signed in (tenant ${tenantId}). Starting read-only discovery …`, 'ok');
      useDiscoveryToken(getOnecomToken);
      finish(await runDiscovery(addLine));
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
      addLine(`ERROR: ${e instanceof Error ? e.message : e}`, 'err');
      setPhase('error');
    } finally {
      clearOnecomToken();
    }
  };

  const lineColor: Record<LogLevel, string> = { cmd: 'text-sky-300', info: 'text-slate-300', ok: 'text-emerald-400', warn: 'text-amber-300', err: 'text-rose-400' };
  const busy = phase === 'collecting' || phase === 'awaiting';

  return (
    <div className="mx-auto max-w-[1100px]">
      <Link to="/tenants" className="mb-3 inline-flex items-center gap-1 text-sm text-slate-500 hover:text-blue-700"><ArrowLeft size={14} /> Source Tenants</Link>

      <div className="mb-5 flex flex-wrap items-center gap-4 rounded-lg border border-slate-300 bg-white p-5 shadow-sm dark:border-slate-700 dark:bg-slate-900">
        <div className="flex h-12 w-12 items-center justify-center rounded-md bg-slate-800 text-white"><Building2 size={22} /></div>
        <div className="mr-auto">
          <h1 className="text-xl font-bold text-slate-900 dark:text-slate-100">Add a source tenant</h1>
          <p className="text-sm text-slate-500">Sign in to the customer's tenant — the read-only Microsoft Graph engine runs here and opens the tenant's dashboard.</p>
        </div>
        <span className="inline-flex items-center gap-1 rounded bg-emerald-50 px-2.5 py-1 text-xs font-medium text-emerald-700 dark:bg-emerald-900/30 dark:text-emerald-300"><ShieldCheck size={13} /> Read-only — no tenant changes</span>
      </div>

      <div className="grid grid-cols-1 gap-5 lg:grid-cols-5">
        <div className="lg:col-span-2 space-y-5">
          {/* Browser sign-in — your location (Belgium), no Denmark, no install */}
          <div className="rounded-lg border-2 border-blue-400 bg-white p-5 shadow-sm dark:border-blue-700 dark:bg-slate-900">
            <div className="mb-1 flex items-center gap-2">
              <h2 className="text-base font-semibold text-slate-900 dark:text-slate-100">Sign in with Microsoft</h2>
              <span className="rounded bg-blue-50 px-2 py-0.5 text-[11px] font-semibold text-blue-700 dark:bg-blue-900/30 dark:text-blue-300">Recommended · YOUR location</span>
            </div>
            <p className="mb-4 text-sm text-slate-500">A normal Microsoft sign-in window opens in your browser — no code, no agent, nothing installed. The sign-in shows <strong>your own location</strong> (not Denmark). This is the original popup flow.</p>

            {configured ? (
              <button onClick={signInPopup} disabled={busy} className="inline-flex w-full items-center justify-center gap-2 rounded-md bg-blue-700 px-4 py-2.5 text-sm font-semibold text-white hover:bg-blue-800 disabled:opacity-50">
                {busy ? <Loader2 size={16} className="animate-spin" /> : <Building2 size={16} />} Sign in &amp; analyze
              </button>
            ) : (
              <div className="space-y-3 rounded-md border border-slate-200 bg-slate-50 p-3 text-sm dark:border-slate-700 dark:bg-slate-800/40">
                <div className="flex items-center gap-1.5 font-medium text-slate-700 dark:text-slate-200"><KeyRound size={14} /> One-time setup (~2 min, once ever)</div>
                <ol className="ml-4 list-decimal space-y-1.5 text-xs text-slate-600 dark:text-slate-300">
                  <li>
                    Create a free app registration in <strong>your</strong> tenant (multitenant):
                    <a href="https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/CreateApplicationBlade/quickStartType~/null/isMSAApp~/false" target="_blank" rel="noreferrer" className="ml-1 inline-flex items-center gap-1 font-medium text-blue-700 hover:underline dark:text-blue-400">open Entra <ExternalLink size={11} /></a>
                  </li>
                  <li>Add a <strong>Single-page application</strong> redirect URI:
                    <div className="mt-1 flex items-center gap-2"><code className="flex-1 break-all rounded bg-white px-2 py-1 text-[11px] dark:bg-slate-900">{redirectUri}</code><CopyBtn text={redirectUri} /></div>
                  </li>
                  <li>Paste the <strong>Application (client) ID</strong> here:</li>
                </ol>
                <div className="flex items-center gap-2">
                  <input value={clientId} onChange={(e) => setClientId(e.target.value)} placeholder="00000000-0000-0000-0000-000000000000" className="flex-1 rounded border border-slate-300 bg-white px-2 py-1.5 font-mono text-xs dark:border-slate-600 dark:bg-slate-900 dark:text-slate-100" />
                  <button onClick={saveClient} disabled={!clientId.trim()} className="rounded-md bg-blue-700 px-3 py-1.5 text-xs font-semibold text-white hover:bg-blue-800 disabled:opacity-50">Save</button>
                </div>
                <p className="text-[11px] text-slate-400">No client secret. The customer admin consents to the read-only scopes the first time they sign in.</p>
              </div>
            )}
            {configured && (
              <button onClick={() => setConfigured(false)} className="mt-2 text-xs text-slate-400 hover:text-slate-600 dark:hover:text-slate-300">Change app registration</button>
            )}
          </div>

          {/* Other ways to sign in (advanced) */}
          <button onClick={() => setShowAdvanced((s) => !s)} className="w-full rounded-lg border border-slate-300 bg-white px-4 py-3 text-left text-sm font-medium text-slate-600 shadow-sm hover:bg-slate-50 dark:border-slate-700 dark:bg-slate-900 dark:text-slate-300 dark:hover:bg-slate-800">
            {showAdvanced ? '▾' : '▸'} Other ways to sign in (code / local agent)
          </button>
          {showAdvanced && <>
          {/* Local agent: code from your own PC */}
          <div className="rounded-lg border border-slate-300 bg-white p-5 shadow-sm dark:border-slate-700 dark:bg-slate-900">
            <div className="mb-1 flex items-center gap-2">
              <h2 className="text-sm font-semibold text-slate-700 dark:text-slate-200">Connect with a code — local agent</h2>
              <span className="rounded bg-slate-100 px-2 py-0.5 text-[11px] font-medium text-slate-500 dark:bg-slate-700/50 dark:text-slate-300">Your PC's location</span>
            </div>
            <p className="mb-3 text-xs text-slate-500">Prefer a code over the popup? Run the agent on this PC so the code is requested locally — Microsoft then shows your location, not Denmark.</p>
            {agentReady ? (
              agentSession && agentSession.phase === 'awaiting-auth' ? (
                <ol className="space-y-3 text-sm">
                  <li><div className="mb-1 font-medium text-slate-700 dark:text-slate-200">1 · Open</div><a href={agentSession.verificationUri} target="_blank" rel="noreferrer" className="inline-flex items-center gap-1.5 text-blue-700 hover:underline dark:text-blue-400">{agentSession.verificationUri || 'microsoft.com/devicelogin'} <ExternalLink size={12} /></a></li>
                  <li><div className="mb-1 font-medium text-slate-700 dark:text-slate-200">2 · Enter code</div><div className="flex items-center gap-2"><span className="rounded border border-slate-300 bg-slate-50 px-3 py-1.5 font-mono text-lg font-bold tracking-widest dark:border-slate-600 dark:bg-slate-800 dark:text-slate-100">{agentSession.userCode}</span><CopyBtn text={agentSession.userCode} /></div></li>
                  <li className="flex items-center gap-2 text-xs text-blue-600"><Loader2 size={13} className="animate-spin" /> Waiting for sign-in…</li>
                </ol>
              ) : (
                <button onClick={agentConnect} disabled={busy} className="inline-flex items-center gap-2 rounded-md border border-slate-300 px-3 py-2 text-sm font-medium text-slate-700 hover:bg-slate-100 disabled:opacity-50 dark:border-slate-600 dark:text-slate-200 dark:hover:bg-slate-800">
                  {busy ? <Loader2 size={15} className="animate-spin" /> : <Terminal size={15} />} Get a code from this PC
                </button>
              )
            ) : (
              <div className="flex flex-wrap gap-2">
                <Link to="/agent" className="rounded-md border border-slate-300 px-3 py-1.5 text-xs font-medium text-slate-700 hover:bg-slate-100 dark:border-slate-600 dark:text-slate-200 dark:hover:bg-slate-800">Download the agent</Link>
                <Link to="/settings" className="rounded-md border border-slate-300 px-3 py-1.5 text-xs font-medium text-slate-700 hover:bg-slate-100 dark:border-slate-600 dark:text-slate-200 dark:hover:bg-slate-800">Configure (Settings)</Link>
              </div>
            )}
          </div>

          {/* Fallback: device code */}
          <div className="rounded-lg border border-slate-300 bg-white p-5 shadow-sm dark:border-slate-700 dark:bg-slate-900">
            <button onClick={() => setShowCode((s) => !s)} className="flex w-full items-center justify-between text-left">
              <span className="text-sm font-semibold text-slate-700 dark:text-slate-200">Or sign in with a code</span>
              <span className="text-xs text-slate-400">{showCode ? 'hide' : 'show'}</span>
            </button>
            {showCode && (
              <div className="mt-3">
                <p className="mb-3 text-xs text-slate-500">No app registration. Run the small <strong>local sign-in helper</strong> on your PC first (double-click START-local-signin.bat) and the sign-in comes from <strong>your location (Belgium)</strong>. Without it running, it uses the server (Denmark).</p>
                {phase === 'awaiting' && code ? (
                  <ol className="space-y-3 text-sm">
                    <li><div className="mb-1 font-medium text-slate-700 dark:text-slate-200">1 · Open</div><a href={code.verification_uri} target="_blank" rel="noreferrer" className="inline-flex items-center gap-1.5 text-blue-700 hover:underline dark:text-blue-400">{code.verification_uri || 'microsoft.com/devicelogin'} <ExternalLink size={12} /></a></li>
                    <li><div className="mb-1 font-medium text-slate-700 dark:text-slate-200">2 · Enter code</div><div className="flex items-center gap-2"><span className="rounded border border-slate-300 bg-slate-50 px-3 py-1.5 font-mono text-lg font-bold tracking-widest dark:border-slate-600 dark:bg-slate-800 dark:text-slate-100">{code.user_code}</span><CopyBtn text={code.user_code} /></div></li>
                    <li className="flex items-center gap-2 text-xs text-blue-600"><Loader2 size={13} className="animate-spin" /> Waiting for sign-in…</li>
                  </ol>
                ) : (
                  <button onClick={startCode} disabled={busy} className="inline-flex items-center gap-2 rounded-md border border-slate-300 px-3 py-2 text-sm font-medium text-slate-700 hover:bg-slate-100 disabled:opacity-50 dark:border-slate-600 dark:text-slate-200 dark:hover:bg-slate-800">
                    {busy ? <Loader2 size={15} className="animate-spin" /> : <Terminal size={15} />} Get a code
                  </button>
                )}
              </div>
            )}
          </div>
          </>}

          {error && <p className="rounded-md bg-rose-50 px-3 py-2 text-sm text-rose-700 dark:bg-rose-900/20 dark:text-rose-300">{error}</p>}
        </div>

        {/* Live console */}
        <div className="lg:col-span-3">
          <div className="overflow-hidden rounded-lg border border-slate-800 shadow-lg">
            <div className="flex items-center gap-2 border-b border-slate-800 bg-slate-900 px-4 py-2.5">
              <span className="h-3 w-3 rounded-full bg-rose-500" /><span className="h-3 w-3 rounded-full bg-amber-400" /><span className="h-3 w-3 rounded-full bg-emerald-500" />
              <span className="ml-2 flex items-center gap-1.5 text-xs font-medium text-slate-300"><Terminal size={13} /> Microsoft Graph — read-only discovery {busy && <span className="ml-1 animate-pulse text-emerald-400">· running</span>}</span>
            </div>
            <div ref={logRef} className="h-[460px] overflow-y-auto bg-[#0b1220] p-4 font-mono text-[12.5px] leading-relaxed">
              {lines.length === 0 ? (
                <div className="text-slate-500">PS C:\WorkPilot&gt; <span className="animate-pulse">_</span><div className="mt-2 text-slate-600">Waiting to connect a source tenant…</div></div>
              ) : lines.map((l, i) => (
                <div key={i} className={`whitespace-pre-wrap break-words ${lineColor[l.level]}`}>
                  {l.level === 'cmd' ? <><span className="text-slate-500">PS C:\WorkPilot&gt; </span>{l.text}</> : l.text}
                </div>
              ))}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
