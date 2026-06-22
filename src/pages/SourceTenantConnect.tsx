import { useEffect, useRef, useState } from 'react';
import { useNavigate, Link } from 'react-router-dom';
import { ArrowLeft, Building2, ShieldCheck, Loader2, Copy, Check, ExternalLink, Terminal } from 'lucide-react';
import { runDiscovery, useDiscoveryToken, LogLevel } from '../services/graphDiscovery';
import { requestDeviceCode, pollForToken, getOnecomToken, clearOnecomToken, DeviceCode } from '../services/onecomDeviceAuth';
import { saveTenantResult } from '../services/tenantStore';
import { save } from '../store/useLocalStorage';

type Line = { text: string; level: LogLevel };

export default function SourceTenantConnect() {
  const nav = useNavigate();
  const [phase, setPhase] = useState<'idle' | 'awaiting' | 'collecting' | 'error'>('idle');
  const [code, setCode] = useState<DeviceCode | null>(null);
  const [lines, setLines] = useState<Line[]>([]);
  const [error, setError] = useState('');
  const [copied, setCopied] = useState(false);
  const logRef = useRef<HTMLDivElement>(null);
  useEffect(() => { logRef.current?.scrollTo({ top: logRef.current.scrollHeight }); }, [lines]);

  const addLine = (text: string, level: LogLevel = 'info') =>
    setLines((p) => [...p, { text: level === 'cmd' ? text : `[${new Date().toLocaleTimeString('en-GB')}] ${text}`, level }]);

  const start = async () => {
    setError(''); setLines([]); setCode(null);
    try {
      addLine('Requesting a device sign-in code from Microsoft …', 'info');
      const dc = await requestDeviceCode();
      setCode(dc); setPhase('awaiting');
      const { tenantId } = await pollForToken(dc);
      setPhase('collecting');
      addLine(`Signed in (tenant ${tenantId}). Starting read-only discovery …`, 'ok');
      useDiscoveryToken(getOnecomToken);
      const r = await runDiscovery(addLine);
      addLine('Read-only assessment complete. No tenant changes were made.', 'ok');
      save('discovery-result', r);
      saveTenantResult(r);
      addLine('Opening the Source Tenant dashboard …', 'ok');
      setTimeout(() => nav(`/tenants/${r.org.tenantId}`), 600);
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
      addLine(`ERROR: ${e instanceof Error ? e.message : e}`, 'err');
      setPhase('error');
    } finally {
      clearOnecomToken();
    }
  };

  const copy = () => { if (code) { navigator.clipboard?.writeText(code.user_code); setCopied(true); setTimeout(() => setCopied(false), 1500); } };
  const lineColor: Record<LogLevel, string> = { cmd: 'text-sky-300', info: 'text-slate-300', ok: 'text-emerald-400', warn: 'text-amber-300', err: 'text-rose-400' };

  return (
    <div className="mx-auto max-w-[1100px]">
      <Link to="/tenants" className="mb-3 inline-flex items-center gap-1 text-sm text-slate-500 hover:text-blue-700"><ArrowLeft size={14} /> Source Tenants</Link>

      {/* Header band */}
      <div className="mb-5 flex flex-wrap items-center gap-4 rounded-lg border border-slate-300 bg-white p-5 shadow-sm dark:border-slate-700 dark:bg-slate-900">
        <div className="flex h-12 w-12 items-center justify-center rounded-md bg-slate-800 text-white"><Building2 size={22} /></div>
        <div className="mr-auto">
          <h1 className="text-xl font-bold text-slate-900 dark:text-slate-100">Add a source tenant</h1>
          <p className="text-sm text-slate-500">Sign in to the customer's tenant with a code — the read-only Microsoft Graph engine runs here and opens the tenant's dashboard.</p>
        </div>
        <span className="inline-flex items-center gap-1 rounded bg-emerald-50 px-2.5 py-1 text-xs font-medium text-emerald-700 dark:bg-emerald-900/30 dark:text-emerald-300"><ShieldCheck size={13} /> Read-only — no tenant changes</span>
      </div>

      <div className="grid grid-cols-1 gap-5 lg:grid-cols-5">
        {/* Connect card */}
        <div className="lg:col-span-2">
          <div className="rounded-lg border border-slate-300 bg-white p-5 shadow-sm dark:border-slate-700 dark:bg-slate-900">
            <h2 className="mb-1 text-base font-semibold text-slate-900 dark:text-slate-100">Connect with a code</h2>
            <p className="mb-4 text-sm text-slate-500">No app registration, no tenant ID, nothing installed. The customer's admin approves once.</p>

            {phase === 'idle' || phase === 'error' ? (
              <button onClick={start} className="inline-flex w-full items-center justify-center gap-2 rounded-md bg-blue-700 px-4 py-2.5 text-sm font-semibold text-white hover:bg-blue-800">
                <Building2 size={16} /> Connect a tenant with a code
              </button>
            ) : phase === 'awaiting' && code ? (
              <ol className="space-y-4 text-sm">
                <li>
                  <div className="mb-1 font-medium text-slate-700 dark:text-slate-200">1 · Open the sign-in page</div>
                  <a href={code.verification_uri} target="_blank" rel="noreferrer" className="inline-flex items-center gap-1.5 rounded-md border border-blue-200 bg-blue-50 px-3 py-1.5 font-medium text-blue-700 hover:bg-blue-100 dark:border-blue-800 dark:bg-blue-900/30 dark:text-blue-300">
                    {code.verification_uri || 'microsoft.com/devicelogin'} <ExternalLink size={13} />
                  </a>
                </li>
                <li>
                  <div className="mb-1 font-medium text-slate-700 dark:text-slate-200">2 · Enter this code</div>
                  <div className="flex items-center gap-2">
                    <span className="rounded-md border border-slate-300 bg-slate-50 px-4 py-2 font-mono text-2xl font-bold tracking-[0.2em] text-slate-900 dark:border-slate-600 dark:bg-slate-800 dark:text-slate-100">{code.user_code}</span>
                    <button onClick={copy} className="rounded-md border border-slate-300 p-2 text-slate-500 hover:bg-slate-100 dark:border-slate-600 dark:hover:bg-slate-800">{copied ? <Check size={16} className="text-emerald-500" /> : <Copy size={16} />}</button>
                  </div>
                </li>
                <li>
                  <div className="font-medium text-slate-700 dark:text-slate-200">3 · Sign in as the customer's admin &amp; approve</div>
                  <div className="mt-1 flex items-center gap-2 text-xs text-blue-600"><Loader2 size={13} className="animate-spin" /> Waiting for sign-in…</div>
                </li>
              </ol>
            ) : (
              <div className="flex items-center gap-2 text-sm font-medium text-blue-700 dark:text-blue-400"><Loader2 size={16} className="animate-spin" /> Signed in — collecting read-only data…</div>
            )}

            {error && <p className="mt-4 rounded-md bg-rose-50 px-3 py-2 text-sm text-rose-700 dark:bg-rose-900/20 dark:text-rose-300">{error}</p>}
          </div>
        </div>

        {/* Live console — the engine running in the tenant */}
        <div className="lg:col-span-3">
          <div className="overflow-hidden rounded-lg border border-slate-800 shadow-lg">
            <div className="flex items-center gap-2 border-b border-slate-800 bg-slate-900 px-4 py-2.5">
              <span className="h-3 w-3 rounded-full bg-rose-500" /><span className="h-3 w-3 rounded-full bg-amber-400" /><span className="h-3 w-3 rounded-full bg-emerald-500" />
              <span className="ml-2 flex items-center gap-1.5 text-xs font-medium text-slate-300"><Terminal size={13} /> Microsoft Graph — read-only discovery {(phase === 'collecting' || phase === 'awaiting') && <span className="ml-1 animate-pulse text-emerald-400">· running</span>}</span>
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
