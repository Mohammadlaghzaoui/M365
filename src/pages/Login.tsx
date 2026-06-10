import { useState } from 'react';
import { Loader2, Lock, LogIn, Mail, ShieldCheck, Sparkles, Terminal, ArrowLeftRight } from 'lucide-react';
import { login, microsoftSession, Session } from '../services/auth';
import { signIn as msSignIn } from '../services/sso';
import { getBranding, getSSOSettings } from '../store/settings';

export default function Login({ onLogin }: { onLogin: (s: Session) => void }) {
  const branding = getBranding();
  const sso = getSSOSettings();
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState('');

  const submit = async (e: React.FormEvent) => {
    e.preventDefault();
    setBusy(true);
    setError('');
    try {
      onLogin(await login(email, password));
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setBusy(false);
    }
  };

  const microsoft = async () => {
    setBusy(true);
    setError('');
    try {
      const account = await msSignIn();
      onLogin(microsoftSession(account.username));
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className="flex min-h-screen bg-slate-950">
      {/* Left brand panel */}
      <div className="relative hidden flex-1 flex-col justify-between overflow-hidden bg-gradient-to-br from-blue-700 via-blue-900 to-slate-950 p-12 lg:flex">
        <div className="absolute -right-32 -top-32 h-96 w-96 rounded-full bg-blue-500/20 blur-3xl" />
        <div className="absolute -bottom-40 -left-20 h-96 w-96 rounded-full bg-violet-500/20 blur-3xl" />
        <div className="relative flex items-center gap-3">
          <div className="grid h-10 w-10 grid-cols-2 grid-rows-2 gap-0.5 rounded-lg bg-white/10 p-1">
            <div className="rounded-sm bg-[#f25022]" /><div className="rounded-sm bg-[#7fba00]" />
            <div className="rounded-sm bg-[#00a4ef]" /><div className="rounded-sm bg-[#ffb900]" />
          </div>
          <div>
            <div className="text-lg font-bold text-white">{branding.portalName}</div>
            <div className="text-xs text-blue-200">{branding.companyName} — Service Provider Portal</div>
          </div>
        </div>
        <div className="relative max-w-md">
          <h1 className="text-4xl font-bold leading-tight text-white">Run your Microsoft 365 practice from one cockpit.</h1>
          <p className="mt-4 text-blue-200">Guided troubleshooting, migrations with test-mode validation, security hardening and AI assistance — for Entra ID, Exchange Online, SharePoint, Teams, BitTitan and Syskit.</p>
          <div className="mt-8 space-y-3">
            {[
              { icon: Terminal, text: '80+ guided workflows with PowerShell, roles & escalation paths' },
              { icon: ArrowLeftRight, text: 'Cross-tenant & hybrid migrations with animated test runs' },
              { icon: ShieldCheck, text: '27-control security hardening baseline with 30/90-day plans' },
              { icon: Sparkles, text: 'AI assistant — OpenRouter, OpenAI or Claude, your choice' },
            ].map((f) => (
              <div key={f.text} className="flex items-center gap-3 text-sm text-blue-100">
                <span className="rounded-lg bg-white/10 p-2"><f.icon size={16} /></span>{f.text}
              </div>
            ))}
          </div>
        </div>
        <div className="relative text-xs text-blue-300/60">© {new Date().getFullYear()} {branding.companyName} · sorrento.cloud</div>
      </div>

      {/* Right login card */}
      <div className="flex w-full items-center justify-center p-6 lg:w-[480px] lg:shrink-0">
        <div className="w-full max-w-sm">
          <div className="mb-8 lg:hidden flex items-center gap-3">
            <div className="grid h-9 w-9 grid-cols-2 grid-rows-2 gap-0.5 rounded-lg bg-white/10 p-1">
              <div className="rounded-sm bg-[#f25022]" /><div className="rounded-sm bg-[#7fba00]" />
              <div className="rounded-sm bg-[#00a4ef]" /><div className="rounded-sm bg-[#ffb900]" />
            </div>
            <span className="text-lg font-bold text-white">{branding.portalName}</span>
          </div>
          <h2 className="text-2xl font-bold text-white">Sign in</h2>
          <p className="mt-1 text-sm text-slate-400">Welcome back — sign in to your workspace.</p>

          <form onSubmit={submit} className="mt-8 space-y-4">
            <label className="block">
              <span className="mb-1.5 block text-sm font-medium text-slate-300">Email</span>
              <div className="relative">
                <Mail size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-500" />
                <input
                  type="email" required value={email} onChange={(e) => setEmail(e.target.value)}
                  placeholder="admin@sorrento.cloud" autoComplete="username"
                  className="w-full rounded-lg border border-slate-700 bg-slate-900 py-2.5 pl-9 pr-3 text-sm text-white placeholder-slate-600 focus:border-blue-500 focus:outline-none focus:ring-1 focus:ring-blue-500"
                />
              </div>
            </label>
            <label className="block">
              <span className="mb-1.5 block text-sm font-medium text-slate-300">Password</span>
              <div className="relative">
                <Lock size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-500" />
                <input
                  type="password" required value={password} onChange={(e) => setPassword(e.target.value)}
                  placeholder="••••••••" autoComplete="current-password"
                  className="w-full rounded-lg border border-slate-700 bg-slate-900 py-2.5 pl-9 pr-3 text-sm text-white placeholder-slate-600 focus:border-blue-500 focus:outline-none focus:ring-1 focus:ring-blue-500"
                />
              </div>
            </label>
            {error && <p className="text-sm text-red-400">{error}</p>}
            <button type="submit" disabled={busy}
              className="flex w-full items-center justify-center gap-2 rounded-lg bg-blue-600 py-2.5 text-sm font-semibold text-white transition-colors hover:bg-blue-500 disabled:opacity-50">
              {busy ? <Loader2 size={16} className="animate-spin" /> : <LogIn size={16} />} Sign in
            </button>
          </form>

          {sso.enabled && sso.clientId && (
            <>
              <div className="my-6 flex items-center gap-3 text-xs text-slate-500">
                <div className="h-px flex-1 bg-slate-800" />or<div className="h-px flex-1 bg-slate-800" />
              </div>
              <button onClick={microsoft} disabled={busy}
                className="flex w-full items-center justify-center gap-2.5 rounded-lg border border-slate-700 bg-slate-900 py-2.5 text-sm font-semibold text-white hover:bg-slate-800 disabled:opacity-50">
                <span className="grid h-4 w-4 grid-cols-2 grid-rows-2 gap-px">
                  <span className="bg-[#f25022]" /><span className="bg-[#7fba00]" /><span className="bg-[#00a4ef]" /><span className="bg-[#ffb900]" />
                </span>
                Sign in with Microsoft
              </button>
            </>
          )}

          <p className="mt-8 text-center text-xs text-slate-600">
            Local sign-in for the first version — switch to Microsoft 365 SSO in Settings for enterprise auth.
          </p>
        </div>
      </div>
    </div>
  );
}
