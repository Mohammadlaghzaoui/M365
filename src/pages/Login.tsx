import { useState } from 'react';
import { Loader2, Lock, LogIn, Mail, ShieldCheck, Sparkles, Terminal, ArrowLeftRight } from 'lucide-react';
import { login, microsoftSession, googleSession, Session } from '../services/auth';
import { signIn as msSignIn } from '../services/sso';
import { googleSignIn } from '../services/googleAuth';
import { getBranding, getSSOSettings, getGoogleSSO } from '../store/settings';
import { LogoFull, LogoMark } from '../components/Logo';

function MicrosoftIcon() {
  return (
    <span className="grid h-4 w-4 grid-cols-2 grid-rows-2 gap-px">
      <span className="bg-[#f25022]" /><span className="bg-[#7fba00]" /><span className="bg-[#00a4ef]" /><span className="bg-[#ffb900]" />
    </span>
  );
}

function GoogleIcon() {
  return (
    <svg width="16" height="16" viewBox="0 0 48 48">
      <path fill="#FFC107" d="M43.6 20.1H42V20H24v8h11.3c-1.6 4.7-6.1 8-11.3 8-6.6 0-12-5.4-12-12s5.4-12 12-12c3.1 0 5.9 1.2 8 3l5.7-5.7C34 6.1 29.3 4 24 4 13 4 4 13 4 24s9 20 20 20 20-9 20-20c0-1.3-.1-2.6-.4-3.9z"/>
      <path fill="#FF3D00" d="M6.3 14.7l6.6 4.8C14.7 15.1 19 12 24 12c3.1 0 5.9 1.2 8 3l5.7-5.7C34 6.1 29.3 4 24 4 16.3 4 9.7 8.3 6.3 14.7z"/>
      <path fill="#4CAF50" d="M24 44c5.2 0 9.9-2 13.4-5.2l-6.2-5.2C29.2 35.1 26.7 36 24 36c-5.2 0-9.6-3.3-11.3-8l-6.5 5C9.5 39.6 16.2 44 24 44z"/>
      <path fill="#1976D2" d="M43.6 20.1H42V20H24v8h11.3c-.8 2.2-2.2 4.2-4.1 5.6l6.2 5.2C36.9 40.2 44 35 44 24c0-1.3-.1-2.6-.4-3.9z"/>
    </svg>
  );
}

export default function Login({ onLogin }: { onLogin: (s: Session) => void }) {
  const branding = getBranding();
  const msSso = getSSOSettings();
  const gSso = getGoogleSSO();
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState('');

  const msConfigured = msSso.enabled && !!msSso.clientId && !!msSso.tenantId;
  const gConfigured = gSso.enabled && !!gSso.clientId;

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
    setBusy(true); setError('');
    try {
      const account = await msSignIn();
      onLogin(microsoftSession(account.username));
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    } finally { setBusy(false); }
  };

  const google = async () => {
    setBusy(true); setError('');
    try {
      const me = await googleSignIn();
      onLogin(googleSession(me.email));
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    } finally { setBusy(false); }
  };

  return (
    <div className="flex min-h-screen bg-slate-950">
      {/* Left brand panel */}
      <div className="relative hidden flex-1 flex-col justify-between overflow-hidden bg-gradient-to-br from-blue-700 via-blue-900 to-slate-950 p-12 lg:flex">
        <div className="absolute -right-32 -top-32 h-96 w-96 rounded-full bg-blue-500/20 blur-3xl" />
        <div className="absolute -bottom-40 -left-20 h-96 w-96 rounded-full bg-violet-500/20 blur-3xl" />
        <div className="relative"><LogoFull height={40} light /></div>
        <div className="relative max-w-md">
          <h1 className="text-4xl font-bold leading-tight text-white">Run your Microsoft 365 practice from one cockpit.</h1>
          <p className="mt-4 text-blue-200">Guided troubleshooting, migrations with test-mode validation, user provisioning, security hardening and AI assistance.</p>
          <div className="mt-8 space-y-3">
            {[
              { icon: Terminal, text: '80+ guided workflows with PowerShell, roles & escalation paths' },
              { icon: ArrowLeftRight, text: 'Migration console with live passes, error handling & reports' },
              { icon: ShieldCheck, text: 'Security hardening baseline & user provisioning via Microsoft Graph' },
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
          <div className="mb-8 lg:hidden"><LogoFull height={34} light /></div>
          <h2 className="text-2xl font-bold text-white">Sign in</h2>
          <p className="mt-1 text-sm text-slate-400">Welcome back — sign in to your workspace.</p>

          {/* SSO buttons — always visible */}
          <div className="mt-7 space-y-2.5">
            <button onClick={microsoft} disabled={busy || !msConfigured}
              title={msConfigured ? '' : 'Configure in Settings → Sign-in (SSO) → Microsoft'}
              className="flex w-full items-center justify-center gap-2.5 rounded-lg border border-slate-700 bg-slate-900 py-2.5 text-sm font-semibold text-white hover:bg-slate-800 disabled:cursor-not-allowed disabled:opacity-40">
              <MicrosoftIcon /> Sign in with Microsoft 365
            </button>
            <button onClick={google} disabled={busy || !gConfigured}
              title={gConfigured ? '' : 'Configure in Settings → Sign-in (SSO) → Google'}
              className="flex w-full items-center justify-center gap-2.5 rounded-lg border border-slate-700 bg-slate-900 py-2.5 text-sm font-semibold text-white hover:bg-slate-800 disabled:cursor-not-allowed disabled:opacity-40">
              <GoogleIcon /> Sign in with Google
            </button>
            {(!msConfigured || !gConfigured) && (
              <p className="text-center text-[11px] text-slate-500">
                {!msConfigured && !gConfigured ? 'SSO buttons activate once configured in Settings → Sign-in (SSO).' : !msConfigured ? 'Microsoft SSO not configured yet (Settings).' : 'Google SSO not configured yet (Settings).'}
              </p>
            )}
          </div>

          <div className="my-6 flex items-center gap-3 text-xs text-slate-500">
            <div className="h-px flex-1 bg-slate-800" />or use a local account<div className="h-px flex-1 bg-slate-800" />
          </div>

          <form onSubmit={submit} className="space-y-4">
            <label className="block">
              <span className="mb-1.5 block text-sm font-medium text-slate-300">Email</span>
              <div className="relative">
                <Mail size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-500" />
                <input type="email" required value={email} onChange={(e) => setEmail(e.target.value)}
                  placeholder="admin@sorrento.cloud" autoComplete="username"
                  className="w-full rounded-lg border border-slate-700 bg-slate-900 py-2.5 pl-9 pr-3 text-sm text-white placeholder-slate-600 focus:border-blue-500 focus:outline-none focus:ring-1 focus:ring-blue-500" />
              </div>
            </label>
            <label className="block">
              <span className="mb-1.5 block text-sm font-medium text-slate-300">Password</span>
              <div className="relative">
                <Lock size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-500" />
                <input type="password" required value={password} onChange={(e) => setPassword(e.target.value)}
                  placeholder="••••••••" autoComplete="current-password"
                  className="w-full rounded-lg border border-slate-700 bg-slate-900 py-2.5 pl-9 pr-3 text-sm text-white placeholder-slate-600 focus:border-blue-500 focus:outline-none focus:ring-1 focus:ring-blue-500" />
              </div>
            </label>
            {error && <p className="text-sm text-red-400">{error}</p>}
            <button type="submit" disabled={busy}
              className="flex w-full items-center justify-center gap-2 rounded-lg bg-blue-600 py-2.5 text-sm font-semibold text-white transition-colors hover:bg-blue-500 disabled:opacity-50">
              {busy ? <Loader2 size={16} className="animate-spin" /> : <LogIn size={16} />} Sign in
            </button>
          </form>

          <div className="mt-8 flex items-center justify-center gap-2 text-xs text-slate-600">
            <LogoMark size={16} /> {branding.portalName} — enterprise service portal
          </div>
        </div>
      </div>
    </div>
  );
}
