import { Link } from 'react-router-dom';
import * as Icons from 'lucide-react';
import { Card, ProgressBar, Badge } from '../components/ui';
import { quickActions, dailyChecklist, portalShortcuts } from '../data/dashboard';
import { useLocalStorage, load } from '../store/useLocalStorage';
import { Ticket, MigrationProject, SecurityControlState, ProvisioningRequest, ConsoleProjectState } from '../types';
import { securityControls } from '../data/security';
import { psTasks } from '../data/psTasks';
import { ExternalLink, Star, Activity } from 'lucide-react';
import { getBranding } from '../store/settings';
import { getSession } from '../services/auth';
import { DonutChart, HBarChart, VBarChart } from '../components/charts';
import { useEffect, useState } from 'react';
import { getTenantInsights, TenantInsights } from '../services/graphInsights';
import { currentAccount } from '../services/sso';
import { Users2, UserX, UserCheck, RefreshCw } from 'lucide-react';

const colorMap: Record<string, string> = {
  blue: 'bg-blue-100 dark:bg-blue-900/40 text-blue-600 dark:text-blue-300',
  green: 'bg-emerald-100 dark:bg-emerald-900/40 text-emerald-600 dark:text-emerald-300',
  orange: 'bg-amber-100 dark:bg-amber-900/40 text-amber-600 dark:text-amber-300',
  purple: 'bg-violet-100 dark:bg-violet-900/40 text-violet-600 dark:text-violet-300',
};

function todayKey() {
  return `daily-checklist:${new Date().toISOString().slice(0, 10)}`;
}

export default function Dashboard() {
  const tickets = load<Ticket[]>('tickets', []);
  const projects = load<MigrationProject[]>('migration-projects', []);
  const secState = load<Record<string, SecurityControlState>>('security-state', {});
  const recent = load<string[]>('recent-workflows', []);
  const favPs = load<string[]>('favorite-ps', []);
  const [checked, setChecked] = useLocalStorage<string[]>(todayKey(), []);
  const [insights, setInsights] = useState<TenantInsights | null>(() => load<TenantInsights | null>('tenant-insights', null));
  const [insightsState, setInsightsState] = useState<'idle' | 'loading' | 'error' | 'nosso'>('idle');

  const refreshInsights = async () => {
    setInsightsState('loading');
    try {
      if (!(await currentAccount())) { setInsightsState('nosso'); return; }
      const data = await getTenantInsights();
      setInsights(data);
      localStorage.setItem('workpilot:tenant-insights', JSON.stringify(data));
      setInsightsState('idle');
    } catch {
      setInsightsState('error');
    }
  };

  useEffect(() => {
    // Auto-refresh live tenant data when signed in with Microsoft SSO.
    currentAccount().then((a) => { if (a) refreshInsights(); }).catch(() => {});
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const provLog = load<ProvisioningRequest[]>('provisioning-requests', []);
  const mwStates = load<Record<string, ConsoleProjectState>>('mw-console', {});
  const mwItems = Object.values(mwStates).flatMap((s) => s.items);
  const branding = getBranding();
  const session = getSession();

  const openTickets = tickets.filter((t) => t.status !== 'closed');
  const urgent = openTickets.filter((t) => t.urgency === 'high' || t.urgency === 'critical');
  const activeProjects = projects.filter((p) => p.status !== 'completed');
  const compliant = securityControls.filter((c) => secState[c.id]?.status === 'compliant').length;
  const secPct = Math.round((compliant / securityControls.length) * 100);
  const dailyPct = Math.round((checked.length / dailyChecklist.length) * 100);
  const favTasks = psTasks.filter((t) => favPs.includes(t.id));
  const mwCompleted = mwItems.filter((i) => i.status === 'Completed').length;
  const mwFailed = mwItems.filter((i) => i.status === 'Failed' || i.status === 'VerifyFailed').length;
  const mwRunning = mwItems.filter((i) => ['Verifying', 'Assessing', 'PreStaging', 'Migrating', 'DeltaSync'].includes(i.status)).length;
  const mwPending = mwItems.length - mwCompleted - mwFailed - mwRunning;

  const secCounts = {
    compliant,
    partial: securityControls.filter((c) => secState[c.id]?.status === 'partial').length,
    nonCompliant: securityControls.filter((c) => secState[c.id]?.status === 'non-compliant').length,
  };
  const secOpen = securityControls.length - secCounts.compliant - secCounts.partial - secCounts.nonCompliant;

  const ticketsByService = Object.entries(
    openTickets.reduce<Record<string, number>>((acc, t) => ({ ...acc, [t.service]: (acc[t.service] ?? 0) + 1 }), {}),
  ).sort((a, b) => b[1] - a[1]).slice(0, 6).map(([label, value]) => ({ label, value }));

  const urgencyColors: Record<string, string> = { critical: '#dc2626', high: '#f59e0b', medium: '#2563eb', low: '#64748b' };
  const ticketsByUrgency = (['critical', 'high', 'medium', 'low'] as const)
    .map((u) => ({ label: u, value: openTickets.filter((t) => t.urgency === u).length, color: urgencyColors[u] }))
    .filter((d) => d.value > 0);

  const activity: { time: string; text: string; color: string }[] = [
    ...provLog.slice(0, 6).map((p) => ({ time: p.createdAt, text: `Provisioning ${p.status}: ${p.displayName || p.mail} (${p.accountType})`, color: p.status === 'Failed' || p.status === 'ValidationError' ? 'bg-red-500' : p.status === 'Created' || p.status === 'Invited' ? 'bg-emerald-500' : 'bg-amber-500' })),
    ...tickets.slice(0, 6).map((t) => ({ time: t.createdAt, text: `Ticket ${t.status}: [${t.service}] ${t.category} — ${t.customer}`, color: t.urgency === 'high' || t.urgency === 'critical' ? 'bg-red-500' : 'bg-blue-500' })),
  ].sort((a, b) => b.time.localeCompare(a.time)).slice(0, 8);

  const toggle = (id: string) =>
    setChecked((prev) => (prev.includes(id) ? prev.filter((c) => c !== id) : [...prev, id]));

  return (
    <div className="space-y-6">
      {/* Hero */}
      <div className="relative overflow-hidden rounded-2xl bg-gradient-to-r from-blue-700 via-blue-800 to-violet-900 p-6 text-white shadow-lg">
        <div className="absolute -right-16 -top-16 h-64 w-64 rounded-full bg-white/10 blur-2xl" />
        <div className="relative flex flex-wrap items-end justify-between gap-4">
          <div>
            <p className="text-xs font-semibold uppercase tracking-widest text-blue-200">{branding.companyName} · {new Date().toLocaleDateString(undefined, { weekday: 'long', day: 'numeric', month: 'long' })}</p>
            <h1 className="mt-1 text-2xl font-bold">Welcome back{session ? `, ${session.email.split('@')[0]}` : ''} 👋</h1>
            <p className="mt-1 text-sm text-blue-100">Your Microsoft 365 service cockpit — {openTickets.length} open ticket{openTickets.length === 1 ? '' : 's'}, {activeProjects.length} active migration project{activeProjects.length === 1 ? '' : 's'}, security baseline at {secPct}%.</p>
          </div>
          <div className="flex gap-6 text-center">
            <div><div className="text-3xl font-bold">{mwCompleted}</div><div className="text-[11px] uppercase tracking-wide text-blue-200">mailboxes migrated</div></div>
            <div><div className={`text-3xl font-bold ${mwFailed ? 'text-amber-300' : ''}`}>{mwFailed}</div><div className="text-[11px] uppercase tracking-wide text-blue-200">migration failures</div></div>
            <div><div className="text-3xl font-bold">{provLog.length}</div><div className="text-[11px] uppercase tracking-wide text-blue-200">provisioning requests</div></div>
          </div>
        </div>
      </div>

      {/* Live tenant insights (real Microsoft Graph data) */}
      <Card className="p-5">
        <div className="mb-3 flex items-center justify-between">
          <h2 className="text-sm font-semibold uppercase tracking-wide text-slate-500 dark:text-slate-400">Live tenant insights — Microsoft Graph</h2>
          <button onClick={refreshInsights} className="flex items-center gap-1.5 text-xs font-semibold text-blue-500 hover:underline" disabled={insightsState === 'loading'}>
            <RefreshCw size={13} className={insightsState === 'loading' ? 'animate-spin' : ''} /> Refresh
          </button>
        </div>
        {insights ? (
          <div className="grid grid-cols-3 gap-4">
            <div className="flex items-center gap-3">
              <span className="rounded-lg bg-blue-100 dark:bg-blue-900/40 p-2.5 text-blue-600 dark:text-blue-300"><Users2 size={18} /></span>
              <div><div className="text-2xl font-bold text-slate-800 dark:text-slate-100">{insights.totalUsers.toLocaleString()}</div><div className="text-xs text-slate-400">users in tenant</div></div>
            </div>
            <div className="flex items-center gap-3">
              <span className="rounded-lg bg-violet-100 dark:bg-violet-900/40 p-2.5 text-violet-600 dark:text-violet-300"><UserCheck size={18} /></span>
              <div><div className="text-2xl font-bold text-slate-800 dark:text-slate-100">{insights.guests.toLocaleString()}</div><div className="text-xs text-slate-400">guest accounts</div></div>
            </div>
            <div className="flex items-center gap-3">
              <span className="rounded-lg bg-amber-100 dark:bg-amber-900/40 p-2.5 text-amber-600 dark:text-amber-300"><UserX size={18} /></span>
              <div><div className="text-2xl font-bold text-slate-800 dark:text-slate-100">{insights.disabled.toLocaleString()}</div><div className="text-xs text-slate-400">disabled accounts</div></div>
            </div>
          </div>
        ) : (
          <p className="text-sm text-slate-400">
            {insightsState === 'loading' ? 'Loading live data from your tenant…'
              : insightsState === 'error' ? 'Could not load tenant data — check Graph permissions (User.Read.All) on the SSO app registration.'
              : 'Sign in with Microsoft 365 (Settings → Sign-in) to show real user/guest/disabled counts straight from your tenant.'}
          </p>
        )}
        {insights && <p className="mt-3 text-[11px] text-slate-400">Real data from Microsoft Graph · last refresh {new Date(insights.fetchedAt).toLocaleString()}</p>}
      </Card>

      {/* Analytics charts */}
      <div className="grid gap-4 lg:grid-cols-3">
        <Card className="p-5">
          <h2 className="mb-3 text-sm font-semibold uppercase tracking-wide text-slate-500 dark:text-slate-400">Security baseline</h2>
          <DonutChart
            centerLabel="of 27 controls compliant"
            centerValue={`${secPct}%`}
            segments={[
              { label: 'Compliant', value: secCounts.compliant, color: '#10b981' },
              { label: 'Partial', value: secCounts.partial, color: '#f59e0b' },
              { label: 'Non-compliant', value: secCounts.nonCompliant, color: '#dc2626' },
              { label: 'Not reviewed', value: secOpen, color: '#94a3b8' },
            ]}
          />
        </Card>
        <Card className="p-5">
          <h2 className="mb-3 text-sm font-semibold uppercase tracking-wide text-slate-500 dark:text-slate-400">Open tickets by service</h2>
          {ticketsByService.length ? <HBarChart data={ticketsByService} /> : <p className="text-xs text-slate-400">No open tickets — nice and quiet. 🎉</p>}
          {ticketsByUrgency.length > 0 && (
            <div className="mt-4">
              <h3 className="mb-2 text-xs font-semibold uppercase text-slate-400">By urgency</h3>
              <HBarChart data={ticketsByUrgency} />
            </div>
          )}
        </Card>
        <Card className="p-5">
          <h2 className="mb-3 text-sm font-semibold uppercase tracking-wide text-slate-500 dark:text-slate-400">Migration pipeline</h2>
          {mwItems.length ? (
            <VBarChart data={[
              { label: 'Pending', value: mwPending },
              { label: 'Running', value: mwRunning },
              { label: 'Completed', value: mwCompleted },
              { label: 'Failed', value: mwFailed },
            ]} />
          ) : (
            <p className="text-xs text-slate-400">No line items yet — open the <Link to="/migration-console" className="text-blue-500 hover:underline">Migration Console</Link> and import users.</p>
          )}
          <div className="mt-4">
            <h3 className="mb-2 text-xs font-semibold uppercase text-slate-400">Daily checks today</h3>
            <ProgressBar value={dailyPct} color="bg-blue-600" />
            <p className="mt-1 text-xs text-slate-400">{checked.length} of {dailyChecklist.length} completed</p>
          </div>
        </Card>
      </div>

      {/* Stats */}
      <div className="grid grid-cols-2 gap-4 lg:grid-cols-5">
        <Card className="p-4">
          <div className="text-xs font-semibold uppercase text-slate-400">Open tickets</div>
          <div className="mt-1 text-3xl font-bold text-blue-600 dark:text-blue-400">{openTickets.length}</div>
          <Link to="/tickets" className="text-xs text-blue-500 hover:underline">Ticket Assistant →</Link>
        </Card>
        <Card className="p-4">
          <div className="text-xs font-semibold uppercase text-slate-400">Urgent tickets</div>
          <div className={`mt-1 text-3xl font-bold ${urgent.length ? 'text-red-500' : 'text-emerald-500'}`}>{urgent.length}</div>
          <span className="text-xs text-slate-400">{urgent.length ? 'needs attention' : 'all clear'}</span>
        </Card>
        <Card className="p-4">
          <div className="text-xs font-semibold uppercase text-slate-400">Migration projects</div>
          <div className="mt-1 text-3xl font-bold text-violet-600 dark:text-violet-400">{activeProjects.length}</div>
          <Link to="/migration" className="text-xs text-blue-500 hover:underline">Projects →</Link>
        </Card>
        <Card className="p-4">
          <div className="text-xs font-semibold uppercase text-slate-400">Security hardening</div>
          <div className="mt-1 text-3xl font-bold text-emerald-600 dark:text-emerald-400">{secPct}%</div>
          <ProgressBar value={secPct} color="bg-emerald-500" />
        </Card>
        <Card className="p-4">
          <div className="text-xs font-semibold uppercase text-slate-400">Daily checks</div>
          <div className="mt-1 text-3xl font-bold text-blue-600 dark:text-blue-400">{checked.length}/{dailyChecklist.length}</div>
          <ProgressBar value={dailyPct} />
        </Card>
      </div>

      {/* Quick actions */}
      <div>
        <h2 className="mb-3 text-sm font-semibold uppercase tracking-wide text-slate-500 dark:text-slate-400">Quick actions</h2>
        <div className="grid grid-cols-2 gap-3 sm:grid-cols-3 lg:grid-cols-5">
          {quickActions.map((a) => {
            const Icon = (Icons as unknown as Record<string, Icons.LucideIcon>)[a.icon] ?? Icons.Zap;
            return (
              <Link key={a.label} to={a.to} className="group rounded-xl border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800 p-4 shadow-sm transition-all hover:-translate-y-0.5 hover:shadow-md">
                <div className={`mb-2 inline-flex rounded-lg p-2 ${colorMap[a.color]}`}><Icon size={18} /></div>
                <div className="text-sm font-semibold text-slate-700 dark:text-slate-200 group-hover:text-blue-600 dark:group-hover:text-blue-400">{a.label}</div>
              </Link>
            );
          })}
        </div>
      </div>

      <div className="grid gap-6 lg:grid-cols-3">
        {/* Daily checklist */}
        <Card className="p-5">
          <h2 className="mb-3 text-sm font-semibold uppercase tracking-wide text-slate-500 dark:text-slate-400">Daily checklist</h2>
          <div className="space-y-1.5">
            {dailyChecklist.map((c) => (
              <label key={c.id} className="flex cursor-pointer items-center gap-2.5 rounded-lg px-2 py-1.5 hover:bg-slate-50 dark:hover:bg-slate-700/50">
                <input type="checkbox" checked={checked.includes(c.id)} onChange={() => toggle(c.id)} className="h-4 w-4 rounded border-slate-300 text-blue-600 focus:ring-blue-500" />
                <span className={`text-sm ${checked.includes(c.id) ? 'text-slate-400 line-through' : 'text-slate-700 dark:text-slate-200'}`}>{c.label}</span>
              </label>
            ))}
          </div>
          <Link to="/monitoring" className="mt-3 inline-block text-xs font-semibold text-blue-500 hover:underline">Open Monitoring Assistant →</Link>
        </Card>

        {/* Recent + favorites */}
        <Card className="p-5">
          <h2 className="mb-3 text-sm font-semibold uppercase tracking-wide text-slate-500 dark:text-slate-400">Recently used workflows</h2>
          {recent.length === 0 ? (
            <p className="text-sm text-slate-400">Workflows you open will appear here.</p>
          ) : (
            <ul className="space-y-2">
              {recent.map((r) => (
                <li key={r} className="text-sm text-slate-700 dark:text-slate-200">• {r}</li>
              ))}
            </ul>
          )}
          <h2 className="mb-2 mt-5 text-sm font-semibold uppercase tracking-wide text-slate-500 dark:text-slate-400">Favorite PowerShell</h2>
          {favTasks.length === 0 ? (
            <p className="text-sm text-slate-400">Star commands in the <Link className="text-blue-500 hover:underline" to="/powershell">PowerShell Generator</Link>.</p>
          ) : (
            <ul className="space-y-1.5">
              {favTasks.map((t) => (
                <li key={t.id}>
                  <Link to="/powershell" className="flex items-center gap-1.5 text-sm text-slate-700 dark:text-slate-200 hover:text-blue-500">
                    <Star size={13} className="text-amber-400 fill-amber-400" />{t.name}
                  </Link>
                </li>
              ))}
            </ul>
          )}
        </Card>

        {/* Portal shortcuts */}
        <Card className="p-5">
          <h2 className="mb-3 text-sm font-semibold uppercase tracking-wide text-slate-500 dark:text-slate-400">Admin portal shortcuts</h2>
          <div className="space-y-1.5">
            {portalShortcuts.map((p) => (
              <a key={p.url} href={p.url} target="_blank" rel="noreferrer" className="flex items-center justify-between rounded-lg px-2 py-1.5 text-sm text-slate-700 dark:text-slate-200 hover:bg-slate-50 dark:hover:bg-slate-700/50 hover:text-blue-500">
                {p.label}
                <ExternalLink size={13} className="text-slate-400" />
              </a>
            ))}
          </div>
        </Card>
      </div>

      {/* Activity feed */}
      {activity.length > 0 && (
        <Card className="p-5">
          <div className="mb-3 flex items-center gap-2 text-sm font-semibold uppercase tracking-wide text-slate-500 dark:text-slate-400"><Activity size={14} /> Recent activity</div>
          <div className="space-y-2.5">
            {activity.map((a, i) => (
              <div key={i} className="flex items-start gap-3 text-sm">
                <span className={`mt-1.5 h-2 w-2 shrink-0 rounded-full ${a.color}`} />
                <span className="flex-1 text-slate-700 dark:text-slate-200">{a.text}</span>
                <span className="shrink-0 text-xs text-slate-400">{new Date(a.time).toLocaleString(undefined, { day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' })}</span>
              </div>
            ))}
          </div>
        </Card>
      )}

      {/* Open tickets table */}
      {openTickets.length > 0 && (
        <Card className="p-5">
          <h2 className="mb-3 text-sm font-semibold uppercase tracking-wide text-slate-500 dark:text-slate-400">Open tickets</h2>
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead>
                <tr className="border-b border-slate-200 dark:border-slate-700 text-left text-xs uppercase text-slate-400">
                  <th className="py-2 pr-4">Customer</th><th className="py-2 pr-4">User</th><th className="py-2 pr-4">Service</th><th className="py-2 pr-4">Category</th><th className="py-2 pr-4">Urgency</th><th className="py-2">Status</th>
                </tr>
              </thead>
              <tbody>
                {openTickets.slice(0, 10).map((t) => (
                  <tr key={t.id} className="border-b border-slate-100 dark:border-slate-700/50">
                    <td className="py-2 pr-4 font-medium text-slate-700 dark:text-slate-200">{t.customer}</td>
                    <td className="py-2 pr-4 text-slate-500 dark:text-slate-300">{t.userName}</td>
                    <td className="py-2 pr-4 text-slate-500 dark:text-slate-300">{t.service}</td>
                    <td className="py-2 pr-4 text-slate-500 dark:text-slate-300">{t.category}</td>
                    <td className="py-2 pr-4"><Badge color={t.urgency === 'critical' || t.urgency === 'high' ? 'red' : t.urgency === 'medium' ? 'orange' : 'gray'}>{t.urgency}</Badge></td>
                    <td className="py-2"><Badge color={t.status === 'escalated' ? 'red' : t.status === 'open' ? 'blue' : 'gray'}>{t.status}</Badge></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </Card>
      )}
    </div>
  );
}
