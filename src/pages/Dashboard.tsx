import { Link } from 'react-router-dom';
import * as Icons from 'lucide-react';
import { Card, ProgressBar, Badge } from '../components/ui';
import { quickActions, dailyChecklist, portalShortcuts } from '../data/dashboard';
import { useLocalStorage, load } from '../store/useLocalStorage';
import { Ticket, MigrationProject, SecurityControlState } from '../types';
import { securityControls } from '../data/security';
import { psTasks } from '../data/psTasks';
import { ExternalLink, Star } from 'lucide-react';

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

  const openTickets = tickets.filter((t) => t.status !== 'closed');
  const urgent = openTickets.filter((t) => t.urgency === 'high' || t.urgency === 'critical');
  const activeProjects = projects.filter((p) => p.status !== 'completed');
  const compliant = securityControls.filter((c) => secState[c.id]?.status === 'compliant').length;
  const secPct = Math.round((compliant / securityControls.length) * 100);
  const dailyPct = Math.round((checked.length / dailyChecklist.length) * 100);
  const favTasks = psTasks.filter((t) => favPs.includes(t.id));

  const toggle = (id: string) =>
    setChecked((prev) => (prev.includes(id) ? prev.filter((c) => c !== id) : [...prev, id]));

  return (
    <div className="space-y-6">
      <div>
        <h1 className="text-2xl font-bold text-slate-800 dark:text-slate-100">Welcome to M365 WorkPilot</h1>
        <p className="mt-1 text-sm text-slate-500 dark:text-slate-400">Your Microsoft 365 Service Provider cockpit — pick a task and let the portal guide you step by step.</p>
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
