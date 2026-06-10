import { useEffect, useMemo, useState } from 'react';
import { NavLink, Outlet, useNavigate } from 'react-router-dom';
import {
  ArrowLeftRight, Bot, BookOpen, Boxes, Cloud, FileText, Gauge, KeyRound, LayoutDashboard,
  Mail, MailQuestion, Menu, MessageSquare, Moon, Search, Settings as SettingsIcon, Shield,
  ShieldCheck, StickyNote, Sun, Terminal, Ticket, Users, Workflow as WorkflowIcon, X, FolderKanban, Send,
} from 'lucide-react';
import { load, save } from '../store/useLocalStorage';
import { getBranding } from '../store/settings';
import { Session, logout } from '../services/auth';
import { UserCircle2, LogOut, MonitorPlay } from 'lucide-react';

export const NAV = [
  { to: '/', label: 'Dashboard', icon: LayoutDashboard, group: 'Overview' },
  { to: '/tickets', label: 'Ticket Assistant', icon: Ticket, group: 'Helpdesk' },
  { to: '/entra', label: 'Entra ID', icon: KeyRound, group: 'Helpdesk' },
  { to: '/exchange', label: 'Exchange Online', icon: Mail, group: 'Helpdesk' },
  { to: '/sharepoint', label: 'SharePoint Online', icon: Boxes, group: 'Helpdesk' },
  { to: '/teams', label: 'Microsoft Teams', icon: Users, group: 'Helpdesk' },
  { to: '/migration', label: 'Migration Projects', icon: FolderKanban, group: 'Migration' },
  { to: '/migration-console', label: 'Migration Console', icon: MonitorPlay, group: 'Migration' },
  { to: '/cross-tenant', label: 'Cross-Tenant Migration', icon: ArrowLeftRight, group: 'Migration' },
  { to: '/hybrid', label: 'Hybrid / On-Prem Connector', icon: Boxes, group: 'Migration' },
  { to: '/bittitan', label: 'BitTitan MigrationWiz', icon: Cloud, group: 'Migration' },
  { to: '/syskit', label: 'Syskit', icon: WorkflowIcon, group: 'Migration' },
  { to: '/provisioning', label: 'User Provisioning', icon: Users, group: 'Operations' },
  { to: '/tenant-admin', label: 'Tenant Administration', icon: SettingsIcon, group: 'Operations' },
  { to: '/monitoring', label: 'Monitoring', icon: Gauge, group: 'Operations' },
  { to: '/security', label: 'Security Hardening', icon: ShieldCheck, group: 'Operations' },
  { to: '/powershell', label: 'PowerShell Generator', icon: Terminal, group: 'Tools' },
  { to: '/mail', label: 'Customer Mail Generator', icon: Send, group: 'Tools' },
  { to: '/chat', label: 'AI Chat Assistant', icon: Bot, group: 'Tools' },
  { to: '/escalation', label: 'Escalation Matrix', icon: Shield, group: 'Reference' },
  { to: '/kb', label: 'Knowledge Base', icon: BookOpen, group: 'Reference' },
  { to: '/notes', label: 'Personal Notes', icon: StickyNote, group: 'Reference' },
  { to: '/settings', label: 'Settings', icon: SettingsIcon, group: 'Reference' },
];

const GROUPS = ['Overview', 'Helpdesk', 'Migration', 'Operations', 'Tools', 'Reference'];

export function Layout({ session, onLogout }: { session: Session; onLogout: () => void }) {
  const [dark, setDark] = useState(() => load('dark-mode', window.matchMedia?.('(prefers-color-scheme: dark)').matches ?? false));
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [query, setQuery] = useState('');
  const branding = getBranding();
  const navigate = useNavigate();

  useEffect(() => {
    document.documentElement.classList.toggle('dark', dark);
    save('dark-mode', dark);
  }, [dark]);

  const results = useMemo(() => {
    if (!query.trim()) return [];
    const q = query.toLowerCase();
    return NAV.filter((n) => n.label.toLowerCase().includes(q)).slice(0, 6);
  }, [query]);

  return (
    <div className="flex min-h-screen bg-slate-50 dark:bg-slate-900">
      {/* Sidebar */}
      <aside className={`fixed inset-y-0 left-0 z-40 w-64 transform border-r border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800 transition-transform lg:translate-x-0 ${sidebarOpen ? 'translate-x-0' : '-translate-x-full'}`}>
        <div className="flex h-16 items-center gap-2.5 border-b border-slate-200 dark:border-slate-700 px-5">
          <div className="grid h-8 w-8 grid-cols-2 grid-rows-2 gap-0.5 rounded p-0.5">
            <div className="rounded-sm bg-[#f25022]" /><div className="rounded-sm bg-[#7fba00]" />
            <div className="rounded-sm bg-[#00a4ef]" /><div className="rounded-sm bg-[#ffb900]" />
          </div>
          <div>
            <div className="text-sm font-bold text-slate-800 dark:text-slate-100 leading-tight">{branding.portalName}</div>
            <div className="text-[10px] text-slate-400 leading-tight">{branding.companyName}</div>
          </div>
          <button className="ml-auto lg:hidden text-slate-400" onClick={() => setSidebarOpen(false)}><X size={18} /></button>
        </div>
        <nav className="h-[calc(100vh-4rem)] overflow-y-auto p-3 pb-10">
          {GROUPS.map((group) => (
            <div key={group} className="mb-3">
              <div className="mb-1 px-3 text-[10px] font-bold uppercase tracking-wider text-slate-400">{group}</div>
              {NAV.filter((n) => n.group === group).map((n) => (
                <NavLink
                  key={n.to}
                  to={n.to}
                  end={n.to === '/'}
                  onClick={() => setSidebarOpen(false)}
                  className={({ isActive }) =>
                    `mb-0.5 flex items-center gap-2.5 rounded-lg px-3 py-2 text-sm font-medium transition-colors ${
                      isActive
                        ? 'bg-blue-600 text-white shadow-sm'
                        : 'text-slate-600 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-700'
                    }`
                  }
                >
                  <n.icon size={16} className="shrink-0" />
                  {n.label}
                </NavLink>
              ))}
            </div>
          ))}
        </nav>
      </aside>

      {sidebarOpen && <div className="fixed inset-0 z-30 bg-black/40 lg:hidden" onClick={() => setSidebarOpen(false)} />}

      {/* Main */}
      <div className="flex min-w-0 flex-1 flex-col lg:pl-64">
        <header className="sticky top-0 z-20 flex h-16 items-center gap-3 border-b border-slate-200 dark:border-slate-700 bg-white/90 dark:bg-slate-800/90 px-4 backdrop-blur lg:px-8">
          <button className="lg:hidden text-slate-500" onClick={() => setSidebarOpen(true)}><Menu size={20} /></button>
          <div className="relative max-w-md flex-1">
            <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
            <input
              value={query}
              onChange={(e) => setQuery(e.target.value)}
              placeholder="Search modules..."
              className="w-full rounded-lg border border-slate-200 dark:border-slate-600 bg-slate-50 dark:bg-slate-900 py-2 pl-9 pr-3 text-sm text-slate-800 dark:text-slate-100 focus:border-blue-500 focus:outline-none"
            />
            {results.length > 0 && (
              <div className="absolute top-full mt-1 w-full rounded-lg border border-slate-200 dark:border-slate-600 bg-white dark:bg-slate-800 py-1 shadow-lg">
                {results.map((r) => (
                  <button
                    key={r.to}
                    onClick={() => { navigate(r.to); setQuery(''); }}
                    className="flex w-full items-center gap-2 px-3 py-2 text-left text-sm text-slate-700 dark:text-slate-200 hover:bg-slate-100 dark:hover:bg-slate-700"
                  >
                    <r.icon size={14} className="text-blue-500" />{r.label}
                  </button>
                ))}
              </div>
            )}
          </div>
          <div className="ml-auto flex items-center gap-2">
            <button
              onClick={() => navigate('/settings')}
              className="hidden sm:flex items-center gap-1.5 rounded-lg border border-slate-200 dark:border-slate-600 px-2.5 py-1.5 text-xs font-medium text-slate-600 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-700"
              title={session.via === 'microsoft' ? 'Signed in via Microsoft 365' : 'Local account — manage in Settings'}
            >
              <UserCircle2 size={15} className="text-emerald-500" />
              {session.email}
            </button>
            <button
              onClick={() => { logout(); onLogout(); }}
              className="rounded-lg border border-slate-200 dark:border-slate-600 p-2 text-slate-500 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-700"
              title="Sign out"
            >
              <LogOut size={16} />
            </button>
            <button
              onClick={() => setDark(!dark)}
              className="rounded-lg border border-slate-200 dark:border-slate-600 p-2 text-slate-500 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-700"
              title="Toggle dark mode"
            >
              {dark ? <Sun size={16} /> : <Moon size={16} />}
            </button>
          </div>
        </header>
        <main className="flex-1 p-4 lg:p-8">
          <Outlet />
        </main>
        <footer className="border-t border-slate-200 dark:border-slate-700 px-8 py-3 text-center text-xs text-slate-400">
          M365 WorkPilot — internal engineer tool. Always verify commands in a test scope before production changes.
        </footer>
      </div>
    </div>
  );
}
