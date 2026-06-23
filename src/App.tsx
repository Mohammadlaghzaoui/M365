import { useState } from 'react';
import { HashRouter, Route, Routes } from 'react-router-dom';
import { Layout } from './components/Layout';
import Dashboard from './pages/Dashboard';
import Tickets from './pages/Tickets';
import { EntraAssistant, ExchangeAssistant, SharePointAssistant, TeamsAssistant, EdgeCasesAssistant } from './pages/Assistants';
import Migration from './pages/Migration';
import MigrationConsole from './pages/MigrationConsole';
import GpoAdvisor from './pages/GpoAdvisor';
import GpoBuilder from './pages/GpoBuilder';
import Discovery from './pages/Discovery';
import TenantPortfolio from './pages/TenantPortfolio';
import TenantDetail from './pages/TenantDetail';
import SourceTenantConnect from './pages/SourceTenantConnect';
import Agent from './pages/Agent';
import CrossTenant from './pages/CrossTenant';
import Hybrid from './pages/Hybrid';
import BitTitan from './pages/BitTitan';
import Syskit from './pages/Syskit';
import TenantAdmin from './pages/TenantAdmin';
import Monitoring from './pages/Monitoring';
import Security from './pages/Security';
import PowerShellGen from './pages/PowerShellGen';
import MailGen from './pages/MailGen';
import Chat from './pages/Chat';
import Escalation from './pages/Escalation';
import KnowledgeBase from './pages/KnowledgeBase';
import Notes from './pages/Notes';
import Settings from './pages/Settings';
import Provisioning from './pages/Provisioning';
import Login from './pages/Login';
import { getSession, Session } from './services/auth';
import { useEffect } from 'react';
import { setSaveHook } from './store/useLocalStorage';
import { pullCloud, schedulePush, setSyncUser } from './services/cloudStore';

export default function App() {
  const [session, setSession] = useState<Session | null>(getSession());
  const [hydrating, setHydrating] = useState(false);

  // Cross-device sync: automatic, linked to the signed-in user — no setup.
  useEffect(() => { setSaveHook(schedulePush); return () => setSaveHook(null); }, []);
  useEffect(() => {
    if (!session) return;
    setSyncUser(session.email);
    setHydrating(true);
    pullCloud().catch(() => {}).finally(() => setHydrating(false));
  }, [session]);

  if (!session) {
    return <Login onLogin={setSession} />;
  }

  if (hydrating) {
    return (
      <div className="flex min-h-screen items-center justify-center bg-slate-100 text-slate-500 dark:bg-slate-950">
        <div className="text-center">
          <div className="mx-auto mb-3 h-8 w-8 animate-spin rounded-full border-2 border-slate-300 border-t-blue-600" />
          Loading your assessments…
        </div>
      </div>
    );
  }

  return (
    <HashRouter>
      <Routes>
        <Route element={<Layout session={session} onLogout={() => setSession(null)} />}>
          <Route path="/" element={<Dashboard />} />
          <Route path="/tickets" element={<Tickets />} />
          <Route path="/entra" element={<EntraAssistant />} />
          <Route path="/exchange" element={<ExchangeAssistant />} />
          <Route path="/sharepoint" element={<SharePointAssistant />} />
          <Route path="/teams" element={<TeamsAssistant />} />
          <Route path="/migration" element={<Migration />} />
          <Route path="/migration-console" element={<MigrationConsole />} />
          <Route path="/gpo-advisor" element={<GpoAdvisor />} />
          <Route path="/gpo-builder" element={<GpoBuilder />} />
          <Route path="/discovery" element={<Discovery />} />
          <Route path="/tenants" element={<TenantPortfolio />} />
          <Route path="/tenants/connect" element={<SourceTenantConnect />} />
          <Route path="/tenants/:tenantId" element={<TenantDetail />} />
          <Route path="/agent" element={<Agent />} />
          <Route path="/edge-cases" element={<EdgeCasesAssistant />} />
          <Route path="/cross-tenant" element={<CrossTenant />} />
          <Route path="/hybrid" element={<Hybrid />} />
          <Route path="/bittitan" element={<BitTitan />} />
          <Route path="/syskit" element={<Syskit />} />
          <Route path="/provisioning" element={<Provisioning />} />
          <Route path="/tenant-admin" element={<TenantAdmin />} />
          <Route path="/monitoring" element={<Monitoring />} />
          <Route path="/security" element={<Security />} />
          <Route path="/powershell" element={<PowerShellGen />} />
          <Route path="/mail" element={<MailGen />} />
          <Route path="/chat" element={<Chat />} />
          <Route path="/escalation" element={<Escalation />} />
          <Route path="/kb" element={<KnowledgeBase />} />
          <Route path="/notes" element={<Notes />} />
          <Route path="/settings" element={<Settings />} />
        </Route>
      </Routes>
    </HashRouter>
  );
}
