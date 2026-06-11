import { useState } from 'react';
import { HashRouter, Route, Routes } from 'react-router-dom';
import { Layout } from './components/Layout';
import Dashboard from './pages/Dashboard';
import Tickets from './pages/Tickets';
import { EntraAssistant, ExchangeAssistant, SharePointAssistant, TeamsAssistant, EdgeCasesAssistant } from './pages/Assistants';
import Migration from './pages/Migration';
import MigrationConsole from './pages/MigrationConsole';
import GpoAdvisor from './pages/GpoAdvisor';
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

export default function App() {
  const [session, setSession] = useState<Session | null>(getSession());

  if (!session) {
    return <Login onLogin={setSession} />;
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
