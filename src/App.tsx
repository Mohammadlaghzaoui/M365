import { HashRouter, Route, Routes } from 'react-router-dom';
import { Layout } from './components/Layout';
import Dashboard from './pages/Dashboard';
import Tickets from './pages/Tickets';
import { EntraAssistant, ExchangeAssistant, SharePointAssistant, TeamsAssistant } from './pages/Assistants';
import Migration from './pages/Migration';
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

export default function App() {
  return (
    <HashRouter>
      <Routes>
        <Route element={<Layout />}>
          <Route path="/" element={<Dashboard />} />
          <Route path="/tickets" element={<Tickets />} />
          <Route path="/entra" element={<EntraAssistant />} />
          <Route path="/exchange" element={<ExchangeAssistant />} />
          <Route path="/sharepoint" element={<SharePointAssistant />} />
          <Route path="/teams" element={<TeamsAssistant />} />
          <Route path="/migration" element={<Migration />} />
          <Route path="/cross-tenant" element={<CrossTenant />} />
          <Route path="/hybrid" element={<Hybrid />} />
          <Route path="/bittitan" element={<BitTitan />} />
          <Route path="/syskit" element={<Syskit />} />
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
