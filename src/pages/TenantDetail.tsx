import { useMemo, useState } from 'react';
import { useParams, useNavigate, Link } from 'react-router-dom';
import {
  ArrowLeft, Building2, Users2, Server, HardDrive, Monitor, Smartphone, ShieldCheck, ShieldAlert,
  FileSpreadsheet, RefreshCw, Boxes, Gauge, ListChecks, TriangleAlert, Search,
} from 'lucide-react';
import { Badge, Button, Card, ProgressBar } from '../components/ui';
import { DonutChart, HBarChart } from '../components/charts';
import { loadTenantResult } from '../services/tenantStore';
import { assess } from '../services/migrationAssessment';
import { exportDiscoveryWorkbook } from '../services/discoveryExport';
import { save } from '../store/useLocalStorage';
import { DiscoveryResult } from '../services/graphDiscovery';

type Cell = string | number;
function DataTable({ columns, rows, max = 1000 }: { columns: string[]; rows: Cell[][]; max?: number }) {
  if (!rows.length) return <p className="px-1 py-6 text-center text-sm text-slate-400">No data.</p>;
  const shown = rows.slice(0, max);
  return (
    <div className="overflow-auto rounded-lg border border-slate-200 dark:border-slate-700" style={{ maxHeight: 460 }}>
      <table className="w-full border-collapse text-sm">
        <thead className="sticky top-0 z-10 bg-slate-100 dark:bg-slate-800">
          <tr>{columns.map((c) => <th key={c} className="whitespace-nowrap px-3 py-2 text-left font-semibold text-slate-600 dark:text-slate-300">{c}</th>)}</tr>
        </thead>
        <tbody>
          {shown.map((r, i) => (
            <tr key={i} className="border-t border-slate-100 dark:border-slate-700/60 hover:bg-slate-50 dark:hover:bg-slate-800/50">
              {r.map((cell, j) => <td key={j} className="whitespace-nowrap px-3 py-1.5 text-slate-700 dark:text-slate-200">{String(cell)}</td>)}
            </tr>
          ))}
        </tbody>
      </table>
      {rows.length > max && <div className="bg-slate-50 dark:bg-slate-800/50 px-3 py-2 text-center text-xs text-slate-400">Showing {max} of {rows.length} rows — export to Excel for the full set.</div>}
    </div>
  );
}

function Panel({ title, sub, children }: { title: string; sub?: string; children: React.ReactNode }) {
  return (
    <Card className="mb-5 p-5">
      <div className="mb-3"><h3 className="font-semibold text-slate-800 dark:text-slate-100">{title}</h3>{sub && <p className="text-xs text-slate-400">{sub}</p>}</div>
      {children}
    </Card>
  );
}

const TABS = ['Overview', 'Users', 'Devices', 'Licenses', 'Groups & Teams', 'Security & Apps', 'Migration'] as const;
type Tab = typeof TABS[number];

export default function TenantDetail() {
  const { tenantId = '' } = useParams();
  const nav = useNavigate();
  const result = useMemo<DiscoveryResult | null>(() => loadTenantResult(tenantId), [tenantId]);
  const analysis = useMemo(() => (result ? assess(result) : null), [result]);
  const [tab, setTab] = useState<Tab>('Overview');
  const [uq, setUq] = useState('');
  const [dq, setDq] = useState('');

  if (!result) {
    return (
      <div className="p-8 text-center">
        <p className="mb-3 text-slate-500">This tenant's assessment was not found (it may have been wiped).</p>
        <Button onClick={() => nav('/tenants')}><ArrowLeft size={15} /> Back to portfolio</Button>
      </div>
    );
  }

  const r = result;
  const fmtData = (gb: number) => (gb >= 1024 ? `${(gb / 1024).toFixed(1)} TB` : `${Math.round(gb)} GB`);
  const totalGB = r.usage.mailboxTotalGB + r.usage.oneDriveTotalGB + r.usage.spoTotalGB;
  const reRun = () => { save('discovery-result', r); nav('/discovery'); };

  const licenseChart = r.licenses.slice(0, 8).map((l) => ({ label: l.skuPartNumber, value: l.consumed }));
  const deviceChart = Object.entries(r.devices.byOs).map(([label, value]) => ({ label, value }));
  const deviceValidation = r.validations.find((v) => /device|intune/i.test(v.workload));

  const kpis = [
    { icon: Building2, label: 'Tenant', value: r.org.displayName, sub: `${r.org.verifiedDomains} verified domains`, color: 'text-blue-600 dark:text-blue-400' },
    { icon: Users2, label: 'Users', value: r.users.length, sub: `${r.guests} guests · ${r.disabled} disabled`, color: 'text-violet-600 dark:text-violet-400' },
    { icon: Boxes, label: 'Groups / Teams', value: r.groups.length, sub: `${r.m365Groups} M365 · ${r.teams} Teams`, color: 'text-emerald-600 dark:text-emerald-400' },
    { icon: Server, label: 'Devices', value: r.devices.total, sub: `${r.devices.compliant ?? 0} compliant`, color: 'text-cyan-600 dark:text-cyan-400' },
    { icon: HardDrive, label: 'Total data', value: fmtData(totalGB), sub: `${r.usage.mailboxCount} mailboxes`, color: 'text-amber-600 dark:text-amber-400' },
  ];

  const userRows = r.users
    .filter((u) => !uq || `${u.displayName} ${u.userPrincipalName} ${u.department} ${u.userCategory}`.toLowerCase().includes(uq.toLowerCase()))
    .map((u) => [u.displayName, u.userPrincipalName, u.userType, u.accountEnabled ? 'Yes' : 'No', u.department, u.licenses, u.workerType ?? '', u.userCategory ?? '', u.appPlatforms ?? '', u.lastOfficeActivity ?? '', u.deviceCount ?? 0, u.lastSignIn] as Cell[]);

  const deviceRows = (r.deviceInventory ?? [])
    .filter((d) => !dq || `${d.deviceName} ${d.user} ${d.os} ${d.model} ${d.suggestedName}`.toLowerCase().includes(dq.toLowerCase()))
    .map((d) => [d.deviceName, d.user, d.os, d.osVersion, d.formFactor ?? '', d.typeCode ?? '', d.suggestedName ?? '', d.compliance, d.ownership, d.model, d.serialNumber, d.source ?? '', d.lastSync] as Cell[]);

  const findingColor: Record<string, string> = { info: 'blue', warn: 'amber', risk: 'orange', blocker: 'red' };

  return (
    <div>
      {/* Header */}
      <div className="mb-5">
        <Link to="/tenants" className="mb-2 inline-flex items-center gap-1 text-sm text-slate-500 hover:text-blue-600"><ArrowLeft size={14} /> Tenant Portfolio</Link>
        <div className="flex flex-wrap items-center gap-3">
          <div className="flex h-12 w-12 items-center justify-center rounded-xl bg-gradient-to-br from-blue-500 to-indigo-600 text-base font-bold text-white shadow">
            {r.org.displayName.split(/\s+/).map((w) => w[0]).join('').slice(0, 2).toUpperCase()}
          </div>
          <div className="mr-auto">
            <h1 className="text-2xl font-bold text-slate-800 dark:text-slate-100">{r.org.displayName}</h1>
            <div className="flex items-center gap-2 text-xs text-slate-400">
              <span className="font-mono">{r.org.tenantId}</span>
              <Badge color="green">Read-only · no tenant changes</Badge>
              <span>Assessed {new Date(r.fetchedAt).toLocaleString()}</span>
            </div>
          </div>
          <Button variant="secondary" onClick={reRun}><RefreshCw size={15} /> Re-run</Button>
          <Button onClick={() => exportDiscoveryWorkbook(r, analysis)}><FileSpreadsheet size={15} /> Export Excel</Button>
        </div>
      </div>

      {/* KPI row */}
      <div className="mb-5 grid grid-cols-2 gap-3 lg:grid-cols-5">
        {kpis.map((k) => (
          <Card key={k.label} className="p-4">
            <div className="flex items-center gap-2 text-slate-400"><k.icon size={16} className={k.color} /><span className="text-xs">{k.label}</span></div>
            <div className="mt-1 truncate text-xl font-bold text-slate-800 dark:text-slate-100" title={String(k.value)}>{k.value}</div>
            <div className="text-xs text-slate-400">{k.sub}</div>
          </Card>
        ))}
      </div>

      {/* Tabs */}
      <div className="mb-5 flex flex-wrap gap-1 border-b border-slate-200 dark:border-slate-700">
        {TABS.map((t) => (
          <button key={t} onClick={() => setTab(t)}
            className={`-mb-px border-b-2 px-3 py-2 text-sm font-medium transition-colors ${tab === t ? 'border-blue-600 text-blue-600 dark:text-blue-400' : 'border-transparent text-slate-500 hover:text-slate-700 dark:hover:text-slate-300'}`}>
            {t}
          </button>
        ))}
      </div>

      {/* ---- Overview ---- */}
      {tab === 'Overview' && (
        <div>
          <div className="mb-5 grid grid-cols-1 gap-5 lg:grid-cols-2">
            <Panel title="User composition">
              <DonutChart centerLabel="total users" centerValue={String(r.users.length)} segments={[
                { label: 'Active members', value: r.users.length - r.guests - r.disabled, color: '#10b981' },
                { label: 'Guests', value: r.guests, color: '#8b5cf6' },
                { label: 'Disabled', value: r.disabled, color: '#94a3b8' },
              ]} />
            </Panel>
            <Panel title="License consumption (top)">
              {licenseChart.length ? <HBarChart data={licenseChart} /> : <p className="text-sm text-slate-400">No licenses found.</p>}
            </Panel>
          </div>

          <div className="mb-5 grid grid-cols-1 gap-5 lg:grid-cols-3">
            <Card className="p-5">
              <div className="mb-1 flex items-center gap-2 text-slate-400"><Monitor size={16} className="text-cyan-500" /><span className="text-xs">Office workers</span></div>
              <div className="text-3xl font-bold text-slate-800 dark:text-slate-100">{r.appUsage?.office ?? 0}</div>
              <div className="text-xs text-slate-400">desktop/laptop — sign in with the Office app ({r.appUsage?.desktopApp ?? 0} confirmed by usage)</div>
            </Card>
            <Card className="p-5">
              <div className="mb-1 flex items-center gap-2 text-slate-400"><Smartphone size={16} className="text-violet-500" /><span className="text-xs">Field workers</span></div>
              <div className="text-3xl font-bold text-slate-800 dark:text-slate-100">{r.appUsage?.field ?? 0}</div>
              <div className="text-xs text-slate-400">mobile-only — use Microsoft 365 only on a phone/tablet</div>
            </Card>
            <Card className="p-5">
              <div className="mb-1 flex items-center gap-2 text-slate-400">{(r.devices.nonCompliant ?? 0) > 0 ? <ShieldAlert size={16} className="text-rose-500" /> : <ShieldCheck size={16} className="text-emerald-500" />}<span className="text-xs">Device compliance</span></div>
              <div className="text-3xl font-bold text-slate-800 dark:text-slate-100">{r.devices.compliant ?? 0}<span className="text-base font-normal text-slate-400"> / {r.devices.total}</span></div>
              <div className="text-xs text-slate-400">{r.devices.nonCompliant ?? 0} non-compliant · {(r.compliancePolicies ?? []).length} policies</div>
            </Card>
          </div>

          {analysis && (
            <Panel title="Migration analysis" sub={`Complexity ${analysis.complexityLabel} · ~${analysis.estimatedDays} working days · ~${fmtData(analysis.totalDataGB)}`}>
              <div className="mb-3 flex items-center gap-3">
                <Gauge size={18} className="text-blue-600" />
                <div className="flex-1"><ProgressBar value={analysis.complexityScore} color={analysis.complexityScore >= 75 ? 'bg-rose-500' : analysis.complexityScore >= 50 ? 'bg-amber-500' : 'bg-emerald-500'} /></div>
                <span className="text-sm font-bold text-slate-700 dark:text-slate-200">{analysis.complexityScore}/100</span>
              </div>
              <p className="mb-3 rounded-lg bg-blue-50 dark:bg-blue-900/20 p-3 text-sm text-blue-800 dark:text-blue-200">{analysis.recommendedApproach}</p>
              <div className="space-y-1.5">
                {analysis.findings.map((f, i) => (
                  <div key={i} className="flex items-start gap-2 text-sm">
                    <Badge color={findingColor[f.level] ?? 'gray'}>{f.level}</Badge>
                    <span className="text-slate-600 dark:text-slate-300"><strong>{f.area}:</strong> {f.text}</span>
                  </div>
                ))}
              </div>
            </Panel>
          )}
        </div>
      )}

      {/* ---- Users ---- */}
      {tab === 'Users' && (
        <Panel title={`Users (${r.users.length})`} sub="Includes who logs in with the Office desktop app vs mobile only, licenses and devices.">
          <div className="mb-3 flex items-center gap-2 rounded-lg border border-slate-200 dark:border-slate-700 px-3 py-1.5">
            <Search size={15} className="text-slate-400" />
            <input value={uq} onChange={(e) => setUq(e.target.value)} placeholder="Search users, department, category…" className="w-full bg-transparent text-sm outline-none" />
          </div>
          <DataTable columns={['Display name', 'UPN', 'Type', 'Enabled', 'Department', 'Licenses', 'Worker type', 'Category', 'App platforms', 'Last Office', 'Devices', 'Last sign-in']} rows={userRows} />
        </Panel>
      )}

      {/* ---- Devices ---- */}
      {tab === 'Devices' && (
        <div>
          {(r.deviceInventory?.length ?? 0) === 0 && (
            <Card className="mb-5 border-amber-300 dark:border-amber-700 p-4">
              <div className="flex items-start gap-2 text-sm text-amber-700 dark:text-amber-300">
                <TriangleAlert size={16} className="mt-0.5 shrink-0" />
                <span><strong>No managed devices returned.</strong> {deviceValidation ? `Reason: ${deviceValidation.note}` : 'This tenant has no Intune-managed devices, or the sign-in account lacked the DeviceManagementManagedDevices.Read.All scope. Re-run and consent to Intune scopes if devices are expected.'}</span>
              </div>
            </Card>
          )}
          <div className="mb-5 grid grid-cols-2 gap-3 lg:grid-cols-4">
            <Card className="p-4"><div className="text-xs text-slate-400">Total devices</div><div className="text-2xl font-bold text-slate-800 dark:text-slate-100">{r.devices.total}</div></Card>
            <Card className="p-4"><div className="text-xs text-slate-400">Compliant</div><div className="text-2xl font-bold text-emerald-600">{r.devices.compliant ?? 0}</div></Card>
            <Card className="p-4"><div className="text-xs text-slate-400">Non-compliant</div><div className="text-2xl font-bold text-rose-600">{r.devices.nonCompliant ?? 0}</div></Card>
            <Card className="p-4"><div className="text-xs text-slate-400">Compliance policies</div><div className="text-2xl font-bold text-slate-800 dark:text-slate-100">{(r.compliancePolicies ?? []).length}</div></Card>
          </div>
          {deviceChart.length > 0 && <Panel title="Devices by operating system"><HBarChart data={deviceChart} color="#0891b2" /></Panel>}
          <Panel title={`Device inventory (${r.deviceInventory?.length ?? 0})`} sub="With the workstation naming convention applied (form factor, type code, suggested name).">
            <div className="mb-3 flex items-center gap-2 rounded-lg border border-slate-200 dark:border-slate-700 px-3 py-1.5">
              <Search size={15} className="text-slate-400" />
              <input value={dq} onChange={(e) => setDq(e.target.value)} placeholder="Search device, user, model…" className="w-full bg-transparent text-sm outline-none" />
            </div>
            <DataTable columns={['Device', 'User', 'OS', 'Version', 'Form factor', 'Code', 'Suggested name', 'Compliance', 'Ownership', 'Model', 'Serial', 'Source', 'Last sync']} rows={deviceRows} />
          </Panel>
          <Panel title="Workstation naming convention" sub="From your Workstation Naming sheet — format: <SITE><n>-<WorkerType><DeviceType>-<Serial> (e.g. AHA1-OW-BC349BC34).">
            <div className="grid grid-cols-1 gap-4 md:grid-cols-2">
              <div>
                <div className="mb-1 text-xs font-semibold uppercase text-slate-400">Worker type</div>
                <DataTable columns={['Type', 'Code']} rows={[['Field', 'F'], ['Office', 'O'], ['Temp', 'T'], ['Kiosk / Common / Shared', 'K']]} />
              </div>
              <div>
                <div className="mb-1 text-xs font-semibold uppercase text-slate-400">Device type</div>
                <DataTable columns={['Type', 'Code']} rows={[['Office Laptop', 'L'], ['Office Desktop', 'D'], ['Engineering Laptop', 'W'], ['Engineering Desktop', 'X'], ['Executive Laptop', 'E'], ['Executive Desktop', 'F'], ['Mac Computers', 'M'], ['Server', 'S'], ['Network', 'N'], ['Appliance / IOT', 'A'], ['Phone', 'P'], ['Tablet', 'T']]} />
              </div>
            </div>
          </Panel>
        </div>
      )}

      {/* ---- Licenses ---- */}
      {tab === 'Licenses' && (
        <Panel title={`Licenses (${r.licenses.length} SKUs)`} sub={`${r.licenses.reduce((a, l) => a + l.consumed, 0)} seats consumed across the tenant.`}>
          <div className="space-y-2">
            {r.licenses.map((l) => (
              <div key={l.skuPartNumber} className="flex items-center gap-3">
                <div className="w-56 truncate text-sm font-medium text-slate-700 dark:text-slate-200" title={l.skuPartNumber}>{l.skuPartNumber}</div>
                <div className="flex-1"><ProgressBar value={l.enabled ? (l.consumed / l.enabled) * 100 : 0} color={l.available <= 0 ? 'bg-rose-500' : 'bg-blue-600'} /></div>
                <div className="w-24 text-right text-sm text-slate-500">{l.consumed}/{l.enabled}</div>
                <div className="w-20 text-right text-xs text-slate-400">{l.available} free</div>
              </div>
            ))}
          </div>
        </Panel>
      )}

      {/* ---- Groups & Teams ---- */}
      {tab === 'Groups & Teams' && (
        <div>
          <Panel title={`Groups (${r.groups.length})`}>
            <DataTable columns={['Display name', 'Mail', 'Type', 'Membership', 'Visibility', 'Is Team']} rows={r.groups.map((g) => [g.displayName, g.mail, g.groupType, g.membershipType, g.visibility, g.isTeam ? 'Yes' : 'No'])} />
          </Panel>
          <Panel title={`Teams detail (${r.teamsDetail.length})`}>
            <DataTable columns={['Team', 'Visibility', 'Owners', 'Members', 'Guests', 'Channels', 'Private/shared']} rows={r.teamsDetail.map((t) => [t.name, t.visibility, t.owners, t.members, t.guests, t.channels, t.privateChannels])} />
          </Panel>
        </div>
      )}

      {/* ---- Security & Apps ---- */}
      {tab === 'Security & Apps' && (
        <div>
          <Panel title={`Conditional Access (${r.caPolicies.length})`}>
            <DataTable columns={['Policy', 'State']} rows={r.caPolicies.map((p) => [p.displayName, p.state])} />
          </Panel>
          <Panel title="SharePoint & sharing" sub={`${r.sharePointSites.length} sites · external sharing exposure (sampled).`}>
            <div className="mb-3 grid grid-cols-3 gap-3 text-center">
              <div className="rounded-lg bg-rose-50 dark:bg-rose-900/20 p-3"><div className="text-xl font-bold text-rose-600">{r.sharing.anonymous}</div><div className="text-xs text-slate-500">"Anyone" links</div></div>
              <div className="rounded-lg bg-amber-50 dark:bg-amber-900/20 p-3"><div className="text-xl font-bold text-amber-600">{r.sharing.organization}</div><div className="text-xs text-slate-500">Org-wide links</div></div>
              <div className="rounded-lg bg-blue-50 dark:bg-blue-900/20 p-3"><div className="text-xl font-bold text-blue-600">{r.sharing.users}</div><div className="text-xs text-slate-500">Specific users</div></div>
            </div>
            <DataTable columns={['Site', 'URL', 'Created', 'Last modified']} rows={r.sharePointSites.map((s) => [s.name, s.webUrl, s.created, s.lastModified])} max={200} />
          </Panel>
          <Panel title={`App registrations (${r.appRegistrations.length})`}>
            <DataTable columns={['Display name', 'App ID', 'Audience', 'Created']} rows={r.appRegistrations.map((a) => [a.displayName, a.appId, a.signInAudience, a.created])} />
          </Panel>
          <Panel title={`Enterprise apps (${r.servicePrincipals.length})`}>
            <DataTable columns={['Display name', 'App ID', 'Type', 'Enabled']} rows={r.servicePrincipals.map((s) => [s.displayName, s.appId, s.type, s.enabled ? 'Yes' : 'No'])} max={300} />
          </Panel>
        </div>
      )}

      {/* ---- Migration ---- */}
      {tab === 'Migration' && analysis && (
        <div>
          <Panel title="Data sizing">
            <DataTable columns={['Workload', 'Total GB', 'Count', 'Notes']} rows={[
              ['Mailboxes', r.usage.mailboxTotalGB, r.usage.mailboxCount, `${r.usage.mailboxOver50GB} over 50GB, ${r.usage.archiveCount} archives`],
              ['OneDrive', r.usage.oneDriveTotalGB, r.usage.oneDriveCount, `${r.usage.oneDriveOver100GB} over 100GB`],
              ['SharePoint', r.usage.spoTotalGB, r.usage.spoSiteCount, ''],
            ]} />
          </Panel>
          <Panel title="Workload readiness">
            <DataTable columns={['Workload', 'Status', 'Count', 'Note']} rows={r.workloads.map((w) => [w.workload, w.status, w.count, w.note])} />
          </Panel>
          <Panel title={`Preparation checklist (${analysis.checklist.length})`}>
            <div className="space-y-1.5">
              {analysis.checklist.map((c, i) => (
                <div key={i} className="flex items-start gap-2 text-sm"><ListChecks size={15} className="mt-0.5 shrink-0 text-blue-500" /><span className="text-slate-600 dark:text-slate-300"><strong>{c.phase}:</strong> {c.task}</span></div>
              ))}
            </div>
          </Panel>
          {r.validations.length > 0 && (
            <Panel title={`Validation items (${r.validations.length})`} sub="Anything Graph could not fully read, and why.">
              <DataTable columns={['Workload', 'Status', 'Note']} rows={r.validations.map((v) => [v.workload, v.status, v.note])} />
            </Panel>
          )}
        </div>
      )}
    </div>
  );
}
