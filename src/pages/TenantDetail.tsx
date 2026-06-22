import { useEffect, useMemo, useState } from 'react';
import { useParams, useNavigate, Link } from 'react-router-dom';
import {
  ArrowLeft, ArrowRight, FileSpreadsheet, RefreshCw, Building2, Users2, Server, HardDrive,
  Mail, ShieldCheck, ShieldAlert, Layers,
} from 'lucide-react';
import { loadTenantResult } from '../services/tenantStore';
import { assess } from '../services/migrationAssessment';
import { exportDiscoveryWorkbook } from '../services/discoveryExport';
import { save } from '../store/useLocalStorage';
import { DiscoveryResult } from '../services/graphDiscovery';
import {
  getSiteCode, setSiteCode, getMigrationTarget, setMigrationTarget, NAMING_FORMAT,
  getTargetDomain, setTargetDomain, mapUpnToTarget,
  endpointClassOf, deviceTypeLabel, suggestedName, proposedLicenseOf, LICENSE_RULES,
} from '../services/naming';

type Tone = 'green' | 'amber' | 'red' | 'blue' | 'slate';
type Cell = React.ReactNode;

const TONES: Record<Tone, string> = {
  green: 'bg-emerald-50 text-emerald-700 ring-emerald-600/20 dark:bg-emerald-900/30 dark:text-emerald-300 dark:ring-emerald-400/20',
  amber: 'bg-amber-50 text-amber-700 ring-amber-600/20 dark:bg-amber-900/30 dark:text-amber-300 dark:ring-amber-400/20',
  red: 'bg-rose-50 text-rose-700 ring-rose-600/20 dark:bg-rose-900/30 dark:text-rose-300 dark:ring-rose-400/20',
  blue: 'bg-blue-50 text-blue-700 ring-blue-600/20 dark:bg-blue-900/30 dark:text-blue-300 dark:ring-blue-400/20',
  slate: 'bg-slate-100 text-slate-600 ring-slate-500/15 dark:bg-slate-700/50 dark:text-slate-300 dark:ring-slate-400/20',
};
function Chip({ label, tone }: { label: string; tone: Tone }) {
  return <span className={`inline-flex items-center rounded px-2 py-0.5 text-xs font-medium ring-1 ring-inset ${TONES[tone]}`}>{label}</span>;
}
const complianceTone = (s: string): Tone => (/^(compliant)$/i.test(s) ? 'green' : /noncompliant|notcompliant/i.test(s) ? 'red' : 'slate');
const workerTone = (s: string): Tone => (s === 'Office' ? 'blue' : s === 'Field' ? 'amber' : 'slate');
const mailboxTone = (s: string): Tone => (s === 'User' ? 'green' : /shared/i.test(s) ? 'amber' : 'slate');
const licenseTone = (s: string): Tone => (/business premium/i.test(s) ? 'blue' : /exchange/i.test(s) ? 'green' : /validate/i.test(s) ? 'amber' : 'slate');

function DataTable({ columns, rows, max = 1000, align }: { columns: string[]; rows: Cell[][]; max?: number; align?: Record<number, 'right'> }) {
  if (!rows.length) return <p className="rounded border border-dashed border-slate-300 px-1 py-6 text-center text-sm text-slate-400 dark:border-slate-700">No rows.</p>;
  const shown = rows.slice(0, max);
  return (
    <div className="overflow-auto rounded-md border border-slate-300 shadow-sm dark:border-slate-700" style={{ maxHeight: 540 }}>
      <table className="w-full border-collapse text-[13px]">
        <thead className="sticky top-0 z-10 bg-slate-100 dark:bg-slate-800">
          <tr>{columns.map((c, i) => <th key={c} className={`whitespace-nowrap border-b border-slate-300 px-3 py-2 font-semibold text-slate-700 dark:border-slate-700 dark:text-slate-200 ${align?.[i] === 'right' ? 'text-right' : 'text-left'}`}>{c}</th>)}</tr>
        </thead>
        <tbody>
          {shown.map((r, i) => (
            <tr key={i} className="odd:bg-white even:bg-slate-50/70 hover:bg-blue-50/40 dark:odd:bg-slate-900 dark:even:bg-slate-800/40 dark:hover:bg-slate-800">
              {r.map((cell, j) => <td key={j} className={`whitespace-nowrap border-b border-slate-100 px-3 py-1.5 text-slate-700 dark:border-slate-800 dark:text-slate-200 ${align?.[j] === 'right' ? 'text-right tabular-nums' : ''}`}>{cell}</td>)}
            </tr>
          ))}
        </tbody>
      </table>
      {rows.length > max && <div className="border-t border-slate-200 bg-slate-50 px-3 py-2 text-center text-xs text-slate-400 dark:border-slate-700 dark:bg-slate-800/50">Showing {max} of {rows.length} rows — export to Excel for the full set.</div>}
    </div>
  );
}

function Section({ id, n, title, desc, children }: { id: string; n: number; title: string; desc?: string; children: React.ReactNode }) {
  return (
    <section id={id} className="mb-6 scroll-mt-32 rounded-lg border border-slate-300 bg-white shadow-sm dark:border-slate-700 dark:bg-slate-900">
      <div className="flex items-start gap-3 border-b border-slate-200 px-5 py-3.5 dark:border-slate-700">
        <span className="mt-0.5 flex h-6 w-6 shrink-0 items-center justify-center rounded-full bg-slate-800 text-xs font-bold text-white dark:bg-slate-200 dark:text-slate-900">{n}</span>
        <div>
          <h2 className="text-base font-semibold leading-tight text-slate-900 dark:text-slate-100">{title}</h2>
          {desc && <p className="mt-0.5 text-xs text-slate-500">{desc}</p>}
        </div>
      </div>
      <div className="p-5">{children}</div>
    </section>
  );
}

const SECTIONS = [
  { id: 'overview', label: 'Overview' },
  { id: 'users', label: 'Users & Licensing' },
  { id: 'endpoints', label: 'Endpoint Plan' },
  { id: 'inventory', label: 'Inventory' },
  { id: 'testmig', label: 'Test Migration' },
  { id: 'risks', label: 'Risks & Issues' },
  { id: 'evidence', label: 'Evidence' },
];

export default function TenantDetail() {
  const { tenantId = '' } = useParams();
  const nav = useNavigate();
  const result = useMemo<DiscoveryResult | null>(() => loadTenantResult(tenantId), [tenantId]);
  const analysis = useMemo(() => (result ? assess(result) : null), [result]);
  const [site, setSite] = useState(() => getSiteCode(tenantId));
  const [target, setTarget] = useState(() => getMigrationTarget(tenantId));
  const [targetDomain, setTargetDomainState] = useState(() => getTargetDomain(tenantId));
  const [pilotSize, setPilotSize] = useState(5);
  const [active, setActive] = useState('overview');

  useEffect(() => {
    const obs = new IntersectionObserver(
      (entries) => { entries.forEach((e) => { if (e.isIntersecting) setActive(e.target.id); }); },
      { rootMargin: '-30% 0px -60% 0px', threshold: 0 },
    );
    SECTIONS.forEach((s) => { const el = document.getElementById(s.id); if (el) obs.observe(el); });
    return () => obs.disconnect();
  }, [result]);

  const workerByUser = useMemo(() => new Map((result?.users ?? []).map((u) => [u.userPrincipalName.toLowerCase(), u.workerType] as const)), [result]);

  if (!result) {
    return (
      <div className="p-8 text-center">
        <p className="mb-3 text-slate-500">This source tenant's assessment was not found (it may have been wiped).</p>
        <Link to="/tenants" className="inline-flex items-center gap-1 text-sm text-blue-700 hover:underline"><ArrowLeft size={15} /> Back to Source Tenants</Link>
      </div>
    );
  }
  const r = result;

  // ---- Endpoint plan ----
  const endpoints = (r.deviceInventory ?? []).map((d) => {
    const wt = workerByUser.get((d.user || '').toLowerCase()) ?? '';
    const sug = suggestedName(site, d, wt);
    return { device: d.deviceName, cls: endpointClassOf(d.formFactor), os: `${d.os}${d.osVersion ? ' ' + d.osVersion : ''}`.trim(), user: d.user || '—', suggested: sug, type: deviceTypeLabel(d.formFactor), compliance: d.compliance, action: sug ? 'Rename to suggested' : 'Validate manually' };
  });
  const renameSuggestions = endpoints.filter((e) => e.suggested).length;
  const namingErrors = (() => { const seen = new Set<string>(); let dup = 0; for (const e of endpoints) { if (!e.suggested) continue; if (seen.has(e.suggested)) dup++; else seen.add(e.suggested); } return dup; })();

  const firstDeviceOf = (upn: string) => { const d = (r.deviceInventory ?? []).find((x) => (x.user || '').toLowerCase() === upn.toLowerCase()); return d ? suggestedName(site, d, workerByUser.get(upn.toLowerCase()) ?? '') : ''; };
  const licCount = (m: string) => r.users.filter((u) => u.accountEnabled && u.userType !== 'Guest' && (u.workerType || '') === m).length;
  const licenseSummary = LICENSE_RULES.map((x) => ({ ...x, count: licCount(x.match) }));

  const candidates = r.users.filter((u) => u.accountEnabled && u.userType !== 'Guest');
  const pilot = (() => {
    const office = candidates.filter((u) => u.workerType === 'Office');
    const field = candidates.filter((u) => u.workerType === 'Field');
    const other = candidates.filter((u) => !u.workerType);
    const mix = [...office.slice(0, Math.ceil(pilotSize / 2)), ...field.slice(0, 1), ...other.slice(0, pilotSize)];
    const seen = new Set<string>();
    return mix.filter((u) => (seen.has(u.userPrincipalName) ? false : (seen.add(u.userPrincipalName), true))).slice(0, pilotSize);
  })();

  const totalGB = r.usage.mailboxTotalGB + r.usage.oneDriveTotalGB + r.usage.spoTotalGB;
  const fmtGB = (gb: number) => (gb >= 1024 ? `${(gb / 1024).toFixed(1)} TB` : `${Math.round(gb)} GB`);
  const readiness = r.workloads.length ? Math.round((r.workloads.filter((w) => w.status === 'Exported').length / r.workloads.length) * 100) : 0;

  const kpis = [
    { icon: Users2, label: 'Users', value: r.users.length, sub: `${r.guests} guests · ${r.disabled} disabled` },
    { icon: Mail, label: 'Mailboxes', value: r.usage.mailboxCount, sub: `${r.usage.mailboxTotalGB.toFixed(0)} GB` },
    { icon: Building2, label: 'Groups / Teams', value: `${r.groups.length} / ${r.teams}`, sub: `${r.m365Groups} M365 groups` },
    { icon: Server, label: 'Devices', value: r.devices.total, sub: `${r.devices.compliant ?? 0} compliant` },
    { icon: HardDrive, label: 'Total data', value: fmtGB(totalGB), sub: `${r.sharePointSites.length} SharePoint sites` },
    { icon: Layers, label: 'Rename plan', value: renameSuggestions, sub: `${namingErrors} naming errors` },
  ];

  const overviewRows: Cell[][] = [
    ['Source tenant', r.org.displayName], ['Migration target', target], ['Tenant ID', r.org.tenantId],
    ['Users', r.users.length], ['Mailboxes', r.usage.mailboxCount], ['Mailbox data (GB)', r.usage.mailboxTotalGB.toFixed(2)],
    ['Groups', r.groups.length], ['Teams', r.teams], ['SharePoint sites', r.sharePointSites.length],
    ['Devices', r.devices.total], ['Rename suggestions', renameSuggestions], ['Naming errors', namingErrors],
    ['Assessed', new Date(r.fetchedAt).toLocaleString()],
  ];

  const findingTone: Record<string, string> = { blocker: 'text-rose-700 dark:text-rose-400', risk: 'text-orange-700 dark:text-orange-400', warn: 'text-amber-700 dark:text-amber-400', info: 'text-slate-600 dark:text-slate-300' };
  const findingChip: Record<string, Tone> = { blocker: 'red', risk: 'red', warn: 'amber', info: 'slate' };

  const evidence: Cell[][] = [
    ['Organisation / domains', 'GET /organization, /domains', <Chip label="Collected" tone="green" />, r.domains.length],
    ['Users', 'GET /users (+manager, proxyAddresses)', <Chip label="Collected" tone="green" />, r.users.length],
    ['Licenses', 'GET /subscribedSkus', <Chip label="Collected" tone="green" />, r.licenses.length],
    ['Groups & Teams', 'GET /groups', <Chip label="Collected" tone="green" />, r.groups.length],
    ['Devices (Intune)', 'GET /deviceManagement/managedDevices', r.deviceInventory.some((d) => d.source === 'Intune') ? <Chip label="Collected" tone="green" /> : <Chip label="None" tone="slate" />, r.deviceInventory.filter((d) => d.source === 'Intune').length],
    ['Devices (Entra)', 'GET /devices', r.deviceInventory.some((d) => d.source === 'Entra') ? <Chip label="Collected" tone="green" /> : <Chip label="None" tone="slate" />, r.deviceInventory.filter((d) => d.source === 'Entra').length],
    ['Office app usage', 'GET /reports/getM365AppUserDetail', r.appUsage?.classified ? <Chip label="Collected" tone="green" /> : <Chip label="Not available" tone="amber" />, r.appUsage?.classified ?? 0],
    ['Mailbox sizing', 'GET /reports/getMailboxUsageDetail', r.usage.available ? <Chip label="Collected" tone="green" /> : <Chip label="Not available" tone="amber" />, r.usage.mailboxCount],
    ['SharePoint sites', 'GET /sites', <Chip label="Collected" tone="green" />, r.sharePointSites.length],
    ['Conditional Access', 'GET /identity/conditionalAccess/policies', <Chip label="Collected" tone="green" />, r.caPolicies.length],
    ['Admin roles', 'GET /directoryRoles?$expand=members', (r.adminRoles ?? []).length ? <Chip label="Collected" tone="green" /> : <Chip label="None" tone="slate" />, (r.adminRoles ?? []).reduce((a, x) => a + x.members.length, 0)],
  ];

  return (
    <div className="mx-auto max-w-[1240px] pb-10">
      <Link to="/tenants" className="mb-3 inline-flex items-center gap-1 text-sm text-slate-500 hover:text-blue-700"><ArrowLeft size={14} /> Source Tenants</Link>

      {/* Header band — source → target */}
      <div className="mb-5 rounded-lg border border-slate-300 bg-white p-5 shadow-sm dark:border-slate-700 dark:bg-slate-900">
        <div className="flex flex-wrap items-center gap-4">
          <div className="flex items-center gap-3">
            <div className="flex h-12 w-12 items-center justify-center rounded-md bg-slate-800 text-base font-bold text-white">{r.org.displayName.slice(0, 2).toUpperCase()}</div>
            <div>
              <div className="text-xs font-medium uppercase tracking-wide text-slate-400">Source tenant</div>
              <div className="text-lg font-bold leading-tight text-slate-900 dark:text-slate-100">{r.org.displayName}</div>
            </div>
          </div>
          <ArrowRight size={22} className="text-slate-300 dark:text-slate-600" />
          <div className="flex items-center gap-2">
            <div className="flex h-12 w-12 items-center justify-center rounded-md bg-blue-700 text-base font-bold text-white">{(target || 'TG').slice(0, 2).toUpperCase()}</div>
            <div>
              <div className="text-xs font-medium uppercase tracking-wide text-slate-400">Migration target</div>
              <input value={target} onChange={(e) => { setTarget(e.target.value); setMigrationTarget(tenantId, e.target.value); }}
                className="w-32 rounded border border-slate-300 bg-white px-2 py-0.5 text-lg font-bold text-slate-900 dark:border-slate-600 dark:bg-slate-800 dark:text-slate-100" />
            </div>
          </div>
          <div className="ml-auto flex items-center gap-2">
            <button onClick={() => { save('discovery-result', r); nav('/discovery'); }} className="inline-flex items-center gap-1.5 rounded-md border border-slate-300 px-3 py-2 text-sm font-medium text-slate-700 hover:bg-slate-100 dark:border-slate-600 dark:text-slate-200 dark:hover:bg-slate-800"><RefreshCw size={15} /> Re-run</button>
            <button onClick={() => exportDiscoveryWorkbook(r, analysis)} className="inline-flex items-center gap-1.5 rounded-md bg-blue-700 px-3 py-2 text-sm font-semibold text-white hover:bg-blue-800"><FileSpreadsheet size={15} /> Export Excel</button>
          </div>
        </div>
        <div className="mt-4 flex flex-wrap items-center gap-x-6 gap-y-2 border-t border-slate-200 pt-3 text-xs text-slate-500 dark:border-slate-700">
          <span className="font-mono">{r.org.tenantId}</span>
          <span className="inline-flex items-center gap-1 rounded bg-emerald-50 px-2 py-0.5 font-medium text-emerald-700 dark:bg-emerald-900/30 dark:text-emerald-300"><ShieldCheck size={12} /> Read-only — no tenant changes</span>
          <span>Assessed {new Date(r.fetchedAt).toLocaleString()}</span>
          <span className="ml-auto flex items-center gap-2">Data collected
            <span className="h-2 w-36 overflow-hidden rounded-full bg-slate-200 dark:bg-slate-700"><span className="block h-full rounded-full bg-blue-600" style={{ width: `${readiness}%` }} /></span>
            <span className="font-semibold text-slate-700 dark:text-slate-200">{readiness}%</span>
          </span>
        </div>
      </div>

      {/* Executive KPI strip */}
      <div className="mb-5 grid grid-cols-2 gap-3 sm:grid-cols-3 lg:grid-cols-6">
        {kpis.map((k) => (
          <div key={k.label} className="rounded-lg border border-slate-300 bg-white p-4 shadow-sm dark:border-slate-700 dark:bg-slate-900">
            <div className="mb-2 flex h-8 w-8 items-center justify-center rounded-md bg-slate-100 text-slate-600 dark:bg-slate-800 dark:text-slate-300"><k.icon size={16} /></div>
            <div className="truncate text-xl font-bold tabular-nums text-slate-900 dark:text-slate-100" title={String(k.value)}>{k.value}</div>
            <div className="text-xs font-medium text-slate-600 dark:text-slate-300">{k.label}</div>
            <div className="truncate text-xs text-slate-400">{k.sub}</div>
          </div>
        ))}
      </div>

      {/* Sticky section nav with active highlight */}
      <div className="sticky top-0 z-20 mb-5 flex flex-wrap gap-1 rounded-lg border border-slate-300 bg-white/95 px-2 py-1.5 shadow-sm backdrop-blur dark:border-slate-700 dark:bg-slate-900/95">
        {SECTIONS.map((s, i) => (
          <button key={s.id} type="button" onClick={() => document.getElementById(s.id)?.scrollIntoView({ behavior: 'smooth', block: 'start' })}
            className={`rounded px-3 py-1.5 text-sm font-medium transition-colors ${active === s.id ? 'bg-blue-700 text-white' : 'text-slate-600 hover:bg-slate-100 dark:text-slate-300 dark:hover:bg-slate-800'}`}>
            <span className="mr-1 opacity-60">{i + 1}.</span>{s.label}
          </button>
        ))}
      </div>

      {/* 1. Overview */}
      <Section id="overview" n={1} title="Source Tenant Overview" desc={`What was collected from ${r.org.displayName} for the migration to ${target}.`}>
        <DataTable columns={['Metric', 'Value']} rows={overviewRows} align={{ 1: 'right' }} />
      </Section>

      {/* 2. Users & Licensing */}
      <Section id="users" n={2} title="Users & Licensing" desc="Proposed target license per user, derived from how each user signs in (desktop/laptop Office app vs mobile only).">
        <div className="mb-5 grid gap-3 sm:grid-cols-3">
          {licenseSummary.map((x) => (
            <div key={x.rule} className="rounded-lg border border-slate-300 bg-slate-50 p-4 dark:border-slate-700 dark:bg-slate-800/40">
              <div className="text-2xl font-bold tabular-nums text-slate-900 dark:text-slate-100">{x.count}</div>
              <div className="text-sm font-medium text-slate-700 dark:text-slate-200">{x.rule}</div>
              <div className="mt-1"><Chip label={x.license} tone={licenseTone(x.license)} /></div>
            </div>
          ))}
        </div>
        <DataTable
          columns={['User', 'UPN', 'Worker type', 'Mailbox', 'Login evidence', 'Current licenses', 'Proposed target license', 'Action']}
          rows={r.users.filter((u) => u.userType !== 'Guest').map((u) => [
            u.displayName, u.userPrincipalName, <Chip label={u.workerType || 'Unknown'} tone={workerTone(u.workerType)} />,
            <Chip label={u.mailboxType || '—'} tone={mailboxTone(u.mailboxType)} />, u.appPlatforms || 'No evidence',
            u.licenses || '—', <Chip label={proposedLicenseOf(u.workerType || '')} tone={licenseTone(proposedLicenseOf(u.workerType || ''))} />,
            u.workerType ? 'Assign on cutover' : 'Validate manually',
          ])}
        />
        <h3 className="mb-2 mt-6 text-sm font-semibold text-slate-700 dark:text-slate-200">Mailboxes &amp; identity detail</h3>
        <DataTable
          columns={['UPN', 'Primary mail', 'Mailbox type', 'Mailbox GB', 'Aliases (SMTP)', 'Identity', 'Manager', 'Company', 'Office', 'Mobile']}
          rows={r.users.filter((u) => u.userType !== 'Guest').map((u) => [
            u.userPrincipalName, u.mail || '—', <Chip label={u.mailboxType || '—'} tone={mailboxTone(u.mailboxType)} />,
            (u.mailboxGB ?? 0) ? (u.mailboxGB ?? 0).toFixed(2) : '—', u.aliases || '—',
            <Chip label={u.hybrid || '—'} tone={/synced/i.test(u.hybrid) ? 'amber' : 'slate'} />, u.manager || '—', u.company || '—', u.office || '—', u.mobile || '—',
          ])}
          align={{ 3: 'right' }}
        />
      </Section>

      {/* 3. Endpoint Plan */}
      <Section id="endpoints" n={3} title="Endpoint Plan" desc="Suggested new workstation names per the naming convention.">
        <div className="mb-4 flex flex-wrap items-end gap-5 rounded-lg border border-slate-200 bg-slate-50 p-4 dark:border-slate-700 dark:bg-slate-800/40">
          <label className="text-sm text-slate-600 dark:text-slate-300">
            <span className="mb-1 block font-medium">Site / company code</span>
            <input value={site} onChange={(e) => { setSite(e.target.value); setSiteCode(tenantId, e.target.value); }} className="w-32 rounded border border-slate-300 bg-white px-2 py-1 text-sm dark:border-slate-600 dark:bg-slate-800 dark:text-slate-100" />
          </label>
          <div className="text-sm text-slate-500">
            <div className="font-medium text-slate-700 dark:text-slate-200">Convention</div>
            <code className="text-xs">{NAMING_FORMAT.replace('[SITE]', site)}</code>
            <div className="mt-1 text-xs">Worker: O=Office, F=Field, T=Temp, K=Kiosk · Device: L=Laptop, D=Desktop, P=Phone, T=Tablet, M=Mac</div>
          </div>
          <div className="ml-auto flex gap-2 text-center text-xs">
            <div className="rounded border border-slate-200 bg-white px-3 py-1.5 dark:border-slate-700 dark:bg-slate-900"><div className="text-base font-bold text-slate-800 dark:text-slate-100">{renameSuggestions}</div>rename plan</div>
            <div className="rounded border border-slate-200 bg-white px-3 py-1.5 dark:border-slate-700 dark:bg-slate-900"><div className="text-base font-bold text-slate-800 dark:text-slate-100">{namingErrors}</div>errors</div>
          </div>
        </div>
        <DataTable
          columns={['Current device name', 'Endpoint class', 'Operating system', 'Likely user', 'Suggested new name', 'Device type', 'Compliance', 'Validation / action']}
          rows={endpoints.map((e) => [e.device, e.cls ? <Chip label={e.cls} tone={e.cls === 'PC' ? 'blue' : 'amber'} /> : '—', e.os, e.user, e.suggested ? <span className="font-mono text-[12px]">{e.suggested}</span> : '—', e.type, <Chip label={e.compliance} tone={complianceTone(e.compliance)} />, e.action])}
        />
      </Section>

      {/* 4. Migration Inventory */}
      <Section id="inventory" n={4} title="Migration Inventory" desc="What will move to the target tenant.">
        <div className="mb-4 grid grid-cols-2 gap-3 sm:grid-cols-4">
          {[
            { l: 'Mailboxes', v: `${r.usage.mailboxCount} · ${r.usage.mailboxTotalGB.toFixed(0)} GB` },
            { l: 'OneDrive', v: `${r.usage.oneDriveCount} · ${r.usage.oneDriveTotalGB.toFixed(0)} GB` },
            { l: 'SharePoint', v: `${r.usage.spoSiteCount} · ${r.usage.spoTotalGB.toFixed(0)} GB` },
            { l: 'Teams', v: `${r.teams}` },
          ].map((k) => (
            <div key={k.l} className="rounded-lg border border-slate-300 bg-slate-50 px-4 py-3 dark:border-slate-700 dark:bg-slate-800/40">
              <div className="text-base font-bold text-slate-900 dark:text-slate-100">{k.v}</div>
              <div className="text-xs text-slate-500">{k.l}</div>
            </div>
          ))}
        </div>
        <div className="grid grid-cols-1 gap-5 lg:grid-cols-2">
          <div>
            <div className="mb-2 text-sm font-medium text-slate-700 dark:text-slate-200">Groups ({r.groups.length})</div>
            <DataTable columns={['Group', 'Type', 'Is Team']} rows={r.groups.map((g) => [g.displayName, g.groupType, g.isTeam ? 'Yes' : 'No'])} max={300} />
          </div>
          <div>
            <div className="mb-2 text-sm font-medium text-slate-700 dark:text-slate-200">SharePoint sites ({r.sharePointSites.length})</div>
            <DataTable columns={['Site', 'URL']} rows={r.sharePointSites.map((s) => [s.name, s.webUrl])} max={300} />
          </div>
        </div>
      </Section>

      {/* 5. Test Migration */}
      <Section id="testmig" n={5} title="Test Migration (Pilot Batch)" desc={`A representative pilot to validate the ${r.org.displayName} → ${target} migration before full cutover. Preview only — no changes are made.`}>
        <div className="mb-4 flex flex-wrap items-end gap-5 rounded-lg border border-slate-200 bg-slate-50 p-4 dark:border-slate-700 dark:bg-slate-800/40">
          <label className="text-sm text-slate-600 dark:text-slate-300">
            <span className="mb-1 block font-medium">Target domain</span>
            <input value={targetDomain} onChange={(e) => { setTargetDomainState(e.target.value); setTargetDomain(tenantId, e.target.value); }} className="w-56 rounded border border-slate-300 bg-white px-2 py-1 text-sm dark:border-slate-600 dark:bg-slate-800 dark:text-slate-100" />
          </label>
          <label className="text-sm text-slate-600 dark:text-slate-300">
            <span className="mb-1 block font-medium">Pilot size</span>
            <input type="number" min={1} max={50} value={pilotSize} onChange={(e) => setPilotSize(Math.max(1, Math.min(50, Number(e.target.value) || 1)))} className="w-20 rounded border border-slate-300 bg-white px-2 py-1 text-sm dark:border-slate-600 dark:bg-slate-800 dark:text-slate-100" />
          </label>
        </div>
        <DataTable
          columns={['Source UPN', 'Target UPN', 'Display name', 'Worker type', 'Proposed license', 'Mailbox GB', 'Aliases (SMTP)', 'Suggested device', 'Action']}
          rows={pilot.map((u) => [
            u.userPrincipalName, <span className="font-medium text-blue-700 dark:text-blue-400">{mapUpnToTarget(u.userPrincipalName, targetDomain) || '—'}</span>, u.displayName,
            <Chip label={u.workerType || 'Validate'} tone={workerTone(u.workerType)} />, <Chip label={proposedLicenseOf(u.workerType || '')} tone={licenseTone(proposedLicenseOf(u.workerType || ''))} />,
            (u.mailboxGB ?? 0) ? (u.mailboxGB ?? 0).toFixed(2) : '—', u.aliases || '—',
            firstDeviceOf(u.userPrincipalName) ? <span className="font-mono text-[12px]">{firstDeviceOf(u.userPrincipalName)}</span> : '—', u.workerType ? 'Pilot wave 1' : 'Validate then pilot',
          ])}
          align={{ 5: 'right' }}
        />
      </Section>

      {/* 6. Risks & Issues */}
      <Section id="risks" n={6} title="Risks & Issues" desc="Migration risks, blockers and privileged accounts.">
        {analysis && (
          <ul className="mb-5 space-y-1.5">
            {analysis.findings.filter((f) => f.level !== 'info').map((f, i) => (
              <li key={i} className="flex items-start gap-2 text-sm">
                <span className="w-16 shrink-0"><Chip label={f.level} tone={findingChip[f.level]} /></span>
                <span className="text-slate-700 dark:text-slate-200"><strong className={findingTone[f.level]}>{f.area}:</strong> {f.text}</span>
              </li>
            ))}
          </ul>
        )}
        <h3 className="mb-2 text-sm font-semibold text-slate-700 dark:text-slate-200">Privileged accounts ({(r.adminRoles ?? []).reduce((a, x) => a + x.members.length, 0)})</h3>
        <DataTable columns={['Admin role', 'Members']} rows={(r.adminRoles ?? []).map((x) => [<Chip label={x.role} tone={/global/i.test(x.role) ? 'red' : 'amber'} />, x.members.join(', ')])} />
        <h3 className="mb-2 mt-6 text-sm font-semibold text-slate-700 dark:text-slate-200">Validation items ({r.validations.length})</h3>
        <DataTable columns={['Data set', 'Status', 'Note']} rows={r.validations.map((v) => [v.workload, v.status || '—', v.note])} />
      </Section>

      {/* 7. Evidence */}
      <Section id="evidence" n={7} title="Data Sources / Evidence" desc="Every figure is collected read-only from Microsoft Graph — this is the provenance for the analysis.">
        <DataTable columns={['Data set', 'Microsoft Graph source', 'Status', 'Records']} rows={evidence} align={{ 3: 'right' }} />
      </Section>
    </div>
  );
}
