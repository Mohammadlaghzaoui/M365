import { useMemo, useState } from 'react';
import { useParams, useNavigate, Link } from 'react-router-dom';
import { ArrowLeft, FileSpreadsheet, RefreshCw } from 'lucide-react';
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

type Cell = string | number;

function DataTable({ columns, rows, max = 1000, align }: { columns: string[]; rows: Cell[][]; max?: number; align?: Record<number, 'right'> }) {
  if (!rows.length) return <p className="px-1 py-6 text-center text-sm text-slate-400">No rows.</p>;
  const shown = rows.slice(0, max);
  return (
    <div className="overflow-auto rounded border border-slate-300 dark:border-slate-700" style={{ maxHeight: 520 }}>
      <table className="w-full border-collapse text-[13px]">
        <thead className="sticky top-0 z-10 bg-slate-100 dark:bg-slate-800">
          <tr>{columns.map((c, i) => <th key={c} className={`whitespace-nowrap border-b border-slate-300 dark:border-slate-700 px-3 py-2 font-semibold text-slate-700 dark:text-slate-200 ${align?.[i] === 'right' ? 'text-right' : 'text-left'}`}>{c}</th>)}</tr>
        </thead>
        <tbody>
          {shown.map((r, i) => (
            <tr key={i} className="odd:bg-white even:bg-slate-50 dark:odd:bg-slate-900 dark:even:bg-slate-800/40">
              {r.map((cell, j) => <td key={j} className={`whitespace-nowrap border-b border-slate-100 dark:border-slate-800 px-3 py-1.5 text-slate-700 dark:text-slate-200 ${align?.[j] === 'right' ? 'text-right tabular-nums' : ''}`}>{String(cell)}</td>)}
            </tr>
          ))}
        </tbody>
      </table>
      {rows.length > max && <div className="border-t border-slate-200 bg-slate-50 px-3 py-2 text-center text-xs text-slate-400 dark:border-slate-700 dark:bg-slate-800/50">Showing {max} of {rows.length} rows — export to Excel for the full set.</div>}
    </div>
  );
}

function Section({ id, title, children }: { id: string; title: string; children: React.ReactNode }) {
  return (
    <section id={id} className="mb-6 scroll-mt-28 rounded border border-slate-300 bg-white p-5 dark:border-slate-700 dark:bg-slate-900">
      <h2 className="mb-4 border-b border-slate-200 pb-2 text-base font-semibold text-slate-800 dark:border-slate-700 dark:text-slate-100">{title}</h2>
      {children}
    </section>
  );
}

const SECTIONS = [
  { id: 'overview', label: '1. Overview' },
  { id: 'users', label: '2. Users & Licensing' },
  { id: 'endpoints', label: '3. Endpoint Plan' },
  { id: 'inventory', label: '4. Migration Inventory' },
  { id: 'testmig', label: '5. Test Migration' },
  { id: 'risks', label: '6. Risks & Issues' },
  { id: 'evidence', label: '7. Data Sources / Evidence' },
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

  if (!result) {
    return (
      <div className="p-8 text-center">
        <p className="mb-3 text-slate-500">This source tenant's assessment was not found (it may have been wiped).</p>
        <Link to="/tenants" className="inline-flex items-center gap-1 text-sm text-blue-700 hover:underline"><ArrowLeft size={15} /> Back to Source Tenants</Link>
      </div>
    );
  }
  const r = result;
  const workerByUser = useMemo(() => new Map(r.users.map((u) => [u.userPrincipalName.toLowerCase(), u.workerType] as const)), [r]);

  // ---- Endpoint plan rows ----
  const endpoints = (r.deviceInventory ?? []).map((d) => {
    const wt = workerByUser.get((d.user || '').toLowerCase()) ?? '';
    const cls = endpointClassOf(d.formFactor);
    const sug = suggestedName(site, d, wt);
    return {
      device: d.deviceName, cls, os: `${d.os}${d.osVersion ? ' ' + d.osVersion : ''}`.trim(),
      user: d.user || '—', suggested: sug, type: deviceTypeLabel(d.formFactor),
      action: sug ? 'Rename to suggested' : 'Validate manually',
    };
  });
  const renameSuggestions = endpoints.filter((e) => e.suggested).length;
  const namingErrors = (() => { const seen = new Set<string>(); let dup = 0; for (const e of endpoints) { if (!e.suggested) continue; if (seen.has(e.suggested)) dup++; else seen.add(e.suggested); } return dup; })();

  // ---- License proposal ----
  const licCount = (m: string) => r.users.filter((u) => u.accountEnabled && u.userType !== 'Guest' && (u.workerType || '') === m).length;
  const licenseSummary = LICENSE_RULES.map((x) => ({ ...x, count: licCount(x.match) }));

  // ---- Test migration pilot: a representative sample with source -> target UPN mapping ----
  const firstDeviceOf = (upn: string) => {
    const d = (r.deviceInventory ?? []).find((x) => (x.user || '').toLowerCase() === upn.toLowerCase());
    return d ? suggestedName(site, d, workerByUser.get(upn.toLowerCase()) ?? '') : '';
  };
  const candidates = r.users.filter((u) => u.accountEnabled && u.userType !== 'Guest');
  const pilot = (() => {
    const office = candidates.filter((u) => u.workerType === 'Office');
    const field = candidates.filter((u) => u.workerType === 'Field');
    const other = candidates.filter((u) => !u.workerType);
    const mix = [...office.slice(0, Math.ceil(pilotSize / 2)), ...field.slice(0, 1), ...other.slice(0, pilotSize)];
    const seen = new Set<string>();
    return mix.filter((u) => (seen.has(u.userPrincipalName) ? false : (seen.add(u.userPrincipalName), true))).slice(0, pilotSize);
  })();

  const overviewRows: Cell[][] = [
    ['Source tenant', r.org.displayName],
    ['Migration target', target],
    ['Tenant ID', r.org.tenantId],
    ['Users', r.users.length],
    ['Mailboxes', r.usage.mailboxCount],
    ['Mailbox data (GB)', r.usage.mailboxTotalGB.toFixed(2)],
    ['Groups', r.groups.length],
    ['Teams', r.teams],
    ['SharePoint sites', r.sharePointSites.length],
    ['Devices', r.devices.total],
    ['Rename suggestions', renameSuggestions],
    ['Naming errors', namingErrors],
    ['Assessed', new Date(r.fetchedAt).toLocaleString()],
  ];

  const findingTone: Record<string, string> = { blocker: 'text-rose-700 dark:text-rose-400', risk: 'text-orange-700 dark:text-orange-400', warn: 'text-amber-700 dark:text-amber-400', info: 'text-slate-600 dark:text-slate-300' };

  // ---- Evidence: which Graph source produced each data set ----
  const evidence: Cell[][] = [
    ['Organisation / domains', 'GET /organization, /domains', 'Collected', r.domains.length],
    ['Users', 'GET /users', 'Collected', r.users.length],
    ['Licenses', 'GET /subscribedSkus', 'Collected', r.licenses.length],
    ['Groups & Teams', 'GET /groups', 'Collected', r.groups.length],
    ['Devices (Intune)', 'GET /deviceManagement/managedDevices', r.deviceInventory.some((d) => d.source === 'Intune') ? 'Collected' : 'None', r.deviceInventory.filter((d) => d.source === 'Intune').length],
    ['Devices (Entra)', 'GET /devices', r.deviceInventory.some((d) => d.source === 'Entra') ? 'Collected' : 'None', r.deviceInventory.filter((d) => d.source === 'Entra').length],
    ['Office app usage', "GET /reports/getM365AppUserDetail", r.appUsage?.classified ? 'Collected' : 'Not available', r.appUsage?.classified ?? 0],
    ['Mailbox sizing', "GET /reports/getMailboxUsageDetail", r.usage.available ? 'Collected' : 'Not available', r.usage.mailboxCount],
    ['SharePoint sites', 'GET /sites', 'Collected', r.sharePointSites.length],
    ['Conditional Access', 'GET /identity/conditionalAccess/policies', 'Collected', r.caPolicies.length],
    ['Admin roles', 'GET /directoryRoles?$expand=members', (r.adminRoles ?? []).length ? 'Collected' : 'None', (r.adminRoles ?? []).reduce((a, x) => a + x.members.length, 0)],
  ];

  return (
    <div className="mx-auto max-w-[1200px]">
      {/* Header */}
      <div className="mb-5">
        <Link to="/tenants" className="mb-2 inline-flex items-center gap-1 text-sm text-slate-500 hover:text-blue-700"><ArrowLeft size={14} /> Source Tenants</Link>
        <div className="flex flex-wrap items-center gap-3 border-b border-slate-300 pb-4 dark:border-slate-700">
          <div className="flex h-11 w-11 shrink-0 items-center justify-center rounded bg-slate-800 text-sm font-bold text-white">
            {r.org.displayName.slice(0, 2).toUpperCase()}
          </div>
          <div className="mr-auto">
            <h1 className="text-xl font-bold text-slate-900 dark:text-slate-100">Source Tenant — {r.org.displayName}</h1>
            <div className="mt-0.5 flex items-center gap-2 text-sm text-slate-500">
              <span>Migration target:</span>
              <input value={target} onChange={(e) => { setTarget(e.target.value); setMigrationTarget(tenantId, e.target.value); }}
                className="w-28 rounded border border-slate-300 bg-white px-2 py-0.5 text-sm font-medium text-slate-800 dark:border-slate-600 dark:bg-slate-800 dark:text-slate-100" />
              <span className="rounded bg-emerald-50 px-2 py-0.5 text-xs font-medium text-emerald-700 dark:bg-emerald-900/30 dark:text-emerald-300">Read-only — no tenant changes</span>
            </div>
          </div>
          <button onClick={() => { save('discovery-result', r); nav('/discovery'); }} className="inline-flex items-center gap-1.5 rounded border border-slate-300 px-3 py-2 text-sm font-medium text-slate-700 hover:bg-slate-100 dark:border-slate-600 dark:text-slate-200 dark:hover:bg-slate-800"><RefreshCw size={15} /> Re-run</button>
          <button onClick={() => exportDiscoveryWorkbook(r, analysis)} className="inline-flex items-center gap-1.5 rounded bg-blue-700 px-3 py-2 text-sm font-semibold text-white hover:bg-blue-800"><FileSpreadsheet size={15} /> Export Excel</button>
        </div>
      </div>

      {/* Section nav — scrolls in-page (no hash navigation; this app uses HashRouter) */}
      <div className="sticky top-0 z-20 mb-5 flex flex-wrap gap-1 border-b border-slate-300 bg-slate-50 py-1 dark:border-slate-700 dark:bg-slate-900">
        {SECTIONS.map((s) => (
          <button key={s.id} type="button"
            onClick={() => document.getElementById(s.id)?.scrollIntoView({ behavior: 'smooth', block: 'start' })}
            className="rounded px-3 py-1.5 text-sm font-medium text-slate-600 hover:bg-slate-200 hover:text-slate-900 dark:text-slate-300 dark:hover:bg-slate-800">{s.label}</button>
        ))}
      </div>

      {/* 1. Overview */}
      <Section id="overview" title="1. Source Tenant Overview">
        <div className="grid grid-cols-1 gap-5 lg:grid-cols-2">
          <DataTable columns={['Metric', 'Value']} rows={overviewRows} align={{ 1: 'right' }} />
          <div className="grid grid-cols-2 gap-3 self-start sm:grid-cols-3">
            {[
              { l: 'Users', v: r.users.length }, { l: 'Mailboxes', v: r.usage.mailboxCount }, { l: 'Mailbox GB', v: r.usage.mailboxTotalGB.toFixed(0) },
              { l: 'Groups', v: r.groups.length }, { l: 'Teams', v: r.teams }, { l: 'SharePoint', v: r.sharePointSites.length },
              { l: 'Devices', v: r.devices.total }, { l: 'Rename plan', v: renameSuggestions }, { l: 'Naming errors', v: namingErrors },
            ].map((k) => (
              <div key={k.l} className="rounded border border-slate-200 bg-slate-50 px-3 py-3 dark:border-slate-700 dark:bg-slate-800/40">
                <div className="text-lg font-bold tabular-nums text-slate-800 dark:text-slate-100">{k.v}</div>
                <div className="text-xs text-slate-500">{k.l}</div>
              </div>
            ))}
          </div>
        </div>
      </Section>

      {/* 2. Users & Licensing */}
      <Section id="users" title="2. Users & Licensing">
        <p className="mb-3 text-sm text-slate-500">Proposed target license is derived from how each user signs in (desktop/laptop Office app vs mobile only). Users with no clear login evidence are flagged for manual validation with {r.org.displayName}.</p>
        <div className="mb-5 overflow-hidden rounded border border-slate-300 dark:border-slate-700">
          <table className="w-full text-sm">
            <thead className="bg-slate-100 dark:bg-slate-800"><tr>
              <th className="px-3 py-2 text-left font-semibold text-slate-700 dark:text-slate-200">User group</th>
              <th className="px-3 py-2 text-left font-semibold text-slate-700 dark:text-slate-200">Proposed target license</th>
              <th className="px-3 py-2 text-right font-semibold text-slate-700 dark:text-slate-200">Users</th>
            </tr></thead>
            <tbody>
              {licenseSummary.map((x) => (
                <tr key={x.rule} className="border-t border-slate-100 dark:border-slate-800">
                  <td className="px-3 py-2 text-slate-700 dark:text-slate-200">{x.rule}</td>
                  <td className="px-3 py-2 text-slate-700 dark:text-slate-200">{x.license}</td>
                  <td className="px-3 py-2 text-right font-semibold tabular-nums text-slate-800 dark:text-slate-100">{x.count}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <DataTable
          columns={['User', 'UPN', 'Worker type', 'Mailbox', 'Login evidence', 'Current licenses', 'Proposed target license', 'Action']}
          rows={r.users.filter((u) => u.userType !== 'Guest').map((u) => [
            u.displayName, u.userPrincipalName, u.workerType || '—', u.mailboxType || '—', u.appPlatforms || 'No evidence',
            u.licenses || '—', proposedLicenseOf(u.workerType || ''),
            (u.workerType ? 'Assign on cutover' : 'Validate manually'),
          ])}
        />

        <h3 className="mb-2 mt-6 text-sm font-semibold text-slate-700 dark:text-slate-200">Mailboxes &amp; identity detail</h3>
        <DataTable
          columns={['UPN', 'Primary mail', 'Mailbox type', 'Mailbox GB', 'Aliases (SMTP)', 'Identity', 'Manager', 'Company', 'Office', 'Mobile']}
          rows={r.users.filter((u) => u.userType !== 'Guest').map((u) => [
            u.userPrincipalName, u.mail || '—', u.mailboxType || '—', (u.mailboxGB ?? 0) ? (u.mailboxGB ?? 0).toFixed(2) : '—',
            u.aliases || '—', u.hybrid || '—', u.manager || '—', u.company || '—', u.office || '—', u.mobile || '—',
          ])}
          align={{ 3: 'right' }}
        />
      </Section>

      {/* 3. Endpoint Plan */}
      <Section id="endpoints" title="3. Endpoint Plan">
        <div className="mb-3 flex flex-wrap items-end gap-4">
          <label className="text-sm text-slate-600 dark:text-slate-300">
            <span className="mb-1 block font-medium">Site / company code</span>
            <input value={site} onChange={(e) => { setSite(e.target.value); setSiteCode(tenantId, e.target.value); }}
              className="w-32 rounded border border-slate-300 bg-white px-2 py-1 text-sm dark:border-slate-600 dark:bg-slate-800 dark:text-slate-100" />
          </label>
          <div className="text-sm text-slate-500">
            <div className="font-medium text-slate-700 dark:text-slate-200">Naming convention</div>
            <code className="text-xs">{NAMING_FORMAT.replace('[SITE]', site)}</code>
            <div className="mt-1 text-xs">Worker: O=Office, F=Field, T=Temp, K=Kiosk · Device: L=Laptop, D=Desktop, P=Phone, T=Tablet, M=Mac</div>
          </div>
        </div>
        <DataTable
          columns={['Current device name', 'Endpoint class', 'Operating system', 'Likely user', 'Suggested new name', 'Device type', 'Naming convention', 'Validation / action']}
          rows={endpoints.map((e) => [e.device, e.cls || '—', e.os, e.user, e.suggested || '—', e.type, e.suggested ? `${site}-[Worker][Device][Last5]` : '—', e.action])}
        />
        <p className="mt-2 text-xs text-slate-400">{renameSuggestions} rename suggestions · {namingErrors} naming errors (duplicate suggested names).</p>
      </Section>

      {/* 4. Migration Inventory */}
      <Section id="inventory" title="4. Migration Inventory">
        <div className="mb-4 grid grid-cols-2 gap-3 sm:grid-cols-4">
          {[
            { l: 'Mailboxes', v: `${r.usage.mailboxCount} · ${r.usage.mailboxTotalGB.toFixed(0)} GB` },
            { l: 'OneDrive', v: `${r.usage.oneDriveCount} · ${r.usage.oneDriveTotalGB.toFixed(0)} GB` },
            { l: 'SharePoint', v: `${r.usage.spoSiteCount} · ${r.usage.spoTotalGB.toFixed(0)} GB` },
            { l: 'Teams', v: `${r.teams}` },
          ].map((k) => (
            <div key={k.l} className="rounded border border-slate-200 bg-slate-50 px-3 py-3 dark:border-slate-700 dark:bg-slate-800/40">
              <div className="text-base font-bold text-slate-800 dark:text-slate-100">{k.v}</div>
              <div className="text-xs text-slate-500">{k.l}</div>
            </div>
          ))}
        </div>
        <div className="grid grid-cols-1 gap-5 lg:grid-cols-2">
          <div>
            <div className="mb-1 text-sm font-medium text-slate-700 dark:text-slate-200">Groups ({r.groups.length})</div>
            <DataTable columns={['Group', 'Type', 'Is Team']} rows={r.groups.map((g) => [g.displayName, g.groupType, g.isTeam ? 'Yes' : 'No'])} max={300} />
          </div>
          <div>
            <div className="mb-1 text-sm font-medium text-slate-700 dark:text-slate-200">SharePoint sites ({r.sharePointSites.length})</div>
            <DataTable columns={['Site', 'URL']} rows={r.sharePointSites.map((s) => [s.name, s.webUrl])} max={300} />
          </div>
        </div>
      </Section>

      {/* 5. Test Migration (pilot) */}
      <Section id="testmig" title="5. Test Migration (Pilot Batch)">
        <p className="mb-3 text-sm text-slate-500">A representative pilot batch to validate the {r.org.displayName} → {target} migration before the full cutover. Source UPNs are mapped to the target domain; verify the mapping and licenses with both tenants.</p>
        <div className="mb-3 flex flex-wrap items-end gap-4">
          <label className="text-sm text-slate-600 dark:text-slate-300">
            <span className="mb-1 block font-medium">Target domain</span>
            <input value={targetDomain} onChange={(e) => { setTargetDomainState(e.target.value); setTargetDomain(tenantId, e.target.value); }}
              className="w-56 rounded border border-slate-300 bg-white px-2 py-1 text-sm dark:border-slate-600 dark:bg-slate-800 dark:text-slate-100" />
          </label>
          <label className="text-sm text-slate-600 dark:text-slate-300">
            <span className="mb-1 block font-medium">Pilot size</span>
            <input type="number" min={1} max={50} value={pilotSize} onChange={(e) => setPilotSize(Math.max(1, Math.min(50, Number(e.target.value) || 1)))}
              className="w-20 rounded border border-slate-300 bg-white px-2 py-1 text-sm dark:border-slate-600 dark:bg-slate-800 dark:text-slate-100" />
          </label>
        </div>
        <DataTable
          columns={['Source UPN', 'Target UPN', 'Display name', 'Worker type', 'Proposed license', 'Mailbox GB', 'Aliases (SMTP)', 'Suggested device name', 'Action']}
          rows={pilot.map((u) => [
            u.userPrincipalName, mapUpnToTarget(u.userPrincipalName, targetDomain) || '—', u.displayName,
            u.workerType || 'Validate', proposedLicenseOf(u.workerType || ''),
            (u.mailboxGB ?? 0) ? (u.mailboxGB ?? 0).toFixed(2) : '—', u.aliases || '—',
            firstDeviceOf(u.userPrincipalName) || '—', u.workerType ? 'Pilot wave 1' : 'Validate then pilot',
          ])}
          align={{ 5: 'right' }}
        />
        <p className="mt-2 text-xs text-slate-400">Pilot is a preview only — no changes are made to either tenant. Use it to confirm UPN mapping, licensing and device naming before scheduling waves.</p>
      </Section>

      {/* 6. Risks & Issues */}
      <Section id="risks" title="6. Risks & Issues">
        {analysis && (
          <ul className="mb-4 space-y-1.5">
            {analysis.findings.filter((f) => f.level !== 'info').map((f, i) => (
              <li key={i} className="flex gap-2 text-sm">
                <span className={`w-16 shrink-0 font-semibold uppercase ${findingTone[f.level]}`}>{f.level}</span>
                <span className="text-slate-700 dark:text-slate-200"><strong>{f.area}:</strong> {f.text}</span>
              </li>
            ))}
          </ul>
        )}
        <h3 className="mb-2 text-sm font-semibold text-slate-700 dark:text-slate-200">Privileged accounts ({(r.adminRoles ?? []).reduce((a, x) => a + x.members.length, 0)})</h3>
        <DataTable columns={['Admin role', 'Members']} rows={(r.adminRoles ?? []).map((x) => [x.role, x.members.join(', ')])} />
        <div className="mb-1 mt-6 text-sm font-medium text-slate-700 dark:text-slate-200">Validation items ({r.validations.length})</div>
        <DataTable columns={['Data set', 'Status', 'Note']} rows={r.validations.map((v) => [v.workload, v.status || '—', v.note])} />
      </Section>

      {/* 7. Data Sources / Evidence */}
      <Section id="evidence" title="7. Data Sources / Evidence">
        <p className="mb-3 text-sm text-slate-500">Every figure above is collected read-only from Microsoft Graph. This is the provenance for the migration analysis.</p>
        <DataTable columns={['Data set', 'Microsoft Graph source', 'Status', 'Records']} rows={evidence} align={{ 3: 'right' }} />
      </Section>
    </div>
  );
}
