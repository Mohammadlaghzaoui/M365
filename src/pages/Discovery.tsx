import { useEffect, useMemo, useRef, useState } from 'react';
import { Radar, ShieldCheck, Loader2, FileSpreadsheet, RefreshCw, Building2, Users2, Boxes, Globe, Sparkles, Gauge, ListChecks, HardDrive } from 'lucide-react';
import { Badge, Button, Card, PageHeader, Section } from '../components/ui';
import { DonutChart, HBarChart } from '../components/charts';
import { runDiscovery, DiscoveryResult, DISCOVERY_SCOPES, LogLevel } from '../services/graphDiscovery';
import { assess } from '../services/migrationAssessment';
import { downloadWorkbook, Sheet } from '../services/excelExport';
import { currentAccount } from '../services/sso';
import { getSSOSettings, aiEnabled } from '../store/settings';
import { load, save } from '../store/useLocalStorage';
import { chat } from '../services/ai';
import { Link } from 'react-router-dom';

export default function Discovery() {
  const [result, setResult] = useState<DiscoveryResult | null>(() => load<DiscoveryResult | null>('discovery-result', null));
  const [busy, setBusy] = useState(false);
  const [lines, setLines] = useState<{ text: string; level: LogLevel }[]>([]);
  const [error, setError] = useState('');
  const logRef = useRef<HTMLDivElement>(null);
  useEffect(() => { logRef.current?.scrollTo({ top: logRef.current.scrollHeight }); }, [lines]);
  const addLine = (text: string, level: LogLevel = 'info') =>
    setLines((prev) => [...prev, { text: level === 'cmd' ? text : `[${new Date().toLocaleTimeString('en-GB')}] ${text}`, level }]);
  const [ai, setAi] = useState('');
  const [aiBusy, setAiBusy] = useState(false);
  const ssoReady = getSSOSettings().enabled && !!getSSOSettings().clientId;
  const analysis = useMemo(() => (result ? assess(result) : null), [result]);

  const run = async () => {
    setBusy(true); setError(''); setAi(''); setLines([]);
    try {
      if (!(await currentAccount())) {
        setError('Sign in with Microsoft 365 first (Settings → Sign-in). Discovery uses your delegated, read-only Graph access.');
        return;
      }
      const r = await runDiscovery(addLine);
      addLine(`Discovery complete — ${r.users.length} users, ${r.groups.length} groups, ~${Math.round((r.usage.mailboxTotalGB + r.usage.oneDriveTotalGB + r.usage.spoTotalGB))} GB data. READ-ONLY: nothing was written.`, 'ok');
      setResult(r);
      save('discovery-result', r);
    } catch (e) {
      addLine(`ERROR: ${e instanceof Error ? e.message : e}`, 'err');
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setBusy(false);
    }
  };

  const exportExcel = () => {
    if (!result) return;
    const sheets: Sheet[] = [
      {
        name: 'Summary',
        columns: ['Metric', 'Value'],
        rows: [
          ['Tenant', result.org.displayName],
          ['Tenant ID', result.org.tenantId],
          ['Verified domains', result.org.verifiedDomains],
          ['Total users', result.users.length],
          ['Guests', result.guests],
          ['Disabled accounts', result.disabled],
          ['Microsoft 365 groups', result.m365Groups],
          ['Teams', result.teams],
          ['Security groups', result.securityGroups],
          ['Distribution groups', result.distributionGroups],
          ['Managed devices', result.devices.total],
          ['Generated', new Date(result.fetchedAt).toLocaleString()],
        ],
      },
      {
        name: 'Users',
        columns: ['Display name', 'UPN', 'Mail', 'Type', 'Enabled', 'Department', 'Job title', 'Usage location', 'Licenses', 'Created', 'Last sign-in'],
        rows: result.users.map((u) => [u.displayName, u.userPrincipalName, u.mail, u.userType, u.accountEnabled ? 'Yes' : 'No', u.department, u.jobTitle, u.usageLocation, u.licenses, u.createdDateTime, u.lastSignIn]),
      },
      {
        name: 'Groups',
        columns: ['Display name', 'Mail', 'Type', 'Membership', 'Visibility', 'Is Team'],
        rows: result.groups.map((g) => [g.displayName, g.mail, g.groupType, g.membershipType, g.visibility, g.isTeam ? 'Yes' : 'No']),
      },
      {
        name: 'Licenses',
        columns: ['SKU', 'Enabled', 'Consumed', 'Available'],
        rows: result.licenses.map((l) => [l.skuPartNumber, l.enabled, l.consumed, l.available]),
      },
      {
        name: 'Domains',
        columns: ['Domain', 'Default', 'Verified', 'Services'],
        rows: result.domains.map((d) => [d.id, d.isDefault ? 'Yes' : 'No', d.isVerified ? 'Yes' : 'No', d.supportedServices]),
      },
      {
        name: 'Devices',
        columns: ['Operating system', 'Count'],
        rows: Object.entries(result.devices.byOs).map(([os, n]) => [os, n]),
      },
      {
        name: 'Data sizing',
        columns: ['Workload', 'Total GB', 'Count', 'Notes'],
        rows: [
          ['Mailboxes', result.usage.mailboxTotalGB, result.usage.mailboxCount, `${result.usage.mailboxOver50GB} over 50GB, ${result.usage.archiveCount} archives`],
          ['OneDrive', result.usage.oneDriveTotalGB, result.usage.oneDriveCount, `${result.usage.oneDriveOver100GB} over 100GB`],
          ['SharePoint', result.usage.spoTotalGB, result.usage.spoSiteCount, ''],
        ],
      },
      {
        name: 'Largest mailboxes',
        columns: ['User', 'Size (GB)'],
        rows: result.usage.mailboxLargest.map((m) => [m.upn, m.gb]),
      },
      ...(analysis ? [{
        name: 'Migration analysis',
        columns: ['Category', 'Detail'],
        rows: [
          ['Complexity', `${analysis.complexityLabel} (${analysis.complexityScore}/100)`],
          ['Total data', `~${analysis.totalDataGB} GB`],
          ['Estimated window', `~${analysis.estimatedDays} working days`],
          ['Recommended approach', analysis.recommendedApproach],
          ...analysis.findings.map((f) => [f.level.toUpperCase() + ' · ' + f.area, f.text] as (string | number)[]),
        ],
      }, {
        name: 'Prep checklist',
        columns: ['Phase', 'Task'],
        rows: analysis.checklist.map((c) => [c.phase, c.task]),
      }] : []),
    ];
    downloadWorkbook(sheets, `migration-discovery-${result.org.displayName.replace(/[^a-z0-9]+/gi, '-')}-${new Date().toISOString().slice(0, 10)}`);
  };

  const analyze = async () => {
    if (!result) return;
    setAiBusy(true);
    try {
      const summary = `Tenant ${result.org.displayName}: ${result.users.length} users (${result.guests} guests, ${result.disabled} disabled), ${result.m365Groups} M365 groups, ${result.teams} Teams, ${result.securityGroups} security groups, ${result.distributionGroups} distribution lists, ${result.devices.total} managed devices, ${result.domains.length} domains. Licenses: ${result.licenses.map((l) => `${l.skuPartNumber} ${l.consumed}/${l.enabled}`).join('; ')}.`;
      setAi(await chat([{ role: 'user', content: `You are a Microsoft 365 migration architect. Based on this read-only tenant discovery, produce a concise migration analysis: complexity, key risks, licensing observations, what to plan for (mailboxes, groups, Teams, devices), and recommended migration approach.\n\n${summary}` }]));
    } catch (e) {
      setAi(`AI error: ${e instanceof Error ? e.message : e}`);
    } finally {
      setAiBusy(false);
    }
  };

  const licenseChart = result?.licenses.slice(0, 8).map((l) => ({ label: l.skuPartNumber, value: l.consumed })) ?? [];
  const deviceChart = result ? Object.entries(result.devices.byOs).map(([label, value]) => ({ label, value })) : [];

  return (
    <div>
      <PageHeader title="Migration Discovery & Analysis" subtitle="Pull your whole tenant via Microsoft Graph — strictly read-only — and turn it into a migration analysis and an Excel discovery workbook." icon={<Radar size={20} />} />

      <Card className="mb-5 p-4 border-emerald-200 dark:border-emerald-800">
        <div className="flex items-start gap-2 text-sm text-emerald-700 dark:text-emerald-300">
          <ShieldCheck size={16} className="mt-0.5 shrink-0" />
          <span><strong>Read-only by design.</strong> Discovery requests only <code>*.Read.All</code> scopes and issues only GET calls — nothing is ever written to your tenant. Scopes: {DISCOVERY_SCOPES.join(', ')}.</span>
        </div>
      </Card>

      <Card className="mb-5 p-5">
        <div className="flex flex-wrap items-center gap-3">
          <Button onClick={run} disabled={busy}>
            {busy ? <Loader2 size={16} className="animate-spin" /> : <RefreshCw size={16} />} {result ? 'Re-run discovery' : 'Run discovery (read-only)'}
          </Button>
          {result && <Button variant="secondary" onClick={exportExcel}><FileSpreadsheet size={15} /> Export Excel workbook</Button>}
          {result && aiEnabled() && <Button variant="ai" onClick={analyze} disabled={aiBusy}>{aiBusy ? <Loader2 size={15} className="animate-spin" /> : <Sparkles size={15} />} AI migration analysis</Button>}
          {busy && <span className="text-sm text-slate-500">Querying Microsoft Graph…</span>}
        </div>
        {!ssoReady && <p className="mt-3 text-sm text-amber-600 dark:text-amber-400">Microsoft 365 SSO is not configured yet — set it up in <Link to="/settings" className="font-semibold hover:underline">Settings → Sign-in</Link>. The SSO app needs admin consent for the read scopes above.</p>}
        {error && <p className="mt-3 text-sm text-red-500">{error}</p>}
      </Card>

      {/* Live Microsoft Graph console */}
      {lines.length > 0 && (
        <Card className="mb-5 overflow-hidden p-0">
          <div className="flex items-center gap-2 border-b border-slate-700 bg-slate-900 px-3 py-2">
            <span className="h-3 w-3 rounded-full bg-red-500" /><span className="h-3 w-3 rounded-full bg-amber-400" /><span className="h-3 w-3 rounded-full bg-emerald-500" />
            <span className="ml-2 text-xs font-medium text-slate-300">Microsoft Graph — live read-only discovery {busy && <span className="animate-pulse">· running</span>}</span>
          </div>
          <div ref={logRef} className="h-72 overflow-y-auto bg-[#012456] p-3 font-mono text-xs leading-relaxed">
            {lines.map((l, i) => (
              <div key={i} className={`whitespace-pre-wrap break-all ${{ cmd: 'text-yellow-300', info: 'text-cyan-300', ok: 'text-emerald-400', warn: 'text-amber-300', err: 'text-red-400' }[l.level]}`}>
                {l.level === 'cmd' ? <><span className="text-white">PS C:\WorkPilot&gt; </span>{l.text}</> : l.text}
              </div>
            ))}
            {busy && <span className="inline-block h-3.5 w-2 animate-pulse bg-slate-200 align-middle" />}
          </div>
        </Card>
      )}

      {result && (
        <div className="space-y-5">
          {/* KPI strip */}
          <div className="grid grid-cols-2 gap-4 lg:grid-cols-4">
            {[
              { icon: Building2, label: 'Tenant', value: result.org.displayName, sub: `${result.org.verifiedDomains} verified domains`, color: 'text-blue-600 dark:text-blue-400' },
              { icon: Users2, label: 'Users', value: result.users.length, sub: `${result.guests} guests · ${result.disabled} disabled`, color: 'text-violet-600 dark:text-violet-400' },
              { icon: Boxes, label: 'Groups / Teams', value: result.groups.length, sub: `${result.m365Groups} M365 · ${result.teams} Teams`, color: 'text-emerald-600 dark:text-emerald-400' },
              { icon: HardDrive, label: 'Total data', value: result.usage.available ? `${(((result.usage.mailboxTotalGB + result.usage.oneDriveTotalGB + result.usage.spoTotalGB) / 1024)).toFixed(1)} TB` : 'n/a', sub: `${result.devices.total} devices`, color: 'text-amber-600 dark:text-amber-400' },
            ].map((k) => (
              <Card key={k.label} className="p-4">
                <div className="flex items-center gap-2 text-xs font-semibold uppercase text-slate-400"><k.icon size={14} /> {k.label}</div>
                <div className={`mt-1 truncate text-2xl font-bold ${k.color}`}>{k.value}</div>
                <div className="text-xs text-slate-400">{k.sub}</div>
              </Card>
            ))}
          </div>

          <div className="grid gap-5 lg:grid-cols-2">
            <Card className="p-5">
              <h3 className="mb-3 text-sm font-semibold uppercase tracking-wide text-slate-500 dark:text-slate-400">Account composition</h3>
              <DonutChart centerLabel="total users" centerValue={String(result.users.length)}
                segments={[
                  { label: 'Active members', value: result.users.length - result.guests - result.disabled, color: '#10b981' },
                  { label: 'Guests', value: result.guests, color: '#8b5cf6' },
                  { label: 'Disabled', value: result.disabled, color: '#94a3b8' },
                ]} />
            </Card>
            <Card className="p-5">
              <h3 className="mb-3 text-sm font-semibold uppercase tracking-wide text-slate-500 dark:text-slate-400">License consumption (top)</h3>
              <HBarChart data={licenseChart} />
            </Card>
          </div>

          {/* ===== MIGRATION ANALYSIS ===== */}
          {analysis && (
            <>
              <div className="grid gap-5 lg:grid-cols-[320px_1fr]">
                <Card className="p-5">
                  <h3 className="mb-3 text-sm font-semibold uppercase tracking-wide text-slate-500 dark:text-slate-400 flex items-center gap-2"><Gauge size={15} /> Migration complexity</h3>
                  <DonutChart centerLabel={analysis.complexityLabel} centerValue={`${analysis.complexityScore}`}
                    segments={[
                      { label: 'Complexity', value: analysis.complexityScore, color: analysis.complexityScore >= 75 ? '#dc2626' : analysis.complexityScore >= 50 ? '#f59e0b' : analysis.complexityScore >= 25 ? '#2563eb' : '#10b981' },
                      { label: 'Headroom', value: 100 - analysis.complexityScore, color: '#e2e8f0' },
                    ]} />
                  <div className="mt-3 space-y-1 text-sm">
                    <div className="flex justify-between"><span className="text-slate-400">Total data</span><span className="font-semibold text-slate-700 dark:text-slate-200">~{analysis.totalDataGB.toLocaleString()} GB</span></div>
                    <div className="flex justify-between"><span className="text-slate-400">Est. window</span><span className="font-semibold text-slate-700 dark:text-slate-200">~{analysis.estimatedDays} days</span></div>
                  </div>
                </Card>
                <Card className="p-5">
                  <h3 className="mb-3 text-sm font-semibold uppercase tracking-wide text-slate-500 dark:text-slate-400 flex items-center gap-2"><HardDrive size={15} /> Sizing & recommended approach</h3>
                  <div className="grid gap-x-6 gap-y-1.5 sm:grid-cols-2">
                    {analysis.sizing.map((s) => (
                      <div key={s.label} className="flex justify-between gap-3 text-sm">
                        <span className="text-slate-400">{s.label}</span>
                        <span className="text-right font-medium text-slate-700 dark:text-slate-200">{s.value}</span>
                      </div>
                    ))}
                  </div>
                  <div className="mt-4 rounded-lg bg-blue-50 dark:bg-blue-900/20 p-3 text-sm text-slate-700 dark:text-slate-200">
                    <span className="font-semibold text-blue-700 dark:text-blue-300">Recommended: </span>{analysis.recommendedApproach}
                  </div>
                </Card>
              </div>

              <Card className="p-5">
                <Section title={`Findings & risks (${analysis.findings.length})`}>
                  <div className="space-y-2">
                    {analysis.findings.map((f, i) => (
                      <div key={i} className="flex items-start gap-2.5 text-sm">
                        <Badge color={f.level === 'blocker' ? 'red' : f.level === 'risk' ? 'orange' : f.level === 'warn' ? 'purple' : 'blue'}>{f.level}</Badge>
                        <span className="font-medium text-slate-600 dark:text-slate-300 shrink-0">{f.area}</span>
                        <span className="text-slate-700 dark:text-slate-200">{f.text}</span>
                      </div>
                    ))}
                  </div>
                </Section>
              </Card>

              <Card className="p-5">
                <Section title="Migration preparation checklist">
                  <div className="space-y-1.5">
                    {analysis.checklist.map((c, i) => (
                      <div key={i} className="flex items-start gap-2.5 text-sm">
                        <ListChecks size={15} className="mt-0.5 shrink-0 text-blue-500" />
                        <span className="w-28 shrink-0 text-xs font-bold uppercase text-slate-400">{c.phase}</span>
                        <span className="text-slate-700 dark:text-slate-200">{c.task}</span>
                      </div>
                    ))}
                  </div>
                </Section>
              </Card>
            </>
          )}

          {ai && (
            <Card className="p-5 border-violet-200 dark:border-violet-800">
              <h3 className="mb-2 flex items-center gap-2 text-sm font-semibold text-violet-700 dark:text-violet-300"><Sparkles size={15} /> AI migration analysis</h3>
              <pre className="whitespace-pre-wrap text-sm text-slate-700 dark:text-slate-200 font-sans leading-relaxed">{ai}</pre>
            </Card>
          )}

          {/* Licenses table */}
          <Card className="p-5">
            <Section title="Licenses">
              <div className="overflow-x-auto">
                <table className="w-full text-sm">
                  <thead><tr className="border-b border-slate-200 dark:border-slate-700 text-left text-xs uppercase text-slate-400"><th className="py-2 pr-3">SKU</th><th className="py-2 pr-3">Enabled</th><th className="py-2 pr-3">Consumed</th><th className="py-2">Available</th></tr></thead>
                  <tbody>
                    {result.licenses.map((l) => (
                      <tr key={l.skuPartNumber} className="border-b border-slate-100 dark:border-slate-700/50">
                        <td className="py-1.5 pr-3 font-medium text-slate-700 dark:text-slate-200">{l.skuPartNumber}</td>
                        <td className="py-1.5 pr-3 text-slate-500">{l.enabled}</td>
                        <td className="py-1.5 pr-3 text-slate-500">{l.consumed}</td>
                        <td className="py-1.5"><Badge color={l.available > 0 ? 'green' : 'orange'}>{l.available}</Badge></td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </Section>
          </Card>

          {/* Domains */}
          <Card className="p-5">
            <Section title="Domains">
              <div className="flex flex-wrap gap-2">
                {result.domains.map((d) => (
                  <span key={d.id} className="flex items-center gap-1.5 rounded-lg border border-slate-200 dark:border-slate-700 px-3 py-1.5 text-sm">
                    <Globe size={13} className="text-blue-500" />{d.id}
                    {d.isDefault && <Badge color="blue">default</Badge>}
                    <Badge color={d.isVerified ? 'green' : 'orange'}>{d.isVerified ? 'verified' : 'unverified'}</Badge>
                  </span>
                ))}
              </div>
            </Section>
          </Card>

          <p className="text-xs text-slate-400">Read-only snapshot · {new Date(result.fetchedAt).toLocaleString()} · {result.users.length} users / {result.groups.length} groups loaded. Export to Excel for the full per-row data.</p>
        </div>
      )}
    </div>
  );
}
