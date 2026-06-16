import { useEffect, useMemo, useRef, useState } from 'react';
import { Radar, ShieldCheck, Loader2, FileSpreadsheet, RefreshCw, Building2, Users2, Boxes, Globe, Sparkles, Gauge, ListChecks, HardDrive, Trash2 } from 'lucide-react';
import { Badge, Button, Card, CopyButton, PageHeader, Section } from '../components/ui';
import { DonutChart, HBarChart } from '../components/charts';
import { runDiscovery, DiscoveryResult, DISCOVERY_SCOPES, LogLevel, useDiscoveryToken } from '../services/graphDiscovery';
import { connectTenant, connectedTenant, disconnectTenant, getDiscoveryToken, discoveryAuthConfigured, getDiscoveryAuth, saveDiscoveryAuth } from '../services/discoveryAuth';
import { saveTenantResult, loadTenantResult, tenantIndex, removeTenantResult, recordExportAudit } from '../services/tenantStore';
import { cloudAgentConfigured } from '../services/cloudDiscovery';
import { requestDeviceCode, pollForToken, getOnecomToken, clearOnecomToken, DeviceCode } from '../services/onecomDeviceAuth';
import { getSession } from '../services/auth';
import { assess } from '../services/migrationAssessment';
import { downloadWorkbook, Sheet } from '../services/excelExport';
import { currentAccount } from '../services/sso';
import { getSSOSettings, aiEnabled } from '../store/settings';
import { load, save } from '../store/useLocalStorage';
import { chat } from '../services/ai';
import { analyzeTemplate, fillAndDownload, TemplateAnalysis, TEMPLATE_DATASETS } from '../services/templateFill';
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
  const analysis = useMemo(() => (result ? assess(result) : null), [result]);
  const [tenants, setTenants] = useState(() => tenantIndex());
  const [connected, setConnected] = useState<{ username: string; tenantId: string } | null>(() => connectedTenant());
  const [connecting, setConnecting] = useState(false);
  const authReady = discoveryAuthConfigured();
  const cloudReady = cloudAgentConfigured();

  // ----- Zero-setup device-code flow via the one.com PHP broker (no agent, no app reg) -----
  const [code, setCode] = useState<DeviceCode | null>(null);
  const [codePhase, setCodePhase] = useState<'idle' | 'awaiting' | 'collecting'>('idle');

  const codeConnect = async () => {
    setError(''); setLines([]); setCode(null); setResult(null);
    try {
      const dc = await requestDeviceCode();
      setCode(dc);
      setCodePhase('awaiting');
      const { tenantId } = await pollForToken(dc);
      setCodePhase('collecting');
      setBusy(true);
      addLine(`Signed in (tenant ${tenantId}). Starting read-only analysis ...`, 'ok');
      useDiscoveryToken(getOnecomToken);
      const r = await runDiscovery(addLine);
      addLine('Read-only assessment complete. No tenant changes were made.', 'ok');
      setResult(r);
      save('discovery-result', r);
      saveTenantResult(r);
      setTenants(tenantIndex());
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
      addLine(`ERROR: ${e instanceof Error ? e.message : e}`, 'err');
    } finally {
      setBusy(false); setCodePhase('idle'); setCode(null); clearOnecomToken();
    }
  };

  const connect = async () => {
    setConnecting(true); setError('');
    try {
      const info = await connectTenant(DISCOVERY_SCOPES);
      setConnected({ username: info.username, tenantId: info.tenantId });
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setConnecting(false);
    }
  };
  const disconnect = async () => {
    await disconnectTenant();
    setConnected(null);
    setResult(null);
    setLines([]);
  };

  /** One-click: sign out of the customer tenant AND erase every trace of its data. */
  const disconnectWipe = async () => {
    const name = result?.org.displayName ?? connected?.tenantId ?? 'this tenant';
    if (!window.confirm(`Disconnect and permanently erase all collected data for "${name}"? This clears the analysis, the stored per-tenant copy, and the sign-in token from this browser.`)) return;
    const tid = result?.org.tenantId ?? connected?.tenantId;
    if (tid) removeTenantResult(tid);
    try { await disconnectTenant(); } catch { /* device-code flow has no MSAL session */ }
    clearOnecomToken();
    save('discovery-result', null);
    setConnected(null);
    setResult(null);
    setAi('');
    setLines([]);
    setTpl(null); setTplDone(null); setTplName('');
    setTenants(tenantIndex());
  };

  // ----- Template-aware export (uses the engineer's own Excel template) -----
  const [tpl, setTpl] = useState<TemplateAnalysis | null>(null);
  const [tplName, setTplName] = useState('');
  const [tplErr, setTplErr] = useState('');
  const [tplDone, setTplDone] = useState<{ sheet: string; rows: number }[] | null>(null);

  const loadTemplate = async (file: File | undefined) => {
    if (!file) return;
    setTplErr(''); setTplDone(null);
    try {
      setTpl(await analyzeTemplate(file));
      setTplName(file.name);
    } catch (e) {
      setTplErr(e instanceof Error ? e.message : String(e));
    }
  };

  const fillTemplate = async () => {
    if (!tpl || !result) return;
    try {
      const summary = await fillAndDownload(tpl, { d: result, a: analysis }, tplName.replace(/\.xlsx?$/i, '') + '-filled.xlsx');
      setTplDone(summary);
    } catch (e) {
      setTplErr(e instanceof Error ? e.message : String(e));
    }
  };

  const run = async () => {
    setBusy(true); setError(''); setAi(''); setLines([]);
    try {
      // Use the connected CUSTOMER tenant token (multi-tenant, no tenant ID).
      // Fall back to the portal SSO token only if no discovery connection is set.
      if (authReady) {
        if (!connected) { setError('Connect a customer tenant first (button above).'); return; }
        useDiscoveryToken(getDiscoveryToken);
      } else {
        if (!(await currentAccount())) {
          setError('Connect a customer tenant (configure the Migration Discovery connection in Settings → Sign-in), or sign in with portal SSO.');
          return;
        }
        useDiscoveryToken(null);
      }
      const r = await runDiscovery(addLine);
      addLine(`Discovery complete — ${r.users.length} users, ${r.groups.length} groups, ~${Math.round((r.usage.mailboxTotalGB + r.usage.oneDriveTotalGB + r.usage.spoTotalGB))} GB data. READ-ONLY: nothing was written.`, 'ok');
      setResult(r);
      save('discovery-result', r);
      saveTenantResult(r); // per-tenant separated storage
      setTenants(tenantIndex());
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
          ['Compliant devices', result.devices.compliant ?? 0],
          ['Non-compliant devices', result.devices.nonCompliant ?? 0],
          ['Compliance policies', (result.compliancePolicies ?? []).length],
          ['Conditional Access policies', result.caPolicies.length],
          ['License SKUs', result.licenses.length],
          ['Licensed seats consumed', result.licenses.reduce((a, l) => a + l.consumed, 0)],
          ['SharePoint sites', result.sharePointSites.length],
          ['Total data (GB)', Math.round(result.usage.mailboxTotalGB + result.usage.oneDriveTotalGB + result.usage.spoTotalGB)],
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
        name: 'Devices by OS',
        columns: ['Operating system', 'Count'],
        rows: Object.entries(result.devices.byOs).map(([os, n]) => [os, n]),
      },
      {
        name: 'Device inventory',
        columns: ['Device', 'User', 'OS', 'OS version', 'Compliance', 'Ownership', 'Manufacturer', 'Model', 'Serial', 'Encrypted', 'Last sync', 'Enrolled'],
        rows: (result.deviceInventory ?? []).map((d) => [d.deviceName, d.user, d.os, d.osVersion, d.compliance, d.ownership, d.manufacturer, d.model, d.serialNumber, d.encrypted ? 'Yes' : 'No', d.lastSync, d.enrolled]),
      },
      {
        name: 'Compliance policies',
        columns: ['Policy', 'Platform'],
        rows: (result.compliancePolicies ?? []).map((p) => [p.name, p.platform]),
      },
      {
        name: 'SharePoint sites',
        columns: ['Name', 'URL', 'Created', 'Last modified'],
        rows: result.sharePointSites.map((s) => [s.name, s.webUrl, s.created, s.lastModified]),
      },
      {
        name: 'Conditional Access',
        columns: ['Policy', 'State'],
        rows: result.caPolicies.map((p) => [p.displayName, p.state]),
      },
      {
        name: 'App registrations',
        columns: ['Display name', 'App ID', 'Audience', 'Created'],
        rows: result.appRegistrations.map((a) => [a.displayName, a.appId, a.signInAudience, a.created]),
      },
      {
        name: 'Enterprise apps',
        columns: ['Display name', 'App ID', 'Type', 'Enabled'],
        rows: result.servicePrincipals.map((s) => [s.displayName, s.appId, s.type, s.enabled ? 'Yes' : 'No']),
      },
      {
        name: 'Teams detail',
        columns: ['Team', 'Visibility', 'Owners', 'Members', 'Guests', 'Channels', 'Private/shared channels'],
        rows: result.teamsDetail.map((t) => [t.name, t.visibility, t.owners, t.members, t.guests, t.channels, t.privateChannels]),
      },
      {
        name: 'Site drives',
        columns: ['Site', 'Drive', 'Type', 'Used GB', 'Total GB'],
        rows: result.siteDrives.map((d) => [d.siteName, d.driveName, d.driveType, d.usedGB, d.totalGB]),
      },
      {
        name: 'Workload readiness',
        columns: ['Workload', 'Status', 'Count', 'Note'],
        rows: result.workloads.map((w) => [w.workload, w.status, w.count, w.note]),
      },
      {
        name: 'Validation items',
        columns: ['Workload', 'Status', 'Note'],
        rows: result.validations.map((v) => [v.workload, v.status, v.note]),
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
    recordExportAudit({ operator: getSession()?.email ?? 'local', tenantId: result.org.tenantId, tenantName: result.org.displayName, scopes: DISCOVERY_SCOPES.length, objects: result.users.length + result.groups.length + result.sharePointSites.length, format: 'xlsx' });
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

      {/* Primary path: enter a code — brokered by the one.com PHP helper (no agent, no app reg) */}
      <Card className="mb-5 p-5 border-emerald-300 dark:border-emerald-700">
        <Section title="Connect a tenant — just enter a code">
          <p className="mb-3 text-sm text-slate-600 dark:text-slate-300">
            No app registration, no tenant ID, nothing installed. Click below, sign in once with the customer's admin
            using the code, and approve the Microsoft Graph consent. The analysis is <strong>read-only</strong>.
          </p>
          {codePhase === 'idle' ? (
            <Button onClick={codeConnect} disabled={busy}>
              {busy ? <Loader2 size={16} className="animate-spin" /> : <Building2 size={16} />} Connect a tenant with a code
            </Button>
          ) : codePhase === 'awaiting' && code ? (
            <div className="rounded-xl border-2 border-emerald-300 dark:border-emerald-700 bg-emerald-50 dark:bg-emerald-900/20 p-4">
              <div className="text-sm text-slate-600 dark:text-slate-300">1. Open this link:</div>
              <a href={code.verification_uri} target="_blank" rel="noreferrer" className="text-lg font-semibold text-blue-600 hover:underline">{code.verification_uri || 'https://microsoft.com/devicelogin'}</a>
              <div className="mt-3 text-sm text-slate-600 dark:text-slate-300">2. Enter this code:</div>
              <div className="flex items-center gap-3">
                <span className="rounded-lg bg-white dark:bg-slate-800 border border-slate-300 dark:border-slate-600 px-4 py-2 font-mono text-2xl font-bold tracking-widest text-slate-800 dark:text-slate-100">{code.user_code}</span>
                <CopyButton text={code.user_code} label="Copy code" />
              </div>
              <div className="mt-3 text-sm text-slate-600 dark:text-slate-300">3. Sign in as the customer's admin and approve. The read-only analysis starts automatically.</div>
              <div className="mt-2 flex items-center gap-2 text-xs text-emerald-700 dark:text-emerald-300"><Loader2 size={13} className="animate-spin" /> Waiting for sign-in…</div>
            </div>
          ) : (
            <div className="flex items-center gap-2 text-sm text-blue-600"><Loader2 size={15} className="animate-spin" /> Signed in — collecting read-only data…</div>
          )}
        </Section>
      </Card>

      {/* Connect the customer tenant — interactive popup login, no tenant ID */}
      <Card className="mb-5 p-5 border-blue-200 dark:border-blue-800">
        <Section title={cloudReady ? 'Or connect with a popup (browser sign-in)' : 'Connect the customer tenant to analyze'}>
          {!authReady ? (
            <SetupWizard onDone={() => { setError(''); }} />
          ) : connected ? (
            <div className="flex flex-wrap items-center gap-3">
              <Badge color="green">Connected</Badge>
              <span className="text-sm font-medium text-slate-700 dark:text-slate-200">{connected.username}</span>
              <span className="text-xs text-slate-400">tenant {connected.tenantId}</span>
              <Button variant="secondary" onClick={disconnect}>Disconnect / switch tenant</Button>
            </div>
          ) : (
            <>
              <p className="mb-3 text-sm text-slate-500 dark:text-slate-400">
                Sign in with the customer's admin account — a Microsoft popup opens. No tenant ID needed. The analysis is <strong>read-only</strong>; nothing in the tenant is changed.
              </p>
              <div className="flex flex-wrap items-center gap-3">
                <Button onClick={connect} disabled={connecting}>
                  {connecting ? <Loader2 size={16} className="animate-spin" /> : <Building2 size={16} />} Connect customer tenant (sign in)
                </Button>
              </div>
            </>
          )}
        </Section>
      </Card>

      <Card className="mb-5 p-5">
        <div className="flex flex-wrap items-center gap-3">
          <Button onClick={run} disabled={busy || (authReady && !connected)}>
            {busy ? <Loader2 size={16} className="animate-spin" /> : <RefreshCw size={16} />} {result ? 'Re-run discovery' : 'Run migration analysis (read-only)'}
          </Button>
          {result && <Button variant="secondary" onClick={exportExcel}><FileSpreadsheet size={15} /> Export Excel workbook</Button>}
          {result && aiEnabled() && <Button variant="ai" onClick={analyze} disabled={aiBusy}>{aiBusy ? <Loader2 size={15} className="animate-spin" /> : <Sparkles size={15} />} AI migration analysis</Button>}
          {(result || connected) && (
            <button onClick={disconnectWipe} disabled={busy} className="inline-flex items-center gap-1.5 rounded-lg border border-red-300 dark:border-red-800 bg-red-50 dark:bg-red-900/20 px-3 py-2 text-sm font-medium text-red-600 dark:text-red-300 hover:bg-red-100 dark:hover:bg-red-900/40 disabled:opacity-50">
              <Trash2 size={15} /> Disconnect &amp; wipe tenant data
            </button>
          )}
          {busy && <span className="text-sm text-slate-500">Querying Microsoft Graph…</span>}
        </div>
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

      {/* Use your own Excel template */}
      {result && (
        <Card className="mb-5 p-5 border-emerald-200 dark:border-emerald-800">
          <Section title="Fill YOUR Excel template (keeps your structure & tabs)">
            <p className="mb-3 text-sm text-slate-500 dark:text-slate-400">
              Upload your own <code>.xlsx</code> template — it stays on your machine (read in the browser, never sent anywhere). WorkPilot detects each tab's header row, maps the columns to the discovered data (English & Dutch headers), fills the rows, and downloads your workbook with the same tabs and layout.
            </p>
            <div className="flex flex-wrap items-center gap-3">
              <label className="inline-flex cursor-pointer items-center gap-2 rounded-lg border border-slate-300 dark:border-slate-600 px-3 py-2 text-sm font-medium text-slate-700 dark:text-slate-200 hover:bg-slate-100 dark:hover:bg-slate-700">
                <FileSpreadsheet size={15} /> Choose template (.xlsx)
                <input type="file" accept=".xlsx,.xls" className="hidden" onChange={(e) => loadTemplate(e.target.files?.[0])} />
              </label>
              {tplName && <span className="text-xs text-slate-400">{tplName}</span>}
              {tplErr && <span className="text-sm text-red-500">{tplErr}</span>}
            </div>

            {tpl && (
              <div className="mt-4 space-y-2">
                <div className="text-xs font-semibold uppercase text-slate-400">Detected tabs — map each to a dataset:</div>
                {tpl.sheets.map((s, i) => (
                  <div key={s.sheetName} className="flex flex-wrap items-center gap-2 rounded-lg border border-slate-200 dark:border-slate-700 p-2.5 text-sm">
                    <span className="font-medium text-slate-700 dark:text-slate-200">{s.sheetName}</span>
                    <span className="text-xs text-slate-400">{s.headers.filter(Boolean).slice(0, 6).join(' · ')}{s.headers.filter(Boolean).length > 6 ? ' …' : ''}</span>
                    <select
                      value={s.datasetId ?? ''}
                      onChange={(e) => setTpl((prev) => prev ? { ...prev, sheets: prev.sheets.map((x, j) => j === i ? { ...x, datasetId: e.target.value || null } : x) } : prev)}
                      className="ml-auto rounded border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-900 px-2 py-1 text-xs"
                    >
                      <option value="">— skip / leave as-is —</option>
                      {TEMPLATE_DATASETS.map((d) => <option key={d.id} value={d.id}>{d.label}</option>)}
                    </select>
                  </div>
                ))}
                <Button onClick={fillTemplate} className="mt-2"><FileSpreadsheet size={15} /> Fill my template & download</Button>
                {tplDone && (
                  <div className="mt-2 rounded-lg bg-emerald-50 dark:bg-emerald-900/20 p-3 text-sm text-emerald-700 dark:text-emerald-300">
                    Done — filled: {tplDone.filter((t) => t.rows > 0).map((t) => `${t.sheet} (${t.rows})`).join(', ') || 'no matching sheets'}. Your template downloaded with its original tabs.
                  </div>
                )}
              </div>
            )}
          </Section>
        </Card>
      )}

      {/* Per-tenant stored assessments */}
      {tenants.length > 0 && (
        <Card className="mb-5 p-4">
          <div className="mb-2 text-xs font-semibold uppercase tracking-wide text-slate-500 dark:text-slate-400">Stored tenant assessments (separated per tenant)</div>
          <div className="flex flex-wrap gap-2">
            {tenants.map((t) => (
              <span key={t.tenantId} className={`flex items-center gap-1.5 rounded-lg border px-2.5 py-1 text-xs ${result?.org.tenantId === t.tenantId ? 'border-blue-400 bg-blue-50 dark:bg-blue-900/20' : 'border-slate-200 dark:border-slate-700'}`}>
                <button onClick={() => { const r = loadTenantResult(t.tenantId); if (r) { setResult(r); save('discovery-result', r); } }} className="font-medium text-slate-700 dark:text-slate-200 hover:text-blue-500">
                  {t.displayName}
                </button>
                <span className="text-slate-400">{t.users}u · {new Date(t.fetchedAt).toLocaleDateString()}</span>
                <button onClick={() => { removeTenantResult(t.tenantId); setTenants(tenantIndex()); }} className="text-slate-400 hover:text-red-500">×</button>
              </span>
            ))}
          </div>
        </Card>
      )}

      {result && (
        <div className="space-y-5">
          {/* Active tenant banner */}
          <Card className="p-4 border-blue-200 dark:border-blue-800">
            <div className="flex flex-wrap items-center gap-2 text-sm">
              <Building2 size={16} className="text-blue-500" />
              <span className="font-semibold text-slate-800 dark:text-slate-100">{result.org.displayName}</span>
              <span className="text-xs text-slate-400">tenant {result.org.tenantId}</span>
              <Badge color="green">Read-only assessment · no tenant changes performed</Badge>
            </div>
          </Card>

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

          {/* Workload readiness + validation */}
          <div className="grid gap-5 lg:grid-cols-2">
            <Card className="p-5">
              <Section title="Migration readiness by workload">
                <div className="space-y-1.5">
                  {result.workloads.map((w) => (
                    <div key={w.workload} className="flex items-center gap-2.5 text-sm">
                      <Badge color={w.status === 'Exported' ? 'green' : w.status === 'Partial' ? 'orange' : w.status === 'Blocked' ? 'red' : 'purple'}>{w.status}</Badge>
                      <span className="font-medium text-slate-700 dark:text-slate-200 w-32 shrink-0">{w.workload}</span>
                      <span className="text-slate-500 dark:text-slate-400">{w.note}</span>
                    </div>
                  ))}
                </div>
              </Section>
            </Card>
            <Card className="p-5">
              <Section title={`Validation items (${result.validations.length})`}>
                {result.validations.length === 0 ? (
                  <p className="text-sm text-emerald-600">No blocked/forbidden endpoints — full read access.</p>
                ) : (
                  <div className="space-y-1.5 max-h-64 overflow-y-auto">
                    {result.validations.map((v, i) => (
                      <div key={i} className="text-sm">
                        <Badge color={v.status === 403 || v.status === 401 ? 'red' : 'orange'}>{v.status || 'err'}</Badge>
                        <span className="ml-2 font-medium text-slate-700 dark:text-slate-200">{v.workload}</span>
                        <div className="text-xs text-slate-400">{v.note}</div>
                      </div>
                    ))}
                  </div>
                )}
              </Section>
            </Card>
          </div>

          {/* Teams detail + SharePoint sharing baseline */}
          {(result.teamsDetail.length > 0 || result.sharing.sampledDrives > 0) && (
            <div className="grid gap-5 lg:grid-cols-2">
              {result.teamsDetail.length > 0 && (
                <Card className="p-5">
                  <Section title={`Teams detail (${result.teamsDetail.length})`}>
                    <div className="overflow-x-auto max-h-72 overflow-y-auto">
                      <table className="w-full text-sm">
                        <thead><tr className="border-b border-slate-200 dark:border-slate-700 text-left text-xs uppercase text-slate-400"><th className="py-1.5 pr-2">Team</th><th className="py-1.5 pr-2">Own.</th><th className="py-1.5 pr-2">Mem.</th><th className="py-1.5 pr-2">Guests</th><th className="py-1.5 pr-2">Chan.</th><th className="py-1.5">Priv.</th></tr></thead>
                        <tbody>
                          {result.teamsDetail.map((t, i) => (
                            <tr key={i} className="border-b border-slate-100 dark:border-slate-700/50">
                              <td className="py-1.5 pr-2 font-medium text-slate-700 dark:text-slate-200">{t.name}</td>
                              <td className="py-1.5 pr-2 text-slate-500">{t.owners}</td>
                              <td className="py-1.5 pr-2 text-slate-500">{t.members}</td>
                              <td className="py-1.5 pr-2">{t.guests > 0 ? <Badge color="orange">{t.guests}</Badge> : <span className="text-slate-400">0</span>}</td>
                              <td className="py-1.5 pr-2 text-slate-500">{t.channels}</td>
                              <td className="py-1.5 text-slate-500">{t.privateChannels}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </Section>
                </Card>
              )}
              {result.sharing.sampledDrives > 0 && (
                <Card className="p-5">
                  <Section title="SharePoint sharing baseline (sampled)">
                    <div className="space-y-2">
                      <div className="flex items-center gap-2 text-sm"><Badge color="red">Anonymous</Badge><span className="font-semibold text-slate-700 dark:text-slate-200">{result.sharing.anonymous}</span><span className="text-slate-400">"Anyone" links</span></div>
                      <div className="flex items-center gap-2 text-sm"><Badge color="orange">Organization</Badge><span className="font-semibold text-slate-700 dark:text-slate-200">{result.sharing.organization}</span><span className="text-slate-400">company-wide links</span></div>
                      <div className="flex items-center gap-2 text-sm"><Badge color="blue">Specific users</Badge><span className="font-semibold text-slate-700 dark:text-slate-200">{result.sharing.users}</span></div>
                      <p className="mt-2 text-xs text-slate-400">Across {result.sharing.sampledDrives} sampled drive(s), {result.sharing.total} permission(s) total. OneDrive sample: {result.oneDriveSample.readable}/{result.oneDriveSample.sampled} readable.</p>
                    </div>
                  </Section>
                </Card>
              )}
            </div>
          )}

          {/* Extended workloads summary */}
          <div className="grid grid-cols-2 gap-4 lg:grid-cols-4">
            {[
              { label: 'SharePoint sites', value: result.sharePointSites.length },
              { label: 'Conditional Access', value: result.caPolicies.length },
              { label: 'App registrations', value: result.appRegistrations.length },
              { label: 'Enterprise apps', value: result.servicePrincipals.length },
              { label: 'Intune configs', value: result.intune.configs },
              { label: 'Compliance policies', value: result.intune.compliance },
              { label: 'Managed devices', value: result.intune.devices },
              { label: 'OneDrive readable', value: result.oneDrive.readable },
            ].map((k) => (
              <Card key={k.label} className="p-4">
                <div className="text-xs font-semibold uppercase text-slate-400">{k.label}</div>
                <div className="mt-1 text-2xl font-bold text-blue-600 dark:text-blue-400">{k.value}</div>
              </Card>
            ))}
          </div>

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

/** First-time, in-portal setup made simple: 3 steps, big buttons, copy fields. */
function SetupWizard({ onDone }: { onDone: () => void }) {
  const [clientId, setClientId] = useState(getDiscoveryAuth().clientId);
  const redirectUri = window.location.origin + window.location.pathname;
  const scopes = DISCOVERY_SCOPES.join(' ');
  const appRegUrl = 'https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/CreateApplicationBlade/quickStartType~/null/isMSAApp~/false';

  return (
    <div className="space-y-4">
      <p className="text-sm text-slate-600 dark:text-slate-300">
        One-time setup (≈3 minutes). After this you just click <strong>Connect</strong> and sign in for every tenant — no tenant ID, no technical steps. The analysis is always <strong>read-only</strong>.
      </p>

      {/* Step 1 */}
      <div className="rounded-xl border border-slate-200 dark:border-slate-700 p-4">
        <div className="mb-2 flex items-center gap-2 text-sm font-semibold text-slate-800 dark:text-slate-100">
          <span className="flex h-6 w-6 items-center justify-center rounded-full bg-blue-600 text-xs font-bold text-white">1</span>
          Open Microsoft and create the connection (once)
        </div>
        <a href={appRegUrl} target="_blank" rel="noreferrer">
          <Button><Building2 size={15} /> Open Microsoft app registration</Button>
        </a>
        <ul className="mt-3 space-y-2 text-sm text-slate-600 dark:text-slate-300">
          <li>• <strong>Name:</strong> WorkPilot Discovery</li>
          <li>• <strong>Supported account types:</strong> choose <em>"Accounts in any organizational directory (multitenant)"</em></li>
          <li className="flex flex-wrap items-center gap-2">• <strong>Redirect URI</strong> (platform "Single-page application"), paste this:
            <code className="rounded bg-slate-100 dark:bg-slate-700 px-1.5 py-0.5 text-xs">{redirectUri}</code>
            <CopyButton text={redirectUri} />
          </li>
          <li className="flex flex-wrap items-center gap-2">• Under <strong>API permissions</strong> → Microsoft Graph → Delegated, add these (paste the list):
            <CopyButton text={scopes} label="Copy permissions" />
          </li>
        </ul>
      </div>

      {/* Step 2 */}
      <div className="rounded-xl border border-slate-200 dark:border-slate-700 p-4">
        <div className="mb-2 flex items-center gap-2 text-sm font-semibold text-slate-800 dark:text-slate-100">
          <span className="flex h-6 w-6 items-center justify-center rounded-full bg-blue-600 text-xs font-bold text-white">2</span>
          Copy the Application (client) ID and paste it here
        </div>
        <div className="flex flex-wrap items-center gap-2">
          <input
            value={clientId}
            onChange={(e) => setClientId(e.target.value.trim())}
            placeholder="00000000-0000-0000-0000-000000000000"
            className="min-w-80 flex-1 rounded-lg border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-900 px-3 py-2 text-sm text-slate-800 dark:text-slate-100 focus:border-blue-500 focus:outline-none"
          />
          <Button onClick={() => { saveDiscoveryAuth({ clientId }); onDone(); }} disabled={!/^[0-9a-f-]{30,}$/i.test(clientId)}>Save</Button>
        </div>
      </div>

      {/* Step 3 */}
      <div className="rounded-xl border border-dashed border-slate-300 dark:border-slate-600 p-4 text-sm text-slate-500 dark:text-slate-400">
        <span className="flex h-6 w-6 items-center justify-center rounded-full bg-slate-400 text-xs font-bold text-white">3</span>
        <span className="ml-2">After saving, the green <strong>Connect customer tenant</strong> button appears. Click it, sign in with the customer's admin, approve once — done. You never touch this setup again.</span>
      </div>
    </div>
  );
}
