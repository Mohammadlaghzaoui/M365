import { useMemo, useRef, useState } from 'react';
import { FileCode2, Upload, Download, Search, Sparkles, Loader2, Gauge, Plus, Trash2, BookOpen, Wand2, FlaskConical, Rocket, CheckCircle2, XCircle, Copy } from 'lucide-react';
import { Badge, Button, Card, CodeBlock, CopyButton, Field, PageHeader, ProgressBar, Section, Select, TextArea } from '../components/ui';
import { DonutChart, HBarChart } from '../components/charts';
import { analyzeGpo, exportCsv, GpoAnalysis, getKnowledgeBase, MatchedSetting, parseGpoXml } from '../services/gpoParser';
import { GpoMapping, gpoTargets, supportColor, seedGpoMappings } from '../data/gpoKnowledge';
import { generateArtifacts, validateProfile, GeneratedArtifacts } from '../services/gpoGenerator';
import { deployProfile } from '../services/gpoDeploy';
import { currentAccount } from '../services/sso';
import { getSSOSettings } from '../store/settings';
import { useLocalStorage, uid } from '../store/useLocalStorage';
import { chat } from '../services/ai';
import { aiEnabled } from '../store/settings';

type Tab = 'analyze' | 'explorer' | 'generator' | 'knowledge';

export default function GpoAdvisor() {
  const [tab, setTab] = useState<Tab>('analyze');
  const [analysis, setAnalysis] = useState<GpoAnalysis | null>(null);
  const [fileNames, setFileNames] = useState<string[]>([]);
  const [error, setError] = useState('');
  const [query, setQuery] = useState('');
  const [supportFilter, setSupportFilter] = useState('all');
  const [aiBusy, setAiBusy] = useState<string | null>(null);
  const [aiText, setAiText] = useState<Record<string, string>>({});

  const handleFiles = async (files: FileList | null) => {
    if (!files?.length) return;
    setError('');
    try {
      const all = [];
      const names: string[] = [];
      for (const f of Array.from(files)) {
        const xml = await f.text();
        all.push(...parseGpoXml(xml, f.name));
        names.push(f.name);
      }
      if (!all.length) { setError('No GPO settings found in the uploaded file(s). Use Get-GPOReport -ReportType XML.'); return; }
      setAnalysis(analyzeGpo(all));
      setFileNames(names);
      setTab('analyze');
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    }
  };

  const loadSample = () => {
    // Built-in sample GPO report so engineers can try the flow without an export.
    const sample = `<?xml version="1.0"?><GPO><Name>WorkPilot Sample Policy</Name>
      <Computer><ExtensionData><Extension><Policy><Name>Turn off real-time protection</Name><State>Enabled</State><Category>Microsoft Defender Antivirus</Category></Policy>
      <Policy><Name>Windows Firewall: Protect all network connections</Name><State>Enabled</State><Category>Windows Firewall</Category></Policy>
      <Policy><Name>Minimum password length</Name><State>Enabled</State><Category>Account Policies</Category></Policy>
      <Policy><Name>Specify intranet Microsoft update service location</Name><State>Enabled</State><Category>Windows Update</Category></Policy></ExtensionData></Computer>
      <User><ExtensionData><Extension><Policy><Name>Prevent access to the command prompt</Name><State>Enabled</State><Category>System</Category></Policy>
      <Policy><Name>Prevent access to registry editing tools</Name><State>Enabled</State><Category>System</Category></Policy>
      <Folder Id="Documents"><Location>\\\\srv\\redirect</Location></Folder>
      <DriveMapSettings><Drive><Properties path="\\\\srv\\share" letter="S"/></Drive></DriveMapSettings>
      <Script><Command>logon.bat</Command></Script>
      <Policy><Name>Internet Explorer Maintenance</Name><State>Enabled</State><Category>Internet Explorer</Category></Policy></ExtensionData></User></GPO>`;
    handleFilesFromText(sample, 'sample-gpo.xml');
  };

  const handleFilesFromText = (xml: string, name: string) => {
    try {
      const parsed = parseGpoXml(xml, name);
      setAnalysis(analyzeGpo(parsed));
      setFileNames([name]);
      setError('');
    } catch (e) { setError(e instanceof Error ? e.message : String(e)); }
  };

  const download = (content: string, filename: string, type: string) => {
    const blob = new Blob([content], { type });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    a.click();
    URL.revokeObjectURL(a.href);
  };

  const exportExcel = (settings: MatchedSetting[]) => {
    // SpreadsheetML (.xls) — opens in Excel without extra libraries.
    const rows = settings.map((s) => `<Row>${[s.name, s.scope, s.category, s.target, s.intunePath, s.setting, s.support, s.confidence + '%'].map((v) => `<Cell><Data ss:Type="String">${String(v).replace(/[<>&]/g, (c) => ({ '<': '&lt;', '>': '&gt;', '&': '&amp;' }[c]!))}</Data></Cell>`).join('')}</Row>`).join('');
    const xml = `<?xml version="1.0"?><?mso-application progid="Excel.Sheet"?><Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"><Worksheet ss:Name="GPO Analysis"><Table><Row>${['Policy', 'Scope', 'Category', 'Target', 'Intune Path', 'Setting', 'Support', 'Confidence'].map((h) => `<Cell><Data ss:Type="String">${h}</Data></Cell>`).join('')}</Row>${rows}</Table></Worksheet></Workbook>`;
    download(xml, 'gpo-analysis.xls', 'application/vnd.ms-excel');
  };

  const filtered = useMemo(() => {
    if (!analysis) return [];
    return analysis.settings.filter((s) => {
      if (supportFilter !== 'all' && s.support !== supportFilter) return false;
      const q = query.toLowerCase();
      return !q || s.name.toLowerCase().includes(q) || s.category.toLowerCase().includes(q) || s.setting.toLowerCase().includes(q);
    });
  }, [analysis, query, supportFilter]);

  const explain = async (s: MatchedSetting) => {
    setAiBusy(s.id);
    try {
      const res = await chat([{ role: 'user', content: `As a Microsoft Intune migration architect, explain how to migrate this GPO setting to Intune and WHY.\n\nGPO setting: "${s.name}" (${s.scope} configuration, category ${s.category}).\nProposed target: ${s.target} — ${s.intunePath}.\nSupport status: ${s.support}.\nGive concrete steps, the exact Intune location, and any gotcha. Be concise.` }]);
      setAiText((prev) => ({ ...prev, [s.id]: res }));
    } catch (e) {
      setAiText((prev) => ({ ...prev, [s.id]: `AI error: ${e instanceof Error ? e.message : e}` }));
    } finally {
      setAiBusy(null);
    }
  };

  return (
    <div>
      <PageHeader title="GPO → Intune Migration Advisor" subtitle="Upload Get-GPOReport XML exports. WorkPilot parses every policy, matches it to its Intune Settings Catalog / Endpoint Security / script / Win32 equivalent, scores migration readiness and exports the report." icon={<FileCode2 size={20} />} />

      <div className="mb-5 flex flex-wrap gap-2">
        {([['analyze', 'Readiness & analysis'], ['explorer', 'Policy explorer'], ['generator', 'Generator & tester'], ['knowledge', 'Knowledge engine']] as [Tab, string][]).map(([id, label]) => (
          <button key={id} onClick={() => setTab(id)}
            className={`rounded-lg px-4 py-2 text-sm font-semibold ${tab === id ? 'bg-blue-600 text-white' : 'border border-slate-300 dark:border-slate-600 text-slate-600 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-700'}`}>
            {label}
          </button>
        ))}
      </div>

      {/* Upload zone */}
      {tab !== 'knowledge' && (
        <Card className="mb-5 p-5">
          <label className="flex cursor-pointer flex-col items-center justify-center gap-2 rounded-xl border-2 border-dashed border-slate-300 dark:border-slate-600 p-8 text-center hover:border-blue-400">
            <Upload size={28} className="text-blue-500" />
            <span className="text-sm font-semibold text-slate-700 dark:text-slate-200">Drop or select GPO XML file(s)</span>
            <span className="text-xs text-slate-400">Get-GPOReport -All -ReportType XML · multiple files supported</span>
            <input type="file" accept=".xml,text/xml" multiple className="hidden" onChange={(e) => handleFiles(e.target.files)} />
          </label>
          <div className="mt-3 flex flex-wrap items-center gap-3">
            <Button variant="secondary" onClick={loadSample}><FileCode2 size={15} /> Try with sample GPO</Button>
            {fileNames.length > 0 && <span className="text-xs text-slate-400">Loaded: {fileNames.join(', ')}</span>}
            {error && <span className="text-sm text-red-500">{error}</span>}
          </div>
        </Card>
      )}

      {/* ===== ANALYZE ===== */}
      {tab === 'analyze' && analysis && (
        <div className="space-y-5">
          <div className="grid gap-4 lg:grid-cols-[320px_1fr]">
            <Card className="p-5">
              <h3 className="mb-3 text-sm font-semibold uppercase tracking-wide text-slate-500 dark:text-slate-400">Migration readiness</h3>
              <DonutChart centerLabel="ready to migrate" centerValue={`${analysis.readiness}%`}
                segments={[
                  { label: 'Supported', value: analysis.supported, color: '#10b981' },
                  { label: 'Script required', value: analysis.scriptRequired, color: '#f59e0b' },
                  { label: 'Manual', value: analysis.manual, color: '#8b5cf6' },
                  { label: 'Unsupported', value: analysis.unsupported, color: '#dc2626' },
                ]} />
            </Card>
            <Card className="p-5">
              <h3 className="mb-3 text-sm font-semibold uppercase tracking-wide text-slate-500 dark:text-slate-400">By recommended target</h3>
              <HBarChart data={analysis.byTarget} />
            </Card>
          </div>
          <div className="grid grid-cols-2 gap-4 lg:grid-cols-5">
            <Card className="p-4"><div className="text-xs uppercase text-slate-400">Total policies</div><div className="mt-1 text-2xl font-bold text-blue-600 dark:text-blue-400">{analysis.total}</div></Card>
            <Card className="p-4"><div className="text-xs uppercase text-slate-400">Supported</div><div className="mt-1 text-2xl font-bold text-emerald-500">{analysis.supported}</div></Card>
            <Card className="p-4"><div className="text-xs uppercase text-slate-400">Script required</div><div className="mt-1 text-2xl font-bold text-amber-500">{analysis.scriptRequired}</div></Card>
            <Card className="p-4"><div className="text-xs uppercase text-slate-400">Manual</div><div className="mt-1 text-2xl font-bold text-violet-500">{analysis.manual}</div></Card>
            <Card className="p-4"><div className="text-xs uppercase text-slate-400">Unsupported</div><div className="mt-1 text-2xl font-bold text-red-500">{analysis.unsupported}</div></Card>
          </div>
          <Card className="p-5 flex flex-wrap items-center gap-3">
            <Gauge size={18} className="text-blue-500" />
            <span className="text-sm text-slate-600 dark:text-slate-300">Open the <button onClick={() => setTab('explorer')} className="font-semibold text-blue-500 hover:underline">Policy Explorer</button> for the per-setting table, AI explanations and export.</span>
            <div className="ml-auto flex gap-2">
              <Button variant="secondary" onClick={() => download(exportCsv(analysis.settings), 'gpo-analysis.csv', 'text/csv')}><Download size={14} /> CSV</Button>
              <Button variant="secondary" onClick={() => exportExcel(analysis.settings)}><Download size={14} /> Excel</Button>
            </div>
          </Card>
        </div>
      )}
      {tab === 'analyze' && !analysis && <Card className="p-8 text-center text-sm text-slate-400">Upload a GPO export or try the sample to see the readiness score.</Card>}

      {/* ===== EXPLORER ===== */}
      {tab === 'explorer' && (
        analysis ? (
          <div className="space-y-4">
            <Card className="p-4 flex flex-wrap items-center gap-3">
              <div className="relative flex-1 min-w-48">
                <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
                <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Search policies…" className="w-full rounded-lg border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-900 py-2 pl-9 pr-3 text-sm text-slate-800 dark:text-slate-100 focus:border-blue-500 focus:outline-none" />
              </div>
              <div className="w-48"><Select label="" value={supportFilter} onChange={setSupportFilter} options={[{ value: 'all', label: 'All statuses' }, ...['Supported', 'ScriptRequired', 'Manual', 'Unsupported'].map((s) => ({ value: s, label: s }))]} /></div>
              <Button variant="secondary" onClick={() => download(exportCsv(filtered), 'gpo-analysis.csv', 'text/csv')}><Download size={14} /> CSV</Button>
              <Button variant="secondary" onClick={() => exportExcel(filtered)}><Download size={14} /> Excel</Button>
            </Card>
            <Card className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="border-b border-slate-200 dark:border-slate-700 text-left text-xs uppercase text-slate-400">
                    <th className="p-3">Policy</th><th className="p-3">Scope</th><th className="p-3">Recommended Intune target</th><th className="p-3">Status</th><th className="p-3">Confidence</th><th className="p-3"></th>
                  </tr>
                </thead>
                <tbody>
                  {filtered.map((s) => (
                    <>
                      <tr key={s.id} className="border-b border-slate-100 dark:border-slate-700/50 align-top">
                        <td className="p-3"><div className="font-medium text-slate-700 dark:text-slate-200">{s.name}</div><div className="text-xs text-slate-400">{s.category}</div></td>
                        <td className="p-3"><Badge color={s.scope === 'Computer' ? 'blue' : 'purple'}>{s.scope}</Badge></td>
                        <td className="p-3"><div className="font-medium text-slate-700 dark:text-slate-200">{s.target}</div><div className="text-xs text-slate-400">{s.intunePath}</div><div className="text-xs text-slate-500 dark:text-slate-300">{s.setting}</div></td>
                        <td className="p-3"><Badge color={supportColor[s.support]}>{s.support}</Badge></td>
                        <td className="p-3">
                          <div className="flex items-center gap-1.5">
                            <div className="h-1.5 w-12 overflow-hidden rounded-full bg-slate-200 dark:bg-slate-700"><div className={`h-full ${s.confidence >= 85 ? 'bg-emerald-500' : s.confidence >= 60 ? 'bg-amber-500' : 'bg-red-500'}`} style={{ width: `${s.confidence}%` }} /></div>
                            <span className="text-xs font-semibold text-slate-600 dark:text-slate-300">{s.confidence}%</span>
                          </div>
                        </td>
                        <td className="p-3">{aiEnabled() && <button onClick={() => explain(s)} disabled={aiBusy === s.id} className="flex items-center gap-1 text-xs font-semibold text-violet-500 hover:underline">{aiBusy === s.id ? <Loader2 size={12} className="animate-spin" /> : <Sparkles size={12} />} Explain</button>}</td>
                      </tr>
                      {(aiText[s.id] || s.notes) && (
                        <tr key={s.id + 'n'} className="border-b border-slate-100 dark:border-slate-700/50">
                          <td colSpan={6} className="px-3 pb-3">
                            {s.notes && <div className="rounded bg-amber-50 dark:bg-amber-900/20 px-3 py-2 text-xs text-amber-700 dark:text-amber-300">{s.notes}</div>}
                            {aiText[s.id] && <div className="mt-2 rounded bg-violet-50 dark:bg-violet-900/20 px-3 py-2 text-xs text-slate-700 dark:text-slate-200 whitespace-pre-wrap">{aiText[s.id]}</div>}
                          </td>
                        </tr>
                      )}
                    </>
                  ))}
                </tbody>
              </table>
            </Card>
          </div>
        ) : <Card className="p-8 text-center text-sm text-slate-400">Upload a GPO export first.</Card>
      )}

      {/* ===== GENERATOR & TESTER ===== */}
      {tab === 'generator' && (analysis ? <GeneratorTab settings={analysis.settings} /> : <Card className="p-8 text-center text-sm text-slate-400">Upload a GPO export (or try the sample) first — the generator builds Intune artifacts from the analysis.</Card>)}

      {/* ===== KNOWLEDGE ENGINE ===== */}
      {tab === 'knowledge' && <KnowledgeEngine />}
    </div>
  );
}

function GeneratorTab({ settings }: { settings: MatchedSetting[] }) {
  const [profileName, setProfileName] = useState('WorkPilot - Migrated GPO');
  const [art, setArt] = useState<GeneratedArtifacts | null>(null);
  const [validation, setValidation] = useState<ReturnType<typeof validateProfile> | null>(null);
  const [log, setLog] = useState<{ text: string; level: string }[]>([]);
  const [busy, setBusy] = useState<'test' | 'deploy' | null>(null);
  const logRef = useRef<HTMLDivElement>(null);
  const ssoReady = getSSOSettings().enabled && !!getSSOSettings().clientId;

  const generate = () => {
    const a = generateArtifacts(settings, profileName);
    setArt(a);
    setValidation(validateProfile(a.settingsCatalogJson));
    setLog([]);
  };

  const addLog = (text: string, level = 'info') => {
    setLog((prev) => [...prev, { text: `[${new Date().toLocaleTimeString('en-GB')}] ${text}`, level }]);
    setTimeout(() => logRef.current?.scrollTo({ top: logRef.current.scrollHeight }), 30);
  };

  const download = (content: string, filename: string, type = 'text/plain') => {
    const blob = new Blob([content], { type });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    a.click();
    URL.revokeObjectURL(a.href);
  };

  const runDeploy = async (mode: 'test' | 'deploy') => {
    if (!art) return;
    setBusy(mode);
    setLog([]);
    try {
      if (!(await currentAccount())) {
        addLog('Not signed in with Microsoft. Sign in via Settings → Sign-in to test/deploy against your tenant.', 'err');
        return;
      }
      const r = await deployProfile(art.settingsCatalogJson, mode,
        (line, level) => addLog(line, level ?? 'info'));
      addLog(mode === 'test'
        ? `RESULT: profile is valid and deployable (${r.verifiedSettings} settings verified, test policy removed).`
        : `RESULT: policy "${r.name}" created in Intune (id ${r.id}), unassigned.`, 'ok');
    } catch (e) {
      addLog(`FAILED: ${e instanceof Error ? e.message : e}`, 'err');
      addLog('Tip: the SSO app needs DeviceManagementConfiguration.ReadWrite.All (admin consent).', 'warn');
    } finally {
      setBusy(null);
    }
  };

  return (
    <div className="space-y-5">
      <Card className="p-5">
        <Section title="Generate Intune artifacts from the analysis">
          <div className="flex flex-wrap items-end gap-3">
            <div className="flex-1 min-w-64"><Field label="Profile name" value={profileName} onChange={setProfileName} /></div>
            <Button onClick={generate}><Wand2 size={16} /> Generate</Button>
          </div>
          {art && (
            <div className="mt-3 flex flex-wrap gap-2 text-xs">
              <Badge color="green">{art.includedCount} settings → Settings Catalog</Badge>
              <Badge color="orange">{art.scriptCount} → remediation scripts</Badge>
              <Badge color="purple">{art.manualCount} → manual</Badge>
            </div>
          )}
        </Section>
      </Card>

      {art && validation && (
        <Card className={`p-5 ${validation.ok ? 'border-emerald-300 dark:border-emerald-700' : 'border-red-300 dark:border-red-700'}`}>
          <Section title="Tester — profile validation">
            <div className="space-y-1.5">
              {validation.checks.map((c) => (
                <div key={c.label} className="flex items-center gap-2 text-sm">
                  {c.pass ? <CheckCircle2 size={15} className="text-emerald-500" /> : <XCircle size={15} className="text-red-500" />}
                  <span className="text-slate-700 dark:text-slate-200">{c.label}</span>
                  <span className="text-slate-400">— {c.detail}</span>
                </div>
              ))}
            </div>
            <div className="mt-4 flex flex-wrap items-center gap-2">
              <Button variant="ai" onClick={() => runDeploy('test')} disabled={!validation.ok || !!busy}>
                {busy === 'test' ? <Loader2 size={15} className="animate-spin" /> : <FlaskConical size={15} />} Test in tenant (create + verify + delete)
              </Button>
              <Button onClick={() => runDeploy('deploy')} disabled={!validation.ok || !!busy}>
                {busy === 'deploy' ? <Loader2 size={15} className="animate-spin" /> : <Rocket size={15} />} Deploy to Intune (unassigned)
              </Button>
              {!ssoReady && <span className="text-xs text-amber-600 dark:text-amber-400">Sign in with Microsoft (Settings → Sign-in) to test/deploy live.</span>}
            </div>
            <p className="mt-2 text-xs text-slate-400">Safe by design: the test creates the policy <strong>unassigned</strong> (no device receives it), verifies it imported, then deletes it. Requires the SSO app to have <code>DeviceManagementConfiguration.ReadWrite.All</code>.</p>
            {log.length > 0 && (
              <div className="mt-3 overflow-hidden rounded-xl border border-slate-700 bg-[#012456]">
                <div className="flex items-center gap-2 border-b border-slate-700 bg-slate-900 px-3 py-2">
                  <span className="h-3 w-3 rounded-full bg-red-500" /><span className="h-3 w-3 rounded-full bg-amber-400" /><span className="h-3 w-3 rounded-full bg-emerald-500" />
                  <span className="ml-2 text-xs font-medium text-slate-300">Intune deployment — live (Microsoft Graph)</span>
                </div>
                <div ref={logRef} className="h-48 overflow-y-auto p-3 font-mono text-xs leading-relaxed">
                  {log.map((l, i) => <div key={i} className={`whitespace-pre-wrap ${{ info: 'text-cyan-300', ok: 'text-emerald-400', err: 'text-red-400', warn: 'text-amber-300' }[l.level] ?? 'text-slate-300'}`}>{l.text}</div>)}
                  {busy && <span className="inline-block h-3.5 w-2 animate-pulse bg-slate-200 align-middle" />}
                </div>
              </div>
            )}
          </Section>
        </Card>
      )}

      {art && (
        <>
          <Card className="p-5">
            <Section title="Settings Catalog profile (Graph JSON)" action={<div className="flex gap-2"><CopyButton text={art.settingsCatalogJson} /><Button variant="secondary" onClick={() => download(art.settingsCatalogJson, `${art.profileName.replace(/[^a-z0-9]+/gi, '-')}.json`, 'application/json')}><Download size={14} /> .json</Button></div>}>
              <pre className="max-h-80 overflow-auto rounded-lg bg-slate-900 dark:bg-slate-950 border border-slate-700 p-3 text-xs leading-relaxed text-emerald-300 font-mono whitespace-pre-wrap">{art.settingsCatalogJson}</pre>
            </Section>
          </Card>
          <Card className="p-5">
            <Section title="PowerShell deployment script" action={<Button variant="secondary" onClick={() => download(art.deployScript, 'deploy-intune-profile.ps1')}><Download size={14} /> .ps1</Button>}>
              <CodeBlock code={art.deployScript} />
            </Section>
          </Card>
          {art.remediationScripts.length > 0 && (
            <Card className="p-5">
              <Section title={`Remediation script stubs (${art.remediationScripts.length})`}>
                <div className="space-y-3">
                  {art.remediationScripts.map((r) => (
                    <div key={r.name}>
                      <div className="mb-1 flex items-center justify-between">
                        <span className="text-sm font-medium text-slate-600 dark:text-slate-300">{r.name}</span>
                        <Button variant="secondary" onClick={() => download(r.script, r.name)}><Download size={13} /> .ps1</Button>
                      </div>
                      <CodeBlock code={r.script} />
                    </div>
                  ))}
                </div>
              </Section>
            </Card>
          )}
          {art.manualSteps.length > 0 && (
            <Card className="p-5">
              <Section title={`Manual migration items (${art.manualSteps.length})`} action={<CopyButton text={art.manualSteps.join('\n')} />}>
                <ul className="space-y-1.5">
                  {art.manualSteps.map((m, i) => <li key={i} className="text-sm text-slate-600 dark:text-slate-300">• {m}</li>)}
                </ul>
              </Section>
            </Card>
          )}
        </>
      )}
    </div>
  );
}

function KnowledgeEngine() {
  const [custom, setCustom] = useLocalStorage<GpoMapping[]>('gpo-custom-mappings', []);
  const [removed, setRemoved] = useLocalStorage<string[]>('gpo-removed-mappings', []);
  const [draft, setDraft] = useState<Partial<GpoMapping>>({ scope: 'Computer', target: 'Settings Catalog', support: 'Supported', confidence: 90 });
  const kb = getKnowledgeBase();

  const add = () => {
    if (!draft.gpoName || !draft.match) return;
    const m: GpoMapping = {
      id: uid(), custom: true,
      match: String(draft.match).split(',').map((x) => x.trim()).filter(Boolean),
      gpoName: draft.gpoName, scope: (draft.scope as GpoMapping['scope']) ?? 'Computer',
      category: draft.category ?? 'Custom', target: (draft.target as GpoMapping['target']) ?? 'Settings Catalog',
      intunePath: draft.intunePath ?? '', setting: draft.setting ?? draft.gpoName,
      support: (draft.support as GpoMapping['support']) ?? 'Supported', confidence: Number(draft.confidence) || 90, notes: draft.notes,
    };
    setCustom((prev) => [m, ...prev]);
    setDraft({ scope: 'Computer', target: 'Settings Catalog', support: 'Supported', confidence: 90 });
  };

  return (
    <div className="space-y-5">
      <Card className="p-5">
        <Section title="Add a knowledge-base mapping">
          <div className="grid gap-3 md:grid-cols-3">
            <Field label="GPO setting name" value={draft.gpoName ?? ''} onChange={(v) => setDraft({ ...draft, gpoName: v })} />
            <Field label="Match keywords (comma separated)" value={String(draft.match ?? '')} onChange={(v) => setDraft({ ...draft, match: v as never })} placeholder="disablecmd, command prompt" />
            <Field label="Category" value={draft.category ?? ''} onChange={(v) => setDraft({ ...draft, category: v })} />
            <Select label="Scope" value={draft.scope as string} onChange={(v) => setDraft({ ...draft, scope: v as never })} options={['Computer', 'User', 'Both'].map((s) => ({ value: s, label: s }))} />
            <Select label="Target" value={draft.target as string} onChange={(v) => setDraft({ ...draft, target: v as never })} options={gpoTargets.map((t) => ({ value: t, label: t }))} />
            <Select label="Support" value={draft.support as string} onChange={(v) => setDraft({ ...draft, support: v as never })} options={['Supported', 'ScriptRequired', 'Manual', 'Unsupported'].map((s) => ({ value: s, label: s }))} />
            <Field label="Intune path" value={draft.intunePath ?? ''} onChange={(v) => setDraft({ ...draft, intunePath: v })} />
            <Field label="Setting" value={draft.setting ?? ''} onChange={(v) => setDraft({ ...draft, setting: v })} />
            <Field label="Confidence %" type="number" value={String(draft.confidence ?? '')} onChange={(v) => setDraft({ ...draft, confidence: Number(v) })} />
          </div>
          <div className="mt-3"><Button onClick={add}><Plus size={15} /> Add mapping</Button></div>
        </Section>
      </Card>
      <Card className="overflow-x-auto">
        <table className="w-full text-sm">
          <thead><tr className="border-b border-slate-200 dark:border-slate-700 text-left text-xs uppercase text-slate-400"><th className="p-3">GPO setting</th><th className="p-3">Target</th><th className="p-3">Status</th><th className="p-3">Conf.</th><th className="p-3"></th></tr></thead>
          <tbody>
            {kb.map((m) => (
              <tr key={m.id} className="border-b border-slate-100 dark:border-slate-700/50">
                <td className="p-3"><div className="font-medium text-slate-700 dark:text-slate-200">{m.gpoName} {m.custom && <Badge color="purple">custom</Badge>}</div><div className="text-xs text-slate-400">{m.intunePath}</div></td>
                <td className="p-3 text-slate-600 dark:text-slate-300">{m.target}</td>
                <td className="p-3"><Badge color={supportColor[m.support]}>{m.support}</Badge></td>
                <td className="p-3 text-slate-500">{m.confidence}%</td>
                <td className="p-3">
                  {m.custom ? <button onClick={() => setCustom((prev) => prev.filter((x) => x.id !== m.id))} className="text-slate-400 hover:text-red-500"><Trash2 size={14} /></button>
                    : <button onClick={() => setRemoved((prev) => [...prev, m.id])} className="text-xs text-slate-400 hover:text-red-500">hide</button>}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </Card>
      {removed.length > 0 && <button onClick={() => setRemoved([])} className="text-xs font-semibold text-blue-500 hover:underline"><BookOpen size={12} className="inline" /> Restore {removed.length} hidden seed mapping(s)</button>}
    </div>
  );
}
