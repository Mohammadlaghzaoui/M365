import { useState } from 'react';
import { FolderKanban, Plus, Trash2, ChevronDown, ChevronRight, FileText } from 'lucide-react';
import { Badge, Button, Card, Field, PageHeader, ProgressBar, Select, TextArea, TextOutput, EmptyState, riskColor } from '../components/ui';
import { MigrationProject, MigrationStage, MigrationType } from '../types';
import { uid, useLocalStorage } from '../store/useLocalStorage';
import { buildStages, migrationTypeLabels } from '../data/migrationStages';
import { t2tManualGuide, gmailToM365Guide, exchangeOnPremGuide, GuideSection } from '../data/migrationGuides';
import { AIHelper } from '../components/AIHelper';

function projectCompletion(p: MigrationProject): number {
  const tasks = p.stages.flatMap((s) => s.tasks);
  if (!tasks.length) return 0;
  return Math.round((tasks.filter((t) => t.done).length / tasks.length) * 100);
}

function generatePlan(p: MigrationProject): string {
  return `MIGRATION PLAN — ${p.name}
Customer: ${p.customer}
Type: ${migrationTypeLabels[p.type]}
Source: ${p.sourceEnv} → Destination: ${p.destTenant}
Domains: ${p.domains}
Scope: ${p.users} users, ${p.sharedMailboxes} shared mailboxes, ${p.groups} groups, ${p.teams} Teams
Tool: ${p.tool} | Cutover: ${p.cutoverDate || 'TBD'}
Mailbox size notes: ${p.mailboxSizeNotes || '-'}

STAGES
${p.stages.map((s, i) => `${i + 1}. ${s.name} [${s.status}] — owner: ${s.owner || 'TBD'}, due: ${s.dueDate || 'TBD'}, risk: ${s.risk}
${s.tasks.map((t) => `   ${t.done ? '[x]' : '[ ]'} ${t.label}`).join('\n')}`).join('\n\n')}

NOTES
${p.notes || '-'}`;
}

function checklist(p: MigrationProject, stageNames: string[], title: string): string {
  const stages = p.stages.filter((s) => stageNames.includes(s.name));
  return `${title} — ${p.name}\n\n${stages.map((s) => `## ${s.name}\n${s.tasks.map((t) => `${t.done ? '[x]' : '[ ]'} ${t.label}`).join('\n')}`).join('\n\n')}`;
}

const rollback = (p: MigrationProject) => `ROLLBACK CHECKLIST — ${p.name}
[ ] Decision point: rollback only BEFORE batch completion / DNS final switch
[ ] Stop/Remove migration batches that are not completed
[ ] Revert DNS: restore original MX/Autodiscover/SPF values (documented in DNS runbook)
[ ] Re-enable source access for users (licenses/sign-in)
[ ] Communicate rollback + next attempt window to users
[ ] Root-cause analysis before re-planning cutover
NOTE (cross-tenant): after Complete-MigrationBatch the source mailbox is converted — rollback means a reverse migration project.`;

const guides: Record<MigrationType, GuideSection[]> = {
  'm365-to-m365': t2tManualGuide,
  'google-to-m365': gmailToM365Guide,
  'exchange-onprem-to-m365': exchangeOnPremGuide,
};

export default function Migration() {
  const [projects, setProjects] = useLocalStorage<MigrationProject[]>('migration-projects', []);
  const [selectedId, setSelectedId] = useState<string | null>(projects[0]?.id ?? null);
  const [creating, setCreating] = useState(projects.length === 0);
  const [openStage, setOpenStage] = useState<string | null>(null);
  const [output, setOutput] = useState<{ title: string; text: string } | null>(null);
  const [draft, setDraft] = useState<Partial<MigrationProject>>({ type: 'm365-to-m365', tool: 'BitTitan MigrationWiz' });

  const selected = projects.find((p) => p.id === selectedId) ?? null;

  const createProject = () => {
    const p: MigrationProject = {
      id: uid(), createdAt: new Date().toISOString(),
      name: draft.name || 'New migration', customer: draft.customer || '',
      type: (draft.type as MigrationType) || 'm365-to-m365',
      sourceEnv: draft.sourceEnv || '', destTenant: draft.destTenant || '',
      domains: draft.domains || '', users: Number(draft.users) || 0,
      sharedMailboxes: Number(draft.sharedMailboxes) || 0, groups: Number(draft.groups) || 0,
      teams: Number(draft.teams) || 0, mailboxSizeNotes: draft.mailboxSizeNotes || '',
      tool: draft.tool || 'BitTitan MigrationWiz', cutoverDate: draft.cutoverDate || '',
      status: 'planning', notes: '', stages: buildStages(),
    };
    setProjects((prev) => [p, ...prev]);
    setSelectedId(p.id);
    setCreating(false);
    setOutput(null);
  };

  const updateProject = (fn: (p: MigrationProject) => MigrationProject) =>
    setProjects((prev) => prev.map((p) => (p.id === selectedId ? fn(p) : p)));

  const updateStage = (stageId: string, fn: (s: MigrationStage) => MigrationStage) =>
    updateProject((p) => ({ ...p, stages: p.stages.map((s) => (s.id === stageId ? fn(s) : s)) }));

  return (
    <div>
      <PageHeader title="Migration Assistant" subtitle="Manage migration projects across all 15 stages, with generated plans, checklists and the manual runbooks per migration type." icon={<FolderKanban size={20} />} />

      <div className="mb-4 flex flex-wrap items-center gap-2">
        {projects.map((p) => (
          <button key={p.id} onClick={() => { setSelectedId(p.id); setCreating(false); setOutput(null); }}
            className={`rounded-lg px-3 py-1.5 text-sm font-medium ${selectedId === p.id && !creating ? 'bg-blue-600 text-white' : 'border border-slate-300 dark:border-slate-600 text-slate-600 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-700'}`}>
            {p.name}
          </button>
        ))}
        <Button variant="secondary" onClick={() => setCreating(true)}><Plus size={15} /> New project</Button>
      </div>

      {creating && (
        <Card className="p-5 mb-6">
          <h3 className="mb-3 font-semibold text-slate-700 dark:text-slate-200">New migration project</h3>
          <div className="grid gap-3 md:grid-cols-3">
            <Field label="Project name" value={draft.name ?? ''} onChange={(v) => setDraft({ ...draft, name: v })} placeholder="Contoso T2T 2026" />
            <Field label="Customer name" value={draft.customer ?? ''} onChange={(v) => setDraft({ ...draft, customer: v })} />
            <Select label="Migration type" value={draft.type as string} onChange={(v) => setDraft({ ...draft, type: v as MigrationType })}
              options={Object.entries(migrationTypeLabels).map(([value, label]) => ({ value, label }))} />
            <Field label="Source environment" value={draft.sourceEnv ?? ''} onChange={(v) => setDraft({ ...draft, sourceEnv: v })} placeholder="old-tenant.onmicrosoft.com / Google / Exchange 2016" />
            <Field label="Destination tenant" value={draft.destTenant ?? ''} onChange={(v) => setDraft({ ...draft, destTenant: v })} placeholder="new-tenant.onmicrosoft.com" />
            <Field label="Domain names" value={draft.domains ?? ''} onChange={(v) => setDraft({ ...draft, domains: v })} placeholder="contoso.com, contoso.nl" />
            <Field label="Number of users" type="number" value={String(draft.users ?? '')} onChange={(v) => setDraft({ ...draft, users: Number(v) })} />
            <Field label="Shared mailboxes" type="number" value={String(draft.sharedMailboxes ?? '')} onChange={(v) => setDraft({ ...draft, sharedMailboxes: Number(v) })} />
            <Field label="Groups" type="number" value={String(draft.groups ?? '')} onChange={(v) => setDraft({ ...draft, groups: Number(v) })} />
            <Field label="Teams" type="number" value={String(draft.teams ?? '')} onChange={(v) => setDraft({ ...draft, teams: Number(v) })} />
            <Field label="Migration tool" value={draft.tool ?? ''} onChange={(v) => setDraft({ ...draft, tool: v })} />
            <Field label="Cutover date" type="date" value={draft.cutoverDate ?? ''} onChange={(v) => setDraft({ ...draft, cutoverDate: v })} />
          </div>
          <div className="mt-3">
            <TextArea label="Mailbox size notes" value={draft.mailboxSizeNotes ?? ''} onChange={(v) => setDraft({ ...draft, mailboxSizeNotes: v })} rows={2} />
          </div>
          <div className="mt-4 flex gap-2">
            <Button onClick={createProject}>Create project (15 stages)</Button>
            {projects.length > 0 && <Button variant="secondary" onClick={() => setCreating(false)}>Cancel</Button>}
          </div>
        </Card>
      )}

      {!creating && !selected && <EmptyState message="Create your first migration project to start tracking stages." />}

      {!creating && selected && (
        <div className="space-y-5">
          <Card className="p-5">
            <div className="flex flex-wrap items-center gap-3">
              <div className="min-w-0 flex-1">
                <h2 className="text-lg font-bold text-slate-800 dark:text-slate-100">{selected.name}</h2>
                <p className="text-sm text-slate-500 dark:text-slate-400">{selected.customer} · {migrationTypeLabels[selected.type]} · {selected.users} users · tool: {selected.tool} · cutover: {selected.cutoverDate || 'TBD'}</p>
              </div>
              <select value={selected.status} onChange={(e) => updateProject((p) => ({ ...p, status: e.target.value as MigrationProject['status'] }))}
                className="rounded-lg border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-900 px-3 py-1.5 text-sm">
                {['planning', 'active', 'cutover', 'completed', 'on-hold'].map((s) => <option key={s} value={s}>{s}</option>)}
              </select>
              <button onClick={() => { setProjects((prev) => prev.filter((p) => p.id !== selected.id)); setSelectedId(null); }} className="text-slate-400 hover:text-red-500"><Trash2 size={16} /></button>
            </div>
            <div className="mt-3 flex items-center gap-3">
              <ProgressBar value={projectCompletion(selected)} />
              <span className="text-sm font-bold text-blue-600 dark:text-blue-400 whitespace-nowrap">{projectCompletion(selected)}%</span>
            </div>
            <div className="mt-4 flex flex-wrap gap-2">
              <Button variant="secondary" onClick={() => setOutput({ title: 'Migration plan', text: generatePlan(selected) })}><FileText size={14} /> Migration plan</Button>
              <Button variant="secondary" onClick={() => setOutput({ title: 'Pre-migration checklist', text: checklist(selected, ['Discovery', 'Planning', 'Source preparation', 'Destination preparation', 'Identity preparation', 'Domain and DNS preparation', 'License preparation', 'BitTitan setup', 'Pilot migration', 'Pre-stage migration'], 'PRE-MIGRATION CHECKLIST') })}>Pre-migration checklist</Button>
              <Button variant="secondary" onClick={() => setOutput({ title: 'Cutover checklist', text: checklist(selected, ['Cutover', 'Final delta'], 'CUTOVER CHECKLIST') })}>Cutover checklist</Button>
              <Button variant="secondary" onClick={() => setOutput({ title: 'Post-migration checklist', text: checklist(selected, ['Post-migration validation', 'User support', 'Closure'], 'POST-MIGRATION CHECKLIST') })}>Post-migration checklist</Button>
              <Button variant="secondary" onClick={() => setOutput({ title: 'Rollback checklist', text: rollback(selected) })}>Rollback checklist</Button>
            </div>
          </Card>

          {output && <TextOutput title={output.title} text={output.text} />}

          {/* Stages */}
          <div className="space-y-2">
            {selected.stages.map((s, i) => {
              const pct = s.tasks.length ? Math.round((s.tasks.filter((t) => t.done).length / s.tasks.length) * 100) : 0;
              const open = openStage === s.id;
              return (
                <Card key={s.id} className="overflow-hidden">
                  <button onClick={() => setOpenStage(open ? null : s.id)} className="flex w-full items-center gap-3 p-4 text-left">
                    {open ? <ChevronDown size={16} className="text-slate-400" /> : <ChevronRight size={16} className="text-slate-400" />}
                    <span className="flex h-7 w-7 shrink-0 items-center justify-center rounded-full bg-blue-100 dark:bg-blue-900/50 text-xs font-bold text-blue-700 dark:text-blue-300">{i + 1}</span>
                    <span className="flex-1 font-semibold text-slate-700 dark:text-slate-200">{s.name}</span>
                    <Badge color={riskColor(s.risk)}>{s.risk} risk</Badge>
                    <Badge color={s.status === 'done' ? 'green' : s.status === 'in-progress' ? 'blue' : s.status === 'blocked' ? 'red' : 'gray'}>{s.status}</Badge>
                    <div className="w-28 hidden sm:block"><ProgressBar value={pct} color={pct === 100 ? 'bg-emerald-500' : 'bg-blue-600'} /></div>
                    <span className="w-10 text-right text-xs font-bold text-slate-500">{pct}%</span>
                  </button>
                  {open && (
                    <div className="border-t border-slate-200 dark:border-slate-700 p-4 space-y-3">
                      <div className="grid gap-3 sm:grid-cols-4">
                        <Select label="Status" value={s.status} onChange={(v) => updateStage(s.id, (st) => ({ ...st, status: v as MigrationStage['status'] }))}
                          options={['not-started', 'in-progress', 'blocked', 'done'].map((x) => ({ value: x, label: x }))} />
                        <Field label="Owner" value={s.owner} onChange={(v) => updateStage(s.id, (st) => ({ ...st, owner: v }))} />
                        <Select label="Risk" value={s.risk} onChange={(v) => updateStage(s.id, (st) => ({ ...st, risk: v as MigrationStage['risk'] }))}
                          options={['low', 'medium', 'high'].map((x) => ({ value: x, label: x }))} />
                        <Field label="Due date" type="date" value={s.dueDate} onChange={(v) => updateStage(s.id, (st) => ({ ...st, dueDate: v }))} />
                      </div>
                      <div className="space-y-1">
                        {s.tasks.map((t) => (
                          <label key={t.id} className="flex cursor-pointer items-start gap-2 rounded px-2 py-1 hover:bg-slate-50 dark:hover:bg-slate-700/40">
                            <input type="checkbox" checked={t.done} onChange={() => updateStage(s.id, (st) => ({ ...st, tasks: st.tasks.map((x) => (x.id === t.id ? { ...x, done: !x.done } : x)) }))}
                              className="mt-0.5 h-4 w-4 rounded border-slate-300 text-blue-600" />
                            <span className={`text-sm ${t.done ? 'text-slate-400 line-through' : 'text-slate-700 dark:text-slate-200'}`}>{t.label}</span>
                          </label>
                        ))}
                      </div>
                      <TextArea label="Stage notes" value={s.notes} onChange={(v) => updateStage(s.id, (st) => ({ ...st, notes: v }))} rows={2} />
                    </div>
                  )}
                </Card>
              );
            })}
          </div>

          {/* Manual runbook for the chosen type */}
          <Card className="p-5">
            <h3 className="mb-3 text-sm font-semibold uppercase tracking-wide text-slate-500 dark:text-slate-400">Manual runbook — {migrationTypeLabels[selected.type]}</h3>
            <div className="space-y-4">
              {guides[selected.type].map((g) => (
                <div key={g.title}>
                  <h4 className="mb-1.5 font-semibold text-slate-700 dark:text-slate-200">{g.title}</h4>
                  <ol className="space-y-1">
                    {g.steps.map((st, i) => (
                      <li key={i} className="flex gap-2 text-sm text-slate-600 dark:text-slate-300">
                        <span className="font-bold text-blue-500">{i + 1}.</span>{st}
                      </li>
                    ))}
                  </ol>
                </div>
              ))}
            </div>
          </Card>

          <AIHelper
            context={generatePlan(selected)}
            defaultPrompt="Review this migration plan: identify missing steps, risks for the cutover date, and write the user communication email for the announcement."
          />
        </div>
      )}
    </div>
  );
}
