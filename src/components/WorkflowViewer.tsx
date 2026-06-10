import { useEffect, useState } from 'react';
import { Workflow } from '../types';
import { AlertTriangle, ArrowUpRight, CheckCircle2, HelpCircle, ListChecks, MapPin, Shield, Terminal } from 'lucide-react';
import { Badge, Card, CodeBlock, CopyButton, Section } from './ui';
import { AIHelper } from './AIHelper';
import { load, save } from '../store/useLocalStorage';

function trackRecent(title: string) {
  const recent = load<string[]>('recent-workflows', []);
  const next = [title, ...recent.filter((r) => r !== title)].slice(0, 8);
  save('recent-workflows', next);
}

export function WorkflowViewer({ workflow }: { workflow: Workflow }) {
  useEffect(() => trackRecent(workflow.title), [workflow.title]);

  return (
    <div className="space-y-5">
      <Card className="p-5">
        <div className="flex items-start justify-between gap-3">
          <div>
            <h2 className="text-lg font-bold text-slate-800 dark:text-slate-100">{workflow.title}</h2>
            <p className="mt-1 text-sm text-slate-600 dark:text-slate-300">{workflow.explanation}</p>
          </div>
          <Badge color="blue">{workflow.requiredRole}</Badge>
        </div>
      </Card>

      <div className="grid gap-5 lg:grid-cols-2">
        <Card className="p-5">
          <Section title="Questions to ask the user" action={<CopyButton text={workflow.questions.map((q) => `- ${q}`).join('\n')} />}>
            <ul className="space-y-2">
              {workflow.questions.map((q, i) => (
                <li key={i} className="flex items-start gap-2 text-sm text-slate-700 dark:text-slate-200">
                  <HelpCircle size={15} className="mt-0.5 shrink-0 text-blue-500" />{q}
                </li>
              ))}
            </ul>
          </Section>
        </Card>

        <Card className="p-5">
          <Section title="Portal path">
            <ul className="space-y-2">
              {workflow.portalPaths.map((p, i) => (
                <li key={i} className="flex items-start gap-2 text-sm font-medium text-slate-700 dark:text-slate-200">
                  <MapPin size={15} className="mt-0.5 shrink-0 text-blue-500" />{p}
                </li>
              ))}
            </ul>
          </Section>
          <Section title="Required admin role">
            <div className="flex items-center gap-2 text-sm text-slate-700 dark:text-slate-200">
              <Shield size={15} className="text-blue-500" />{workflow.requiredRole}
            </div>
          </Section>
        </Card>
      </div>

      <Card className="p-5">
        <Section title="Step-by-step checks" action={<CopyButton text={workflow.steps.map((s, i) => `${i + 1}. ${s}`).join('\n')} />}>
          <ol className="space-y-2.5">
            {workflow.steps.map((s, i) => (
              <li key={i} className="flex items-start gap-3 text-sm text-slate-700 dark:text-slate-200">
                <span className="flex h-6 w-6 shrink-0 items-center justify-center rounded-full bg-blue-100 dark:bg-blue-900/50 text-xs font-bold text-blue-700 dark:text-blue-300">{i + 1}</span>
                <span className="pt-0.5">{s}</span>
              </li>
            ))}
          </ol>
        </Section>
      </Card>

      {workflow.powershell.length > 0 && (
        <Card className="p-5">
          <Section title="PowerShell commands">
            <div className="space-y-3">
              {workflow.powershell.map((ps, i) => (
                <div key={i}>
                  <div className="mb-1 flex items-center gap-2 text-sm font-medium text-slate-600 dark:text-slate-300">
                    <Terminal size={14} className="text-violet-500" />{ps.label}
                  </div>
                  <CodeBlock code={ps.command} />
                </div>
              ))}
            </div>
          </Section>
        </Card>
      )}

      <div className="grid gap-5 lg:grid-cols-2">
        <Card className="p-5 border-amber-200 dark:border-amber-800">
          <Section title="Dangerous — do not change without approval">
            <ul className="space-y-2">
              {workflow.dangers.map((d, i) => (
                <li key={i} className="flex items-start gap-2 text-sm text-amber-700 dark:text-amber-300">
                  <AlertTriangle size={15} className="mt-0.5 shrink-0" />{d}
                </li>
              ))}
            </ul>
          </Section>
        </Card>

        <Card className="p-5 border-red-200 dark:border-red-800">
          <Section title="Escalate when">
            <ul className="space-y-2">
              {workflow.escalation.map((e, i) => (
                <li key={i} className="flex items-start gap-2 text-sm text-red-600 dark:text-red-300">
                  <ArrowUpRight size={15} className="mt-0.5 shrink-0" />{e}
                </li>
              ))}
            </ul>
          </Section>
        </Card>
      </div>

      <Card className="p-5">
        <Section title="Ticket update template" action={<CopyButton text={workflow.ticketTemplate} />}>
          <pre className="whitespace-pre-wrap rounded-lg bg-slate-50 dark:bg-slate-900 p-3 text-sm text-slate-600 dark:text-slate-300 font-sans leading-relaxed">{workflow.ticketTemplate}</pre>
        </Section>
      </Card>

      <AIHelper context={`Workflow: ${workflow.title}\n${workflow.explanation}\nSteps:\n${workflow.steps.join('\n')}`} />
    </div>
  );
}

export function WorkflowList({ workflows, selectedId, onSelect }: { workflows: Workflow[]; selectedId: string | null; onSelect: (id: string) => void }) {
  const [filter, setFilter] = useState('');
  const filtered = workflows.filter((w) => w.title.toLowerCase().includes(filter.toLowerCase()));
  return (
    <div>
      <input
        value={filter}
        onChange={(e) => setFilter(e.target.value)}
        placeholder="Filter workflows..."
        className="mb-3 w-full rounded-lg border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-900 px-3 py-2 text-sm text-slate-800 dark:text-slate-100 focus:border-blue-500 focus:outline-none"
      />
      <div className="space-y-1.5">
        {filtered.map((w) => (
          <button
            key={w.id}
            onClick={() => onSelect(w.id)}
            className={`flex w-full items-center gap-2 rounded-lg px-3 py-2.5 text-left text-sm font-medium transition-colors ${
              selectedId === w.id
                ? 'bg-blue-600 text-white'
                : 'text-slate-700 dark:text-slate-200 hover:bg-slate-100 dark:hover:bg-slate-700'
            }`}
          >
            {selectedId === w.id ? <CheckCircle2 size={15} className="shrink-0" /> : <ListChecks size={15} className="shrink-0 text-slate-400" />}
            {w.title}
          </button>
        ))}
      </div>
    </div>
  );
}
