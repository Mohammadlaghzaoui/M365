import { useState } from 'react';
import { Settings } from 'lucide-react';
import { Badge, Card, CodeBlock, CopyButton, PageHeader, Section, riskColor } from '../components/ui';
import { tenantTasks } from '../data/tenantAdmin';
import { AIHelper } from '../components/AIHelper';

export default function TenantAdmin() {
  const [selected, setSelected] = useState(tenantTasks[0].id);
  const task = tenantTasks.find((t) => t.id === selected)!;

  return (
    <div>
      <PageHeader title="Tenant Administration Assistant" subtitle="Standard tenant admin procedures with portal paths, roles, PowerShell, risk level and approval requirements." icon={<Settings size={20} />} />
      <div className="grid gap-6 lg:grid-cols-[280px_1fr]">
        <div className="space-y-1.5 lg:sticky lg:top-20 lg:self-start lg:max-h-[calc(100vh-6rem)] lg:overflow-y-auto pr-1">
          {tenantTasks.map((t) => (
            <button key={t.id} onClick={() => setSelected(t.id)}
              className={`w-full rounded-lg px-3 py-2.5 text-left text-sm font-medium ${selected === t.id ? 'bg-blue-600 text-white' : 'border border-slate-200 dark:border-slate-700 text-slate-600 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-700'}`}>
              {t.name}
            </button>
          ))}
        </div>
        <div className="space-y-4">
          <Card className="p-5">
            <div className="mb-3 flex flex-wrap items-center gap-2">
              <h2 className="text-lg font-bold text-slate-800 dark:text-slate-100 flex-1">{task.name}</h2>
              <Badge color={riskColor(task.risk)}>{task.risk} risk</Badge>
              <Badge color={task.approvalNeeded ? 'orange' : 'green'}>{task.approvalNeeded ? 'Approval required' : 'No approval needed'}</Badge>
            </div>
            <Section title="Portal path"><p className="text-sm font-medium text-blue-600 dark:text-blue-400">{task.portalPath}</p></Section>
            <Section title="Required role"><p className="text-sm text-slate-700 dark:text-slate-200">{task.requiredRole}</p></Section>
            <Section title="Step-by-step actions">
              <ol className="space-y-2">
                {task.steps.map((s, i) => (
                  <li key={i} className="flex items-start gap-3 text-sm text-slate-700 dark:text-slate-200">
                    <span className="flex h-6 w-6 shrink-0 items-center justify-center rounded-full bg-blue-100 dark:bg-blue-900/50 text-xs font-bold text-blue-700 dark:text-blue-300">{i + 1}</span>
                    <span className="pt-0.5">{s}</span>
                  </li>
                ))}
              </ol>
            </Section>
            {task.powershell.length > 0 && (
              <Section title="PowerShell">
                {task.powershell.map((ps, i) => (
                  <div key={i} className="mb-2">
                    <div className="mb-1 text-sm font-medium text-slate-600 dark:text-slate-300">{ps.label}</div>
                    <CodeBlock code={ps.command} />
                  </div>
                ))}
              </Section>
            )}
            <Section title="Ticket note" action={<CopyButton text={task.ticketNote} />}>
              <pre className="whitespace-pre-wrap rounded-lg bg-slate-50 dark:bg-slate-900 p-3 text-sm text-slate-600 dark:text-slate-300 font-sans">{task.ticketNote}</pre>
            </Section>
          </Card>
          <AIHelper context={`Tenant admin task: ${task.name}\nSteps:\n${task.steps.join('\n')}`} />
        </div>
      </div>
    </div>
  );
}
