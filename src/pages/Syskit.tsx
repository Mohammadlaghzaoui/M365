import { useState } from 'react';
import { Workflow } from 'lucide-react';
import { Badge, Card, CopyButton, PageHeader, Section, riskColor } from '../components/ui';
import { syskitTasks } from '../data/syskit';
import { AIHelper } from '../components/AIHelper';

export default function Syskit() {
  const [selected, setSelected] = useState(syskitTasks[0].id);
  const task = syskitTasks.find((t) => t.id === selected)!;

  return (
    <div>
      <PageHeader title="Syskit Assistant" subtitle="Governance and reporting playbook: which Syskit report to run, what to look for, and what to do with the findings." icon={<Workflow size={20} />} />
      <div className="grid gap-6 lg:grid-cols-[300px_1fr]">
        <div className="space-y-1.5 lg:sticky lg:top-20 lg:self-start">
          {syskitTasks.map((t) => (
            <button key={t.id} onClick={() => setSelected(t.id)}
              className={`flex w-full items-center justify-between gap-2 rounded-lg px-3 py-2.5 text-left text-sm font-medium ${selected === t.id ? 'bg-blue-600 text-white' : 'border border-slate-200 dark:border-slate-700 text-slate-600 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-700'}`}>
              {t.name}
              <Badge color={selected === t.id ? 'gray' : riskColor(t.risk)}>{t.risk}</Badge>
            </button>
          ))}
        </div>
        <div className="space-y-4">
          <Card className="p-5">
            <div className="mb-3 flex items-center justify-between">
              <h2 className="text-lg font-bold text-slate-800 dark:text-slate-100">{task.name}</h2>
              <Badge color={riskColor(task.risk)}>{task.risk} risk</Badge>
            </div>
            <Section title="Why run this report"><p className="text-sm text-slate-700 dark:text-slate-200">{task.why}</p></Section>
            <Section title="Where to find it in Syskit"><p className="text-sm font-medium text-blue-600 dark:text-blue-400">{task.where}</p></Section>
            <Section title="What to look for"><p className="text-sm text-slate-700 dark:text-slate-200">{task.lookFor}</p></Section>
            <Section title="Recommended action"><p className="text-sm text-slate-700 dark:text-slate-200">{task.action}</p></Section>
            <Section title="Escalation trigger"><p className="text-sm text-red-600 dark:text-red-300">{task.escalationTrigger}</p></Section>
            <Section title="Ticket note" action={<CopyButton text={task.ticketNote} />}>
              <pre className="whitespace-pre-wrap rounded-lg bg-slate-50 dark:bg-slate-900 p-3 text-sm text-slate-600 dark:text-slate-300 font-sans">{task.ticketNote}</pre>
            </Section>
          </Card>
          <AIHelper context={`Syskit governance task: ${task.name}\nPurpose: ${task.why}\nLooking for: ${task.lookFor}`} defaultPrompt="I ran this Syskit report — help me interpret findings and draft the customer summary. I will paste the findings." />
        </div>
      </div>
    </div>
  );
}
