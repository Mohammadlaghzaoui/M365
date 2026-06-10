import { useState } from 'react';
import { Gauge } from 'lucide-react';
import { Badge, Card, CopyButton, PageHeader, Section } from '../components/ui';
import { monitoringChecks } from '../data/monitoring';
import { AIHelper } from '../components/AIHelper';

export default function Monitoring() {
  const [selected, setSelected] = useState(monitoringChecks[0].id);
  const check = monitoringChecks.find((c) => c.id === selected)!;

  return (
    <div>
      <PageHeader title="Monitoring Assistant" subtitle="Daily and weekly health checks: what to look at, what is normal, what is risky, and when to escalate." icon={<Gauge size={20} />} />
      <div className="grid gap-6 lg:grid-cols-[300px_1fr]">
        <div className="space-y-1.5 lg:sticky lg:top-20 lg:self-start lg:max-h-[calc(100vh-6rem)] lg:overflow-y-auto pr-1">
          {monitoringChecks.map((c) => (
            <button key={c.id} onClick={() => setSelected(c.id)}
              className={`flex w-full items-center justify-between gap-2 rounded-lg px-3 py-2.5 text-left text-sm font-medium ${selected === c.id ? 'bg-blue-600 text-white' : 'border border-slate-200 dark:border-slate-700 text-slate-600 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-700'}`}>
              {c.name}
            </button>
          ))}
        </div>
        <div className="space-y-4">
          <Card className="p-5">
            <div className="mb-3 flex items-center justify-between">
              <h2 className="text-lg font-bold text-slate-800 dark:text-slate-100">{check.name}</h2>
              <Badge color="blue">{check.frequency}</Badge>
            </div>
            <Section title="Portal path"><p className="text-sm font-medium text-blue-600 dark:text-blue-400">{check.portalPath}</p></Section>
            <Section title="What to check"><p className="text-sm text-slate-700 dark:text-slate-200">{check.whatToCheck}</p></Section>
            <div className="grid gap-4 sm:grid-cols-2">
              <div className="rounded-lg border border-emerald-200 dark:border-emerald-800 bg-emerald-50 dark:bg-emerald-900/20 p-3">
                <div className="mb-1 text-xs font-bold uppercase text-emerald-600 dark:text-emerald-400">Normal</div>
                <p className="text-sm text-slate-700 dark:text-slate-200">{check.normal}</p>
              </div>
              <div className="rounded-lg border border-red-200 dark:border-red-800 bg-red-50 dark:bg-red-900/20 p-3">
                <div className="mb-1 text-xs font-bold uppercase text-red-600 dark:text-red-400">Risky</div>
                <p className="text-sm text-slate-700 dark:text-slate-200">{check.risky}</p>
              </div>
            </div>
            <div className="mt-4">
              <Section title="Escalation trigger"><p className="text-sm text-amber-700 dark:text-amber-300">{check.escalation}</p></Section>
              <Section title="Ticket note" action={<CopyButton text={check.ticketNote} />}>
                <pre className="whitespace-pre-wrap rounded-lg bg-slate-50 dark:bg-slate-900 p-3 text-sm text-slate-600 dark:text-slate-300 font-sans">{check.ticketNote}</pre>
              </Section>
            </div>
          </Card>
          <AIHelper context={`Monitoring check: ${check.name}\nLooking at: ${check.whatToCheck}\nRisky signs: ${check.risky}`} defaultPrompt="I found something during this check — help me triage it. I will paste what I see." />
        </div>
      </div>
    </div>
  );
}
