import { useState } from 'react';
import { Cloud } from 'lucide-react';
import { Card, CodeBlock, PageHeader, Section, TextOutput } from '../components/ui';
import { btGuide, btChecklists, btErrors } from '../data/bittitan';
import { AIHelper } from '../components/AIHelper';

type Tab = 'guide' | 'checklists' | 'errors';

export default function BitTitan() {
  const [tab, setTab] = useState<Tab>('guide');
  const [openStep, setOpenStep] = useState<string>(btGuide[0].id);
  const [selectedError, setSelectedError] = useState(btErrors[0].id);
  const err = btErrors.find((e) => e.id === selectedError)!;

  return (
    <div>
      <PageHeader title="BitTitan MigrationWiz Assistant" subtitle="Guided MigrationWiz process for M365↔M365, Google Workspace→M365 and Exchange on-prem→M365, with checklists and error helper." icon={<Cloud size={20} />} />

      <div className="mb-5 flex flex-wrap gap-2">
        {([['guide', 'Step-by-step guide'], ['checklists', 'Checklists'], ['errors', 'Error helper']] as [Tab, string][]).map(([id, label]) => (
          <button key={id} onClick={() => setTab(id)}
            className={`rounded-lg px-4 py-2 text-sm font-semibold ${tab === id ? 'bg-blue-600 text-white' : 'border border-slate-300 dark:border-slate-600 text-slate-600 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-700'}`}>
            {label}
          </button>
        ))}
      </div>

      {tab === 'guide' && (
        <div className="space-y-2">
          {btGuide.map((step) => (
            <Card key={step.id} className="overflow-hidden">
              <button onClick={() => setOpenStep(openStep === step.id ? '' : step.id)} className="w-full p-4 text-left font-semibold text-slate-700 dark:text-slate-200">
                {step.title}
              </button>
              {openStep === step.id && (
                <div className="border-t border-slate-200 dark:border-slate-700 p-4">
                  <ul className="space-y-2">
                    {step.details.map((d, i) => (
                      <li key={i} className="flex items-start gap-2 text-sm text-slate-600 dark:text-slate-300">
                        <span className="mt-1.5 h-1.5 w-1.5 shrink-0 rounded-full bg-blue-500" />{d}
                      </li>
                    ))}
                  </ul>
                </div>
              )}
            </Card>
          ))}
          <AIHelper context={`BitTitan MigrationWiz project. Current guide:\n${btGuide.map((g) => g.title).join('\n')}`} defaultPrompt="Create a MigrationWiz project plan for a 120-user Google Workspace to M365 migration with a cutover in 4 weeks." />
        </div>
      )}

      {tab === 'checklists' && (
        <div className="grid gap-4 lg:grid-cols-3">
          <TextOutput title="Permission checklist" text={btChecklists.permissions.map((c) => `[ ] ${c}`).join('\n')} />
          <TextOutput title="Cutover checklist" text={btChecklists.cutover.map((c) => `[ ] ${c}`).join('\n')} />
          <TextOutput title="Post-migration checklist" text={btChecklists.postMigration.map((c) => `[ ] ${c}`).join('\n')} />
        </div>
      )}

      {tab === 'errors' && (
        <div className="grid gap-6 lg:grid-cols-[320px_1fr]">
          <div className="space-y-1.5 lg:sticky lg:top-20 lg:self-start">
            {btErrors.map((e) => (
              <button key={e.id} onClick={() => setSelectedError(e.id)}
                className={`w-full rounded-lg px-3 py-2.5 text-left text-xs font-medium ${selectedError === e.id ? 'bg-red-600 text-white' : 'border border-slate-200 dark:border-slate-700 text-slate-600 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-700'}`}>
                {e.name}
              </button>
            ))}
          </div>
          <div className="space-y-4">
            <Card className="p-5">
              <h3 className="text-sm font-bold text-red-600 dark:text-red-400">{err.name}</h3>
              <Section title="Meaning"><p className="text-sm text-slate-700 dark:text-slate-200">{err.meaning}</p></Section>
              <Section title="Likely cause"><p className="text-sm text-slate-700 dark:text-slate-200">{err.cause}</p></Section>
              <Section title="Fix steps">
                <ol className="space-y-2">
                  {err.fix.map((f, i) => (
                    <li key={i} className="flex items-start gap-2.5 text-sm text-slate-700 dark:text-slate-200">
                      <span className="flex h-5 w-5 shrink-0 items-center justify-center rounded-full bg-emerald-100 dark:bg-emerald-900/50 text-xs font-bold text-emerald-700 dark:text-emerald-300">{i + 1}</span>{f}
                    </li>
                  ))}
                </ol>
              </Section>
              {err.powershell && <Section title="PowerShell"><CodeBlock code={err.powershell} /></Section>}
              <Section title="Prevention"><p className="text-sm text-slate-700 dark:text-slate-200">{err.prevention}</p></Section>
              <Section title="Escalation note"><p className="text-sm text-amber-700 dark:text-amber-300">{err.escalationNote}</p></Section>
            </Card>
            <AIHelper context={`BitTitan MigrationWiz error: ${err.name}\n${err.meaning}\nCause: ${err.cause}`} defaultPrompt="Summarize this BitTitan error for the customer ticket and give me the fastest fix path." />
          </div>
        </div>
      )}
    </div>
  );
}
