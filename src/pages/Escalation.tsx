import { useState } from 'react';
import { Shield } from 'lucide-react';
import { Badge, Card, CopyButton, PageHeader, Section } from '../components/ui';
import { escalationMatrix } from '../data/escalation';

export default function Escalation() {
  const [selected, setSelected] = useState<string | null>(null);

  return (
    <div>
      <PageHeader title="Escalation Matrix" subtitle="When to escalate, with which evidence, to which team — including ready-to-copy example escalation notes." icon={<Shield size={20} />} />

      <Card className="overflow-x-auto">
        <table className="w-full text-sm">
          <thead>
            <tr className="border-b border-slate-200 dark:border-slate-700 text-left text-xs uppercase text-slate-400">
              <th className="p-3">Service</th>
              <th className="p-3">Issue type</th>
              <th className="p-3">When to escalate</th>
              <th className="p-3">Escalation team</th>
              <th className="p-3"></th>
            </tr>
          </thead>
          <tbody>
            {escalationMatrix.map((e) => (
              <tr key={e.id} className="border-b border-slate-100 dark:border-slate-700/50 align-top">
                <td className="p-3"><Badge color="blue">{e.service}</Badge></td>
                <td className="p-3 font-medium text-slate-700 dark:text-slate-200">{e.issueType}</td>
                <td className="p-3 max-w-md text-slate-500 dark:text-slate-300">{e.whenToEscalate}</td>
                <td className="p-3 text-slate-500 dark:text-slate-300">{e.team}</td>
                <td className="p-3">
                  <button onClick={() => setSelected(selected === e.id ? null : e.id)} className="text-xs font-semibold text-blue-500 hover:underline whitespace-nowrap">
                    {selected === e.id ? 'Hide details' : 'Details + note'}
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </Card>

      {selected && (() => {
        const e = escalationMatrix.find((x) => x.id === selected)!;
        return (
          <Card className="mt-4 p-5">
            <h3 className="mb-3 font-bold text-slate-800 dark:text-slate-100">{e.service} — {e.issueType}</h3>
            <Section title="Required evidence"><p className="text-sm text-slate-700 dark:text-slate-200">{e.evidence}</p></Section>
            <Section title="Example escalation note" action={<CopyButton text={e.exampleNote} />}>
              <pre className="whitespace-pre-wrap rounded-lg bg-slate-50 dark:bg-slate-900 p-3 text-sm text-slate-600 dark:text-slate-300 font-sans leading-relaxed">{e.exampleNote}</pre>
            </Section>
          </Card>
        );
      })()}
    </div>
  );
}
