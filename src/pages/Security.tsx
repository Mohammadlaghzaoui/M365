import { useState } from 'react';
import { ShieldCheck } from 'lucide-react';
import { Badge, Card, PageHeader, ProgressBar, TextOutput, riskColor } from '../components/ui';
import { securityControls } from '../data/security';
import { SecurityControlState } from '../types';
import { useLocalStorage } from '../store/useLocalStorage';
import { AIHelper } from '../components/AIHelper';

const defaultState: SecurityControlState = { status: 'not-reviewed', evidence: '', notes: '' };

const statusOptions = ['not-reviewed', 'compliant', 'partial', 'non-compliant', 'n/a'] as const;

export default function Security() {
  const [state, setState] = useLocalStorage<Record<string, SecurityControlState>>('security-state', {});
  const [filter, setFilter] = useState<'all' | 'quickwins' | '30' | '90' | 'open'>('all');
  const [openId, setOpenId] = useState<string | null>(null);

  const get = (id: string) => state[id] ?? defaultState;
  const update = (id: string, patch: Partial<SecurityControlState>) =>
    setState((prev) => ({ ...prev, [id]: { ...(prev[id] ?? defaultState), ...patch } }));

  const filtered = securityControls.filter((c) => {
    if (filter === 'quickwins') return c.quickWin;
    if (filter === '30') return c.plan === '30';
    if (filter === '90') return c.plan === '90';
    if (filter === 'open') return get(c.id).status !== 'compliant' && get(c.id).status !== 'n/a';
    return true;
  });

  const compliant = securityControls.filter((c) => get(c.id).status === 'compliant').length;
  const pct = Math.round((compliant / securityControls.length) * 100);

  const quickWinsText = securityControls.filter((c) => c.quickWin).map((c) => `[${get(c.id).status === 'compliant' ? 'x' : ' '}] ${c.name} — ${c.recommended}`).join('\n');
  const plan30 = securityControls.filter((c) => c.plan === '30').map((c) => `[${get(c.id).status === 'compliant' ? 'x' : ' '}] ${c.name} (${c.priority}) — ${c.portalPath}`).join('\n');
  const plan90 = securityControls.filter((c) => c.plan === '90').map((c) => `[${get(c.id).status === 'compliant' ? 'x' : ' '}] ${c.name} (${c.priority}) — ${c.portalPath}`).join('\n');

  return (
    <div>
      <PageHeader title="Security Hardening Assistant" subtitle="Track 27 baseline controls per tenant: status, evidence, priorities, plus quick wins and 30/90-day plans." icon={<ShieldCheck size={20} />} />

      <Card className="p-5 mb-5">
        <div className="flex items-center gap-4">
          <div className="flex-1">
            <div className="mb-1 flex justify-between text-sm">
              <span className="font-semibold text-slate-700 dark:text-slate-200">Hardening progress</span>
              <span className="font-bold text-emerald-600">{compliant}/{securityControls.length} compliant ({pct}%)</span>
            </div>
            <ProgressBar value={pct} color="bg-emerald-500" />
          </div>
        </div>
      </Card>

      <div className="mb-5 flex flex-wrap gap-2">
        {([['all', 'All controls'], ['quickwins', 'Quick wins'], ['30', '30-day plan'], ['90', '90-day plan'], ['open', 'Open items']] as const).map(([id, label]) => (
          <button key={id} onClick={() => setFilter(id)}
            className={`rounded-lg px-4 py-2 text-sm font-semibold ${filter === id ? 'bg-blue-600 text-white' : 'border border-slate-300 dark:border-slate-600 text-slate-600 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-700'}`}>
            {label}
          </button>
        ))}
      </div>

      {filter === 'quickwins' && <div className="mb-4"><TextOutput title="Quick wins checklist (copy for the customer)" text={quickWinsText} /></div>}
      {filter === '30' && <div className="mb-4"><TextOutput title="30-day security plan" text={plan30} /></div>}
      {filter === '90' && <div className="mb-4"><TextOutput title="90-day security plan" text={plan90} /></div>}

      <div className="space-y-2">
        {filtered.map((c) => {
          const s = get(c.id);
          const open = openId === c.id;
          return (
            <Card key={c.id} className="overflow-hidden">
              <button onClick={() => setOpenId(open ? null : c.id)} className="flex w-full flex-wrap items-center gap-2 p-4 text-left">
                <span className="flex-1 font-semibold text-slate-700 dark:text-slate-200">{c.name}</span>
                {c.quickWin && <Badge color="green">quick win</Badge>}
                <Badge color={riskColor(c.priority)}>{c.priority}</Badge>
                <Badge color={s.status === 'compliant' ? 'green' : s.status === 'non-compliant' ? 'red' : s.status === 'partial' ? 'orange' : 'gray'}>{s.status}</Badge>
              </button>
              {open && (
                <div className="space-y-3 border-t border-slate-200 dark:border-slate-700 p-4">
                  <div className="grid gap-3 text-sm sm:grid-cols-2">
                    <div><span className="font-semibold text-slate-500 dark:text-slate-400">Risk if not configured: </span><span className="text-slate-700 dark:text-slate-200">{c.risk}</span></div>
                    <div><span className="font-semibold text-slate-500 dark:text-slate-400">Recommended: </span><span className="text-slate-700 dark:text-slate-200">{c.recommended}</span></div>
                    <div><span className="font-semibold text-slate-500 dark:text-slate-400">Portal: </span><span className="text-blue-600 dark:text-blue-400">{c.portalPath}</span></div>
                    <div><span className="font-semibold text-slate-500 dark:text-slate-400">Required role: </span><span className="text-slate-700 dark:text-slate-200">{c.requiredRole}</span></div>
                  </div>
                  <div className="grid gap-3 sm:grid-cols-3">
                    <label className="block">
                      <span className="mb-1 block text-sm font-medium text-slate-700 dark:text-slate-300">Status</span>
                      <select value={s.status} onChange={(e) => update(c.id, { status: e.target.value as SecurityControlState['status'] })}
                        className="w-full rounded-lg border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-900 px-3 py-2 text-sm">
                        {statusOptions.map((o) => <option key={o} value={o}>{o}</option>)}
                      </select>
                    </label>
                    <label className="block">
                      <span className="mb-1 block text-sm font-medium text-slate-700 dark:text-slate-300">Evidence</span>
                      <input value={s.evidence} onChange={(e) => update(c.id, { evidence: e.target.value })} placeholder="screenshot ref / export / date"
                        className="w-full rounded-lg border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-900 px-3 py-2 text-sm" />
                    </label>
                    <label className="block">
                      <span className="mb-1 block text-sm font-medium text-slate-700 dark:text-slate-300">Notes</span>
                      <input value={s.notes} onChange={(e) => update(c.id, { notes: e.target.value })}
                        className="w-full rounded-lg border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-900 px-3 py-2 text-sm" />
                    </label>
                  </div>
                </div>
              )}
            </Card>
          );
        })}
      </div>

      <div className="mt-6">
        <AIHelper
          context={`Security hardening status: ${compliant}/${securityControls.length} compliant.\nOpen items:\n${securityControls.filter((c) => get(c.id).status !== 'compliant').map((c) => `- ${c.name} (${c.priority}): ${get(c.id).status}`).join('\n')}`}
          defaultPrompt="Prioritize my open hardening items and write a customer-facing summary with business risk per item."
        />
      </div>
    </div>
  );
}
