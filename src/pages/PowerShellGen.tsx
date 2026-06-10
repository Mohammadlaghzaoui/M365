import { useMemo, useState } from 'react';
import { Terminal, Star, AlertTriangle } from 'lucide-react';
import { Card, CodeBlock, PageHeader, Section, Select } from '../components/ui';
import { psTasks, psServices } from '../data/psTasks';
import { useLocalStorage } from '../store/useLocalStorage';
import { AIHelper } from '../components/AIHelper';

export default function PowerShellGen() {
  const [service, setService] = useState(psServices[0]);
  const tasks = useMemo(() => psTasks.filter((t) => t.service === service), [service]);
  const [taskId, setTaskId] = useState(tasks[0]?.id ?? '');
  const task = psTasks.find((t) => t.id === taskId) ?? tasks[0];
  const [params, setParams] = useState<Record<string, string>>({});
  const [favorites, setFavorites] = useLocalStorage<string[]>('favorite-ps', []);

  const changeService = (s: string) => {
    setService(s);
    const first = psTasks.find((t) => t.service === s);
    setTaskId(first?.id ?? '');
    setParams({});
  };

  const command = useMemo(() => {
    if (!task) return '';
    let cmd = task.template;
    for (const p of task.params) {
      const v = params[p.key]?.trim();
      cmd = cmd.split(`{${p.key}}`).join(v || `<${p.label}>`);
    }
    return cmd;
  }, [task, params]);

  const toggleFav = () => {
    if (!task) return;
    setFavorites((prev) => (prev.includes(task.id) ? prev.filter((f) => f !== task.id) : [...prev, task.id]));
  };

  if (!task) return null;

  return (
    <div>
      <PageHeader title="PowerShell Command Generator" subtitle="Pick a service and task, fill in the parameters — get the command with explanation, module, role and warnings." icon={<Terminal size={20} />} />
      <div className="grid gap-6 lg:grid-cols-[360px_1fr]">
        <Card className="p-5 self-start space-y-3">
          <Select label="Service" value={service} onChange={changeService} options={psServices.map((s) => ({ value: s, label: s }))} />
          <Select label="Task" value={task.id} onChange={(v) => { setTaskId(v); setParams({}); }} options={tasks.map((t) => ({ value: t.id, label: t.name }))} />
          {task.params.map((p) => (
            <label key={p.key} className="block">
              <span className="mb-1 block text-sm font-medium text-slate-700 dark:text-slate-300">{p.label}</span>
              <input value={params[p.key] ?? ''} onChange={(e) => setParams({ ...params, [p.key]: e.target.value })} placeholder={p.placeholder}
                className="w-full rounded-lg border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-900 px-3 py-2 text-sm text-slate-800 dark:text-slate-100 focus:border-blue-500 focus:outline-none" />
            </label>
          ))}
          <button onClick={toggleFav} className="flex items-center gap-2 text-sm font-medium text-slate-600 dark:text-slate-300 hover:text-amber-500">
            <Star size={16} className={favorites.includes(task.id) ? 'fill-amber-400 text-amber-400' : ''} />
            {favorites.includes(task.id) ? 'Remove from favorites' : 'Add to dashboard favorites'}
          </button>
        </Card>

        <div className="space-y-4">
          <Card className="p-5">
            <Section title="Generated command">
              <CodeBlock code={command} />
            </Section>
            <Section title="Explanation"><p className="text-sm text-slate-700 dark:text-slate-200">{task.explanation}</p></Section>
            <div className="grid gap-4 sm:grid-cols-2">
              <div>
                <div className="text-xs font-semibold uppercase text-slate-400">Required module</div>
                <p className="mt-1 text-sm font-medium text-slate-700 dark:text-slate-200">{task.module}</p>
              </div>
              <div>
                <div className="text-xs font-semibold uppercase text-slate-400">Required admin role</div>
                <p className="mt-1 text-sm font-medium text-slate-700 dark:text-slate-200">{task.requiredRole}</p>
              </div>
            </div>
            <div className="mt-4 flex items-start gap-2 rounded-lg border border-amber-200 dark:border-amber-800 bg-amber-50 dark:bg-amber-900/20 p-3 text-sm text-amber-700 dark:text-amber-300">
              <AlertTriangle size={16} className="mt-0.5 shrink-0" />{task.warning}
            </div>
          </Card>
          <AIHelper context={`PowerShell command:\n${command}\nPurpose: ${task.explanation}`} defaultPrompt="Explain this command parameter by parameter and what the output will look like." />
        </div>
      </div>
    </div>
  );
}
