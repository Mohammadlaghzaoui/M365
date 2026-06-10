import { useState } from 'react';
import { Settings as SettingsIcon, Plug, Loader2, CheckCircle2, XCircle, Trash2 } from 'lucide-react';
import { Button, Card, Field, PageHeader, Section, Select, TextArea } from '../components/ui';
import { AISettings, AIProvider } from '../types';
import { DEFAULT_AI_SETTINGS, DEFAULT_SYSTEM_PROMPT, getAISettings, saveAISettings } from '../store/settings';
import { testConnection } from '../services/ai';

const providerInfo: Record<AIProvider, { label: string; modelHint: string; keyHint: string }> = {
  openrouter: { label: 'OpenRouter (recommended first)', modelHint: 'e.g. anthropic/claude-sonnet-4.5, openai/gpt-4o, meta-llama/llama-3.3-70b-instruct', keyHint: 'sk-or-v1-… (openrouter.ai/keys)' },
  openai: { label: 'OpenAI API', modelHint: 'e.g. gpt-4o, gpt-4o-mini', keyHint: 'sk-… (platform.openai.com)' },
  claude: { label: 'Claude API (Anthropic)', modelHint: 'e.g. claude-sonnet-4-5, claude-haiku-4-5', keyHint: 'sk-ant-… (console.anthropic.com)' },
  disabled: { label: 'Disabled (templates only)', modelHint: '', keyHint: '' },
};

export default function Settings() {
  const [s, setS] = useState<AISettings>(getAISettings());
  const [testState, setTestState] = useState<'idle' | 'testing' | 'ok' | 'fail'>('idle');
  const [testMsg, setTestMsg] = useState('');
  const [saved, setSaved] = useState(false);

  const set = <K extends keyof AISettings>(k: K, v: AISettings[K]) => {
    setS((prev) => ({ ...prev, [k]: v }));
    setSaved(false);
  };

  const save = () => {
    saveAISettings(s);
    setSaved(true);
    setTimeout(() => setSaved(false), 2000);
  };

  const test = async () => {
    saveAISettings(s);
    setTestState('testing');
    setTestMsg('');
    try {
      const reply = await testConnection(s);
      setTestState('ok');
      setTestMsg(`Model replied: "${reply.slice(0, 120)}"`);
    } catch (e) {
      setTestState('fail');
      setTestMsg(e instanceof Error ? e.message : String(e));
    }
  };

  const clearData = () => {
    if (confirm('Clear ALL local portal data (tickets, projects, notes, KB customs, settings)? This cannot be undone.')) {
      Object.keys(localStorage).filter((k) => k.startsWith('workpilot:')).forEach((k) => localStorage.removeItem(k));
      location.reload();
    }
  };

  const info = providerInfo[s.provider];

  return (
    <div className="max-w-3xl">
      <PageHeader title="Settings" subtitle="AI provider configuration and local data management. Keys are stored only in your browser's localStorage." icon={<SettingsIcon size={20} />} />

      <Card className="p-5 space-y-4">
        <Section title="AI provider">
          <Select label="Provider" value={s.provider} onChange={(v) => set('provider', v as AIProvider)}
            options={(Object.keys(providerInfo) as AIProvider[]).map((p) => ({ value: p, label: providerInfo[p].label }))} />
          <p className="mt-2 text-xs text-slate-400">
            Switching provider only changes the endpoint and auth header — your prompts and workflows stay identical. Start with OpenRouter; move to OpenAI or Claude API by selecting it here and pasting that key.
          </p>
        </Section>

        {s.provider !== 'disabled' && (
          <>
            <Field label={`API key ${info.keyHint && `(${info.keyHint})`}`} type="password" value={s.apiKey} onChange={(v) => set('apiKey', v)} placeholder="Paste your API key" />
            <Field label={`Model name ${info.modelHint && `— ${info.modelHint}`}`} value={s.model} onChange={(v) => set('model', v)} />
            <label className="block">
              <span className="mb-1 block text-sm font-medium text-slate-700 dark:text-slate-300">Temperature: {s.temperature.toFixed(1)} (lower = more precise)</span>
              <input type="range" min={0} max={1} step={0.1} value={s.temperature} onChange={(e) => set('temperature', Number(e.target.value))} className="w-full accent-violet-600" />
            </label>
            <TextArea label="System prompt" value={s.systemPrompt} onChange={(v) => set('systemPrompt', v)} rows={6} />
            <button onClick={() => set('systemPrompt', DEFAULT_SYSTEM_PROMPT)} className="text-xs font-semibold text-blue-500 hover:underline">Reset to default system prompt</button>
          </>
        )}

        <div className="flex flex-wrap items-center gap-3 border-t border-slate-200 dark:border-slate-700 pt-4">
          <Button onClick={save}>{saved ? 'Saved ✓' : 'Save settings'}</Button>
          {s.provider !== 'disabled' && (
            <Button variant="ai" onClick={test} disabled={testState === 'testing' || !s.apiKey}>
              {testState === 'testing' ? <Loader2 size={16} className="animate-spin" /> : <Plug size={16} />} Test AI connection
            </Button>
          )}
          {testState === 'ok' && <span className="flex items-center gap-1.5 text-sm text-emerald-600"><CheckCircle2 size={16} /> Connected</span>}
          {testState === 'fail' && <span className="flex items-center gap-1.5 text-sm text-red-500"><XCircle size={16} /> Failed</span>}
        </div>
        {testMsg && <p className={`text-xs ${testState === 'ok' ? 'text-emerald-600' : 'text-red-500'}`}>{testMsg}</p>}
        {s.provider === 'claude' && (
          <p className="text-xs text-amber-600 dark:text-amber-400">Note: the Claude API is called directly from the browser (CORS-enabled via the anthropic-dangerous-direct-browser-access header). For production team use, route keys through a small backend proxy instead of the browser.</p>
        )}
      </Card>

      <Card className="mt-5 p-5">
        <Section title="Local data">
          <p className="mb-3 text-sm text-slate-500 dark:text-slate-400">
            All portal data (tickets, migration projects, notes, KB articles, security status, settings) lives in your browser's localStorage under the <code className="rounded bg-slate-100 dark:bg-slate-700 px-1">workpilot:</code> prefix. The storage layer is a single module (src/store) so it can be swapped for SQLite/PostgreSQL later without touching the pages.
          </p>
          <Button variant="danger" onClick={clearData}><Trash2 size={15} /> Clear all local data</Button>
        </Section>
      </Card>
    </div>
  );
}
