import { useState } from 'react';
import { Sparkles, Loader2 } from 'lucide-react';
import { chat } from '../services/ai';
import { aiEnabled } from '../store/settings';
import { Button, Card, CopyButton } from './ui';
import { Link } from 'react-router-dom';

/**
 * Reusable "Ask AI" panel. Pass a context string (workflow, error, ticket...)
 * and it sends it to the configured provider with a task-specific instruction.
 */
export function AIHelper({ context, defaultPrompt = 'Explain this issue in simple language and give me extra troubleshooting steps.' }: { context: string; defaultPrompt?: string }) {
  const [prompt, setPrompt] = useState(defaultPrompt);
  const [reply, setReply] = useState('');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');

  if (!aiEnabled()) {
    return (
      <Card className="p-4 border-violet-200 dark:border-violet-800">
        <div className="flex items-center gap-2 text-sm text-slate-500 dark:text-slate-400">
          <Sparkles size={16} className="text-violet-500" />
          AI assistance is disabled. Enable it in{' '}
          <Link to="/settings" className="font-semibold text-violet-600 dark:text-violet-400 hover:underline">Settings</Link>
          {' '}— the portal works fully without it using built-in templates.
        </div>
      </Card>
    );
  }

  const ask = async () => {
    setLoading(true);
    setError('');
    try {
      const res = await chat([
        { role: 'user', content: `${prompt}\n\n--- CONTEXT ---\n${context}` },
      ]);
      setReply(res);
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setLoading(false);
    }
  };

  return (
    <Card className="p-4 border-violet-200 dark:border-violet-800">
      <div className="mb-2 flex items-center gap-2 text-sm font-semibold text-violet-700 dark:text-violet-300">
        <Sparkles size={16} /> Ask AI about this
      </div>
      <div className="flex gap-2">
        <input
          value={prompt}
          onChange={(e) => setPrompt(e.target.value)}
          onKeyDown={(e) => e.key === 'Enter' && !loading && ask()}
          className="flex-1 rounded-lg border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-900 px-3 py-2 text-sm text-slate-800 dark:text-slate-100 focus:border-violet-500 focus:outline-none"
        />
        <Button variant="ai" onClick={ask} disabled={loading}>
          {loading ? <Loader2 size={16} className="animate-spin" /> : <Sparkles size={16} />}
          Ask
        </Button>
      </div>
      {error && <p className="mt-2 text-sm text-red-500">{error}</p>}
      {reply && (
        <div className="mt-3 rounded-lg bg-violet-50 dark:bg-violet-900/20 p-3">
          <div className="mb-1 flex justify-end"><CopyButton text={reply} /></div>
          <pre className="whitespace-pre-wrap text-sm text-slate-700 dark:text-slate-200 font-sans leading-relaxed">{reply}</pre>
        </div>
      )}
    </Card>
  );
}
