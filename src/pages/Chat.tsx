import { useRef, useState } from 'react';
import { Bot, Loader2, SendHorizonal, Trash2 } from 'lucide-react';
import { Link } from 'react-router-dom';
import { Card, CopyButton, PageHeader } from '../components/ui';
import { chat } from '../services/ai';
import { aiEnabled } from '../store/settings';
import { ChatMessage } from '../types';
import { useLocalStorage } from '../store/useLocalStorage';

const quickPrompts = [
  'Explain this error in simple words: ',
  'Generate troubleshooting steps for: ',
  'Generate a ticket note for: ',
  'Generate a customer email about: ',
  'Explain this PowerShell command: ',
  'Summarize this BitTitan error: ',
  'Explain this migration error: ',
  'Give escalation advice for: ',
];

export default function Chat() {
  const [messages, setMessages] = useLocalStorage<ChatMessage[]>('chat-history', []);
  const [input, setInput] = useState('');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const bottomRef = useRef<HTMLDivElement>(null);

  const send = async (text?: string) => {
    const content = (text ?? input).trim();
    if (!content || loading) return;
    setInput('');
    setError('');
    const next: ChatMessage[] = [...messages, { role: 'user', content }];
    setMessages(next);
    setLoading(true);
    try {
      const reply = await chat(next.slice(-12));
      setMessages([...next, { role: 'assistant', content: reply }]);
      setTimeout(() => bottomRef.current?.scrollIntoView({ behavior: 'smooth' }), 50);
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="flex h-[calc(100vh-10rem)] flex-col">
      <PageHeader title="AI Chat Assistant" subtitle="Ask anything: paste errors, tickets, BitTitan failures or migration logs. The assistant answers with portal paths, PowerShell, roles, risks and escalation advice." icon={<Bot size={20} />} />

      {!aiEnabled() && (
        <Card className="p-6 text-center">
          <p className="text-sm text-slate-500 dark:text-slate-400">
            AI is not configured yet. Go to <Link to="/settings" className="font-semibold text-violet-600 hover:underline">Settings</Link>, choose a provider (OpenRouter / OpenAI / Claude), paste your API key and test the connection.
            <br />Everything else in the portal works without AI.
          </p>
        </Card>
      )}

      {aiEnabled() && (
        <>
          <div className="mb-3 flex flex-wrap gap-2">
            {quickPrompts.map((p) => (
              <button key={p} onClick={() => setInput(p)} className="rounded-full border border-violet-200 dark:border-violet-800 bg-violet-50 dark:bg-violet-900/20 px-3 py-1 text-xs font-medium text-violet-700 dark:text-violet-300 hover:bg-violet-100 dark:hover:bg-violet-900/40">
                {p.replace(': ', '')}
              </button>
            ))}
            {messages.length > 0 && (
              <button onClick={() => setMessages([])} className="ml-auto flex items-center gap-1 rounded-full border border-slate-200 dark:border-slate-700 px-3 py-1 text-xs text-slate-500 hover:text-red-500">
                <Trash2 size={12} /> Clear chat
              </button>
            )}
          </div>

          <Card className="flex-1 overflow-y-auto p-4">
            {messages.length === 0 && (
              <div className="flex h-full items-center justify-center text-sm text-slate-400">
                Start by pasting an error message, a ticket, or asking a question.
              </div>
            )}
            <div className="space-y-4">
              {messages.map((m, i) => (
                <div key={i} className={`flex ${m.role === 'user' ? 'justify-end' : 'justify-start'}`}>
                  <div className={`max-w-[85%] rounded-2xl px-4 py-3 text-sm leading-relaxed ${m.role === 'user' ? 'bg-blue-600 text-white' : 'bg-slate-100 dark:bg-slate-700 text-slate-800 dark:text-slate-100'}`}>
                    <pre className="whitespace-pre-wrap font-sans">{m.content}</pre>
                    {m.role === 'assistant' && <div className="mt-2"><CopyButton text={m.content} /></div>}
                  </div>
                </div>
              ))}
              {loading && (
                <div className="flex items-center gap-2 text-sm text-slate-400"><Loader2 size={16} className="animate-spin" /> Thinking…</div>
              )}
              {error && <p className="text-sm text-red-500">{error}</p>}
              <div ref={bottomRef} />
            </div>
          </Card>

          <div className="mt-3 flex gap-2">
            <textarea
              value={input}
              onChange={(e) => setInput(e.target.value)}
              onKeyDown={(e) => { if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); send(); } }}
              rows={2}
              placeholder="Ask a question or paste an error / ticket / log… (Enter to send, Shift+Enter for newline)"
              className="flex-1 rounded-xl border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-900 px-4 py-3 text-sm text-slate-800 dark:text-slate-100 focus:border-violet-500 focus:outline-none"
            />
            <button onClick={() => send()} disabled={loading || !input.trim()}
              className="self-end rounded-xl bg-violet-600 p-3 text-white hover:bg-violet-700 disabled:opacity-50">
              <SendHorizonal size={18} />
            </button>
          </div>
        </>
      )}
    </div>
  );
}
