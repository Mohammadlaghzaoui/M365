import { useState } from 'react';
import { StickyNote, Plus, Pin, Trash2 } from 'lucide-react';
import { Button, Card, PageHeader, EmptyState } from '../components/ui';
import { Note } from '../types';
import { uid, useLocalStorage } from '../store/useLocalStorage';

export default function Notes() {
  const [notes, setNotes] = useLocalStorage<Note[]>('notes', []);
  const [selectedId, setSelectedId] = useState<string | null>(notes[0]?.id ?? null);
  const selected = notes.find((n) => n.id === selectedId) ?? null;

  const add = () => {
    const n: Note = { id: uid(), title: 'New note', body: '', updatedAt: new Date().toISOString(), pinned: false };
    setNotes((prev) => [n, ...prev]);
    setSelectedId(n.id);
  };

  const update = (patch: Partial<Note>) =>
    setNotes((prev) => prev.map((n) => (n.id === selectedId ? { ...n, ...patch, updatedAt: new Date().toISOString() } : n)));

  const sorted = [...notes].sort((a, b) => Number(b.pinned) - Number(a.pinned) || b.updatedAt.localeCompare(a.updatedAt));

  return (
    <div>
      <PageHeader title="Personal Notes" subtitle="Quick scratchpad for tenant specifics, customer quirks and handy snippets. Stored locally in your browser." icon={<StickyNote size={20} />} />
      <div className="mb-4"><Button onClick={add}><Plus size={15} /> New note</Button></div>
      {notes.length === 0 ? (
        <EmptyState message="No notes yet — create your first one." />
      ) : (
        <div className="grid gap-6 lg:grid-cols-[300px_1fr]">
          <div className="space-y-2">
            {sorted.map((n) => (
              <button key={n.id} onClick={() => setSelectedId(n.id)}
                className={`w-full rounded-lg border p-3 text-left ${selectedId === n.id ? 'border-blue-500 bg-blue-50 dark:bg-blue-900/20' : 'border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800'}`}>
                <div className="flex items-center gap-2">
                  {n.pinned && <Pin size={12} className="text-amber-500" />}
                  <span className="flex-1 truncate text-sm font-semibold text-slate-700 dark:text-slate-200">{n.title || 'Untitled'}</span>
                </div>
                <div className="mt-1 truncate text-xs text-slate-400">{n.body.slice(0, 60) || 'empty'}</div>
              </button>
            ))}
          </div>
          {selected && (
            <Card className="p-5 self-start">
              <div className="mb-3 flex items-center gap-2">
                <input value={selected.title} onChange={(e) => update({ title: e.target.value })}
                  className="flex-1 rounded-lg border border-transparent bg-transparent px-2 py-1 text-lg font-bold text-slate-800 dark:text-slate-100 hover:border-slate-200 dark:hover:border-slate-600 focus:border-blue-500 focus:outline-none" />
                <button onClick={() => update({ pinned: !selected.pinned })} title="Pin">
                  <Pin size={17} className={selected.pinned ? 'fill-amber-400 text-amber-500' : 'text-slate-300'} />
                </button>
                <button onClick={() => { setNotes((prev) => prev.filter((n) => n.id !== selected.id)); setSelectedId(null); }} title="Delete">
                  <Trash2 size={17} className="text-slate-400 hover:text-red-500" />
                </button>
              </div>
              <textarea value={selected.body} onChange={(e) => update({ body: e.target.value })} rows={18} placeholder="Write your note…"
                className="w-full rounded-lg border border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-900 p-4 text-sm leading-relaxed text-slate-700 dark:text-slate-200 focus:border-blue-500 focus:outline-none" />
              <div className="mt-2 text-xs text-slate-400">Last updated {new Date(selected.updatedAt).toLocaleString()}</div>
            </Card>
          )}
        </div>
      )}
    </div>
  );
}
