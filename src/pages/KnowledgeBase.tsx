import { useMemo, useState } from 'react';
import { BookOpen, Plus, Star, Pencil, Trash2 } from 'lucide-react';
import { Badge, Button, Card, CopyButton, Field, PageHeader, Select, TextArea } from '../components/ui';
import { seedArticles } from '../data/kbArticles';
import { KBArticle } from '../types';
import { uid, useLocalStorage } from '../store/useLocalStorage';

const services = ['All', 'Entra ID', 'Exchange Online', 'SharePoint Online', 'Teams', 'Migration', 'Cross-Tenant Migration', 'BitTitan', 'Syskit', 'Security'];

export default function KnowledgeBase() {
  // Custom articles + overrides (favorites/edits of seeds) are stored locally.
  const [custom, setCustom] = useLocalStorage<KBArticle[]>('kb-custom', []);
  const [overrides, setOverrides] = useLocalStorage<Record<string, Partial<KBArticle>>>('kb-overrides', {});
  const [query, setQuery] = useState('');
  const [service, setService] = useState('All');
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [editing, setEditing] = useState<KBArticle | null>(null);
  const [favOnly, setFavOnly] = useState(false);

  const articles = useMemo<KBArticle[]>(() => {
    const seeds = seedArticles.map((a) => ({ ...a, ...overrides[a.id] }));
    return [...custom, ...seeds];
  }, [custom, overrides]);

  const filtered = articles.filter((a) => {
    if (favOnly && !a.favorite) return false;
    if (service !== 'All' && a.service !== service) return false;
    const q = query.toLowerCase();
    return !q || a.title.toLowerCase().includes(q) || a.body.toLowerCase().includes(q) || a.tags.some((t) => t.includes(q));
  });

  const selected = articles.find((a) => a.id === selectedId) ?? null;

  const toggleFav = (a: KBArticle) => {
    if (a.custom) setCustom((prev) => prev.map((c) => (c.id === a.id ? { ...c, favorite: !c.favorite } : c)));
    else setOverrides((prev) => ({ ...prev, [a.id]: { ...prev[a.id], favorite: !a.favorite } }));
  };

  const saveEdit = () => {
    if (!editing) return;
    if (editing.custom) {
      setCustom((prev) => {
        const exists = prev.some((c) => c.id === editing.id);
        return exists ? prev.map((c) => (c.id === editing.id ? editing : c)) : [editing, ...prev];
      });
    } else {
      setOverrides((prev) => ({ ...prev, [editing.id]: { title: editing.title, body: editing.body, tags: editing.tags, service: editing.service } }));
    }
    setSelectedId(editing.id);
    setEditing(null);
  };

  return (
    <div>
      <PageHeader title="Knowledge Base" subtitle="Searchable how-to articles per service. Add your own, edit, favorite, and copy steps straight into tickets." icon={<BookOpen size={20} />} />

      <div className="mb-4 flex flex-wrap items-center gap-3">
        <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Search articles, tags, content..."
          className="w-72 rounded-lg border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-900 px-3 py-2 text-sm focus:border-blue-500 focus:outline-none text-slate-800 dark:text-slate-100" />
        <div className="w-56"><Select label="" value={service} onChange={setService} options={services.map((s) => ({ value: s, label: s }))} /></div>
        <button onClick={() => setFavOnly(!favOnly)} className={`flex items-center gap-1.5 rounded-lg border px-3 py-2 text-sm font-medium ${favOnly ? 'border-amber-400 text-amber-500' : 'border-slate-300 dark:border-slate-600 text-slate-500'}`}>
          <Star size={14} className={favOnly ? 'fill-amber-400 text-amber-400' : ''} /> Favorites
        </button>
        <Button variant="secondary" onClick={() => setEditing({ id: uid(), title: '', service: 'Entra ID', tags: [], body: '', favorite: false, custom: true })}>
          <Plus size={15} /> Add article
        </Button>
      </div>

      {editing && (
        <Card className="mb-5 p-5 space-y-3">
          <Field label="Title" value={editing.title} onChange={(v) => setEditing({ ...editing, title: v })} />
          <div className="grid gap-3 sm:grid-cols-2">
            <Select label="Service" value={editing.service} onChange={(v) => setEditing({ ...editing, service: v })} options={services.filter((s) => s !== 'All').map((s) => ({ value: s, label: s }))} />
            <Field label="Tags (comma separated)" value={editing.tags.join(', ')} onChange={(v) => setEditing({ ...editing, tags: v.split(',').map((t) => t.trim()).filter(Boolean) })} />
          </div>
          <TextArea label="Content (steps, commands, notes)" value={editing.body} onChange={(v) => setEditing({ ...editing, body: v })} rows={10} />
          <div className="flex gap-2">
            <Button onClick={saveEdit}>Save article</Button>
            <Button variant="secondary" onClick={() => setEditing(null)}>Cancel</Button>
          </div>
        </Card>
      )}

      <div className="grid gap-6 lg:grid-cols-[360px_1fr]">
        <div className="space-y-2 lg:max-h-[calc(100vh-16rem)] lg:overflow-y-auto pr-1">
          {filtered.map((a) => (
            <button key={a.id} onClick={() => setSelectedId(a.id)}
              className={`w-full rounded-lg border p-3 text-left ${selectedId === a.id ? 'border-blue-500 bg-blue-50 dark:bg-blue-900/20' : 'border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800 hover:border-blue-300'}`}>
              <div className="flex items-center gap-2">
                <span className="flex-1 text-sm font-semibold text-slate-700 dark:text-slate-200">{a.title}</span>
                {a.favorite && <Star size={13} className="fill-amber-400 text-amber-400" />}
              </div>
              <div className="mt-1 flex flex-wrap items-center gap-1.5">
                <Badge color="blue">{a.service}</Badge>
                {a.custom && <Badge color="purple">custom</Badge>}
                {a.tags.slice(0, 3).map((t) => <span key={t} className="text-xs text-slate-400">#{t}</span>)}
              </div>
            </button>
          ))}
          {filtered.length === 0 && <p className="p-4 text-sm text-slate-400">No articles match.</p>}
        </div>

        {selected ? (
          <Card className="p-5 self-start">
            <div className="mb-3 flex items-center gap-2">
              <h2 className="flex-1 text-lg font-bold text-slate-800 dark:text-slate-100">{selected.title}</h2>
              <button onClick={() => toggleFav(selected)} title="Favorite"><Star size={18} className={selected.favorite ? 'fill-amber-400 text-amber-400' : 'text-slate-300'} /></button>
              <button onClick={() => setEditing({ ...selected })} title="Edit"><Pencil size={16} className="text-slate-400 hover:text-blue-500" /></button>
              {selected.custom && (
                <button onClick={() => { setCustom((prev) => prev.filter((c) => c.id !== selected.id)); setSelectedId(null); }} title="Delete">
                  <Trash2 size={16} className="text-slate-400 hover:text-red-500" />
                </button>
              )}
              <CopyButton text={selected.body} label="Copy steps" />
            </div>
            <pre className="whitespace-pre-wrap rounded-lg bg-slate-50 dark:bg-slate-900 p-4 text-sm leading-relaxed text-slate-700 dark:text-slate-200 font-sans">{selected.body}</pre>
          </Card>
        ) : (
          <Card className="p-10 text-center text-sm text-slate-400 self-start">Select an article to read it.</Card>
        )}
      </div>
    </div>
  );
}
