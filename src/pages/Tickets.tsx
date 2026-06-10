import { useState } from 'react';
import { Ticket as TicketIcon, Wand2, Trash2, Send, Loader2, ExternalLink } from 'lucide-react';
import { Badge, Button, Card, Field, PageHeader, Select, TextArea, TextOutput } from '../components/ui';
import { AIHelper } from '../components/AIHelper';
import { generateTicket, GeneratedTicket, serviceOptions, categoryOptions } from '../data/ticketGenerator';
import { Ticket } from '../types';
import { uid, useLocalStorage } from '../store/useLocalStorage';
import { serviceNowEnabled } from '../store/settings';
import { createIncident, SNIncidentResult } from '../services/servicenow';

const emptyTicket = (): Ticket => ({
  id: uid(),
  createdAt: new Date().toISOString(),
  customer: '',
  userName: '',
  userEmail: '',
  service: 'Entra ID',
  category: '',
  errorMessage: '',
  impact: '',
  urgency: 'medium',
  tried: '',
  screenshots: false,
  language: 'en',
  status: 'open',
});

export default function Tickets() {
  const [tickets, setTickets] = useLocalStorage<Ticket[]>('tickets', []);
  const [form, setForm] = useState<Ticket>(emptyTicket());
  const [generated, setGenerated] = useState<GeneratedTicket | null>(null);
  const [snBusy, setSnBusy] = useState(false);
  const [snResult, setSnResult] = useState<SNIncidentResult | null>(null);
  const [snError, setSnError] = useState('');

  const sendToServiceNow = async () => {
    if (!generated) return;
    setSnBusy(true);
    setSnError('');
    try {
      setSnResult(await createIncident(form, `${generated.fullDescription}\n\nFirst troubleshooting steps:\n${generated.troubleshooting}`));
    } catch (e) {
      setSnError(e instanceof Error ? e.message : String(e));
    } finally {
      setSnBusy(false);
    }
  };

  const set = <K extends keyof Ticket>(k: K, v: Ticket[K]) => setForm((f) => ({ ...f, [k]: v }));

  const generate = () => {
    const t = { ...form, id: form.id || uid(), category: form.category || (categoryOptions[form.service]?.[0] ?? '') };
    setGenerated(generateTicket(t));
    setTickets((prev) => {
      const exists = prev.some((p) => p.id === t.id);
      return exists ? prev.map((p) => (p.id === t.id ? t : p)) : [t, ...prev];
    });
  };

  const setStatus = (id: string, status: Ticket['status']) =>
    setTickets((prev) => prev.map((t) => (t.id === id ? { ...t, status } : t)));

  return (
    <div>
      <PageHeader title="Ticket Assistant" subtitle="Fill in the ticket details — the assistant generates the full ticket package: descriptions, troubleshooting, PowerShell, customer update, escalation and closure notes." icon={<TicketIcon size={20} />} />
      <div className="grid gap-6 xl:grid-cols-[420px_1fr]">
        <Card className="p-5 self-start">
          <div className="space-y-3">
            <Field label="Customer / tenant name" value={form.customer} onChange={(v) => set('customer', v)} placeholder="Contoso BV" />
            <div className="grid grid-cols-2 gap-3">
              <Field label="User full name" value={form.userName} onChange={(v) => set('userName', v)} placeholder="Jane Doe" />
              <Field label="User email" value={form.userEmail} onChange={(v) => set('userEmail', v)} placeholder="jane@contoso.com" />
            </div>
            <Select label="Affected service" value={form.service} onChange={(v) => { set('service', v); set('category', ''); }} options={serviceOptions.map((s) => ({ value: s, label: s }))} />
            <Select label="Issue category" value={form.category} onChange={(v) => set('category', v)} options={[{ value: '', label: '— select —' }, ...(categoryOptions[form.service] ?? []).map((c) => ({ value: c, label: c }))]} />
            <TextArea label="Error message (exact text)" value={form.errorMessage} onChange={(v) => set('errorMessage', v)} rows={2} placeholder="Paste the exact error..." />
            <Field label="Business impact" value={form.impact} onChange={(v) => set('impact', v)} placeholder="User cannot work / department blocked..." />
            <div className="grid grid-cols-2 gap-3">
              <Select label="Urgency" value={form.urgency} onChange={(v) => set('urgency', v as Ticket['urgency'])} options={['low', 'medium', 'high', 'critical'].map((u) => ({ value: u, label: u }))} />
              <Select label="Language output" value={form.language} onChange={(v) => set('language', v as 'en' | 'nl')} options={[{ value: 'en', label: 'English' }, { value: 'nl', label: 'Dutch' }]} />
            </div>
            <TextArea label="What has already been tried" value={form.tried} onChange={(v) => set('tried', v)} rows={2} />
            <label className="flex items-center gap-2 text-sm text-slate-700 dark:text-slate-300">
              <input type="checkbox" checked={form.screenshots} onChange={(e) => set('screenshots', e.target.checked)} className="h-4 w-4 rounded border-slate-300 text-blue-600" />
              Screenshots available
            </label>
            <Button onClick={generate} className="w-full"><Wand2 size={16} /> Generate ticket package</Button>
          </div>
        </Card>

        <div className="space-y-4">
          {!generated && (
            <Card className="p-8 text-center text-sm text-slate-400">Fill in the form and click <strong>Generate</strong>. Every section gets its own copy button. Works fully offline — enable AI in Settings for enhanced versions.</Card>
          )}
          {generated && (
            <>
              {serviceNowEnabled() && (
                <Card className="p-4 flex flex-wrap items-center gap-3">
                  <Button onClick={sendToServiceNow} disabled={snBusy}>
                    {snBusy ? <Loader2 size={15} className="animate-spin" /> : <Send size={15} />} Create incident in ServiceNow
                  </Button>
                  {snResult && (
                    <a href={snResult.link} target="_blank" rel="noreferrer" className="flex items-center gap-1.5 text-sm font-semibold text-emerald-600 hover:underline">
                      {snResult.number} created <ExternalLink size={13} />
                    </a>
                  )}
                  {snError && <span className="text-sm text-red-500">{snError}</span>}
                </Card>
              )}
              <TextOutput title="Short description" text={generated.shortDescription} />
              <TextOutput title="Full ticket description" text={generated.fullDescription} />
              <TextOutput title="First troubleshooting steps" text={generated.troubleshooting} />
              <div className="grid gap-4 lg:grid-cols-2">
                <TextOutput title="Possible causes" text={generated.causes} />
                <TextOutput title="Microsoft portal paths" text={generated.portals} />
              </div>
              {generated.powershell && <TextOutput title="PowerShell commands" text={generated.powershell} />}
              <TextOutput title="Evidence needed" text={generated.evidence} />
              <TextOutput title="Customer update" text={generated.customerUpdate} />
              <TextOutput title="Escalation note" text={generated.escalationNote} />
              <TextOutput title="Closure note" text={generated.closureNote} />
              <AIHelper
                context={`Ticket:\n${generated.fullDescription}\n\nPlanned troubleshooting:\n${generated.troubleshooting}`}
                defaultPrompt="Improve the troubleshooting plan for this ticket and write a more specific customer update."
              />
            </>
          )}

          {tickets.length > 0 && (
            <Card className="p-5">
              <h3 className="mb-3 text-sm font-semibold uppercase tracking-wide text-slate-500 dark:text-slate-400">Ticket log ({tickets.length})</h3>
              <div className="space-y-2">
                {tickets.map((t) => (
                  <div key={t.id} className="flex flex-wrap items-center gap-2 rounded-lg border border-slate-200 dark:border-slate-700 px-3 py-2 text-sm">
                    <span className="font-medium text-slate-700 dark:text-slate-200">{t.customer || '—'}</span>
                    <span className="text-slate-400">·</span>
                    <span className="text-slate-500 dark:text-slate-300">{t.service}: {t.category}</span>
                    <Badge color={t.urgency === 'high' || t.urgency === 'critical' ? 'red' : 'gray'}>{t.urgency}</Badge>
                    <div className="ml-auto flex items-center gap-2">
                      <select value={t.status} onChange={(e) => setStatus(t.id, e.target.value as Ticket['status'])} className="rounded border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-900 px-2 py-1 text-xs">
                        {['open', 'in-progress', 'waiting', 'escalated', 'closed'].map((s) => <option key={s} value={s}>{s}</option>)}
                      </select>
                      <button onClick={() => setTickets((prev) => prev.filter((p) => p.id !== t.id))} className="text-slate-400 hover:text-red-500"><Trash2 size={14} /></button>
                    </div>
                  </div>
                ))}
              </div>
            </Card>
          )}
        </div>
      </div>
    </div>
  );
}
