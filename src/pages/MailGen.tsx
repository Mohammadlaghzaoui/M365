import { useState } from 'react';
import { Send, Sparkles, Loader2 } from 'lucide-react';
import { Button, Card, Field, PageHeader, Select, TextOutput } from '../components/ui';
import { generateMail, mailTypeLabels, MailVars } from '../data/mailTemplates';
import { MailType } from '../types';
import { chat } from '../services/ai';
import { aiEnabled } from '../store/settings';

export default function MailGen() {
  const [type, setType] = useState<MailType>('ack');
  const [lang, setLang] = useState<'en' | 'nl'>('en');
  const [tone, setTone] = useState<'friendly' | 'professional' | 'short' | 'detailed'>('professional');
  const [vars, setVars] = useState<MailVars>({ customer: '', user: '', ticketId: '', issue: '', date: '', engineer: '' });
  const [aiVersion, setAiVersion] = useState('');
  const [loading, setLoading] = useState(false);

  const mail = generateMail(type, lang, tone, vars);

  const enhance = async () => {
    setLoading(true);
    setAiVersion('');
    try {
      const res = await chat([{
        role: 'user',
        content: `Rewrite this customer email. Keep the language (${lang === 'nl' ? 'Dutch' : 'English'}) and the ${tone} tone, make it natural and specific to the issue "${vars.issue}".\n\nSubject: ${mail.subject}\n\n${mail.body}`,
      }]);
      setAiVersion(res);
    } catch (e) {
      setAiVersion(`AI error: ${e instanceof Error ? e.message : e}`);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div>
      <PageHeader title="Customer Mail Generator" subtitle="Professional customer emails in English or Dutch — acknowledgements, updates, migration communications and security recommendations." icon={<Send size={20} />} />
      <div className="grid gap-6 lg:grid-cols-[380px_1fr]">
        <Card className="p-5 self-start space-y-3">
          <Select label="Email type" value={type} onChange={(v) => setType(v as MailType)} options={mailTypeLabels} />
          <div className="grid grid-cols-2 gap-3">
            <Select label="Language" value={lang} onChange={(v) => setLang(v as 'en' | 'nl')} options={[{ value: 'en', label: 'English' }, { value: 'nl', label: 'Dutch (Nederlands)' }]} />
            <Select label="Tone" value={tone} onChange={(v) => setTone(v as typeof tone)} options={['friendly', 'professional', 'short', 'detailed'].map((t) => ({ value: t, label: t }))} />
          </div>
          <Field label="Customer / organization" value={vars.customer} onChange={(v) => setVars({ ...vars, customer: v })} placeholder="Contoso BV" />
          <Field label="Recipient name" value={vars.user} onChange={(v) => setVars({ ...vars, user: v })} placeholder="Jane" />
          <Field label="Ticket ID" value={vars.ticketId} onChange={(v) => setVars({ ...vars, ticketId: v })} placeholder="INC-10421" />
          <Field label="Issue / topic" value={vars.issue} onChange={(v) => setVars({ ...vars, issue: v })} placeholder="Outlook cannot send email" />
          <Field label="Date (update ETA / migration date)" value={vars.date} onChange={(v) => setVars({ ...vars, date: v })} placeholder="Friday 26 June" />
          <Field label="Your name (signature)" value={vars.engineer} onChange={(v) => setVars({ ...vars, engineer: v })} placeholder="Mohamed — Service Desk" />
          {aiEnabled() && (
            <Button variant="ai" onClick={enhance} disabled={loading} className="w-full">
              {loading ? <Loader2 size={16} className="animate-spin" /> : <Sparkles size={16} />} Enhance with AI
            </Button>
          )}
        </Card>
        <div className="space-y-4">
          <TextOutput title={`Subject`} text={mail.subject} />
          <TextOutput title={`Email body (${lang.toUpperCase()} · ${tone})`} text={mail.body} />
          {aiVersion && <TextOutput title="AI-enhanced version" text={aiVersion} />}
        </div>
      </div>
    </div>
  );
}
