import { MailType } from '../types';

export interface MailVars {
  customer: string;
  user: string;
  ticketId: string;
  issue: string;
  date: string;
  engineer: string;
}

type Lang = 'en' | 'nl';
type Tone = 'friendly' | 'professional' | 'short' | 'detailed';

export const mailTypeLabels: { value: MailType; label: string }[] = [
  { value: 'ack', label: 'Initial acknowledgement' },
  { value: 'more-info', label: 'Request more information' },
  { value: 'investigating', label: 'Issue under investigation' },
  { value: 'resolved', label: 'Issue resolved' },
  { value: 'escalation', label: 'Escalation update' },
  { value: 'migration-announce', label: 'Migration announcement' },
  { value: 'migration-reminder', label: 'Migration reminder' },
  { value: 'cutover-done', label: 'Cutover completed' },
  { value: 'post-migration', label: 'Post-migration support' },
  { value: 'security-recommendation', label: 'Security recommendation' },
];

const greetings: Record<Lang, Record<Tone, (v: MailVars) => string>> = {
  en: {
    friendly: (v) => `Hi ${v.user || 'there'},`,
    professional: (v) => `Dear ${v.user || 'customer'},`,
    short: (v) => `Hi ${v.user || ''},`.replace(' ,', ','),
    detailed: (v) => `Dear ${v.user || 'customer'},`,
  },
  nl: {
    friendly: (v) => `Hoi ${v.user || 'daar'},`,
    professional: (v) => `Beste ${v.user || 'klant'},`,
    short: (v) => `Hoi ${v.user || ''},`.replace(' ,', ','),
    detailed: (v) => `Beste ${v.user || 'klant'},`,
  },
};

const closings: Record<Lang, Record<Tone, (v: MailVars) => string>> = {
  en: {
    friendly: (v) => `Thanks and kind regards,\n${v.engineer || 'Your IT Support Team'}`,
    professional: (v) => `Kind regards,\n${v.engineer || 'IT Support'}\nService Desk`,
    short: (v) => `Regards,\n${v.engineer || 'IT Support'}`,
    detailed: (v) => `If you have any questions, simply reply to this email or call the service desk.\n\nKind regards,\n${v.engineer || 'IT Support'}\nService Desk`,
  },
  nl: {
    friendly: (v) => `Bedankt en vriendelijke groet,\n${v.engineer || 'Je IT Support Team'}`,
    professional: (v) => `Met vriendelijke groet,\n${v.engineer || 'IT Support'}\nServicedesk`,
    short: (v) => `Groet,\n${v.engineer || 'IT Support'}`,
    detailed: (v) => `Heb je vragen? Reageer gerust op deze e-mail of bel de servicedesk.\n\nMet vriendelijke groet,\n${v.engineer || 'IT Support'}\nServicedesk`,
  },
};

const bodies: Record<MailType, Record<Lang, { subject: (v: MailVars) => string; body: (v: MailVars, tone: Tone) => string }>> = {
  ack: {
    en: {
      subject: (v) => `[${v.ticketId || 'Ticket'}] We received your request — ${v.issue || 'your issue'}`,
      body: (v, tone) =>
        tone === 'short'
          ? `We received your request about "${v.issue}" and registered it as ${v.ticketId || 'a ticket'}. We will get back to you shortly.`
          : `Thank you for contacting the service desk. We have received your request regarding "${v.issue}" and registered it under ticket number ${v.ticketId || '<ticket>'}.\n\nAn engineer has been assigned and will start the investigation. We will keep you informed of the progress.${tone === 'detailed' ? '\n\nTo speed things up, it helps if you can already send us:\n- A screenshot of the exact error message\n- The time the issue last occurred\n- Whether colleagues experience the same issue' : ''}`,
    },
    nl: {
      subject: (v) => `[${v.ticketId || 'Ticket'}] Je melding is ontvangen — ${v.issue || 'je melding'}`,
      body: (v, tone) =>
        tone === 'short'
          ? `We hebben je melding over "${v.issue}" ontvangen en geregistreerd als ${v.ticketId || 'ticket'}. We komen er snel bij je op terug.`
          : `Bedankt voor je melding bij de servicedesk. We hebben je verzoek over "${v.issue}" ontvangen en geregistreerd onder ticketnummer ${v.ticketId || '<ticket>'}.\n\nEen engineer is toegewezen en start met het onderzoek. We houden je op de hoogte van de voortgang.${tone === 'detailed' ? '\n\nOm sneller te kunnen helpen, ontvangen we graag alvast:\n- Een screenshot van de exacte foutmelding\n- Het tijdstip waarop het probleem zich voor het laatst voordeed\n- Of collega\'s hetzelfde probleem ervaren' : ''}`,
    },
  },
  'more-info': {
    en: {
      subject: (v) => `[${v.ticketId || 'Ticket'}] More information needed — ${v.issue || ''}`,
      body: (v, tone) =>
        `To continue the investigation of "${v.issue}" we need some additional information:\n\n1. A screenshot of the exact error message\n2. The date and time the issue last occurred\n3. Does the issue occur on all devices (also via the browser)?\n4. Have any recent changes been made (new phone, password change)?\n\n${tone === 'short' ? 'Could you send this to us?' : 'As soon as we receive this information, we will continue immediately. The ticket remains open in the meantime.'}`,
    },
    nl: {
      subject: (v) => `[${v.ticketId || 'Ticket'}] Aanvullende informatie nodig — ${v.issue || ''}`,
      body: (v, tone) =>
        `Om het onderzoek naar "${v.issue}" voort te zetten hebben we wat aanvullende informatie nodig:\n\n1. Een screenshot van de exacte foutmelding\n2. Datum en tijdstip waarop het probleem zich voor het laatst voordeed\n3. Treedt het probleem op alle apparaten op (ook via de browser)?\n4. Zijn er recent wijzigingen geweest (nieuwe telefoon, wachtwoordwijziging)?\n\n${tone === 'short' ? 'Kun je dit naar ons sturen?' : 'Zodra we deze informatie ontvangen, gaan we direct verder. Het ticket blijft ondertussen open staan.'}`,
    },
  },
  investigating: {
    en: {
      subject: (v) => `[${v.ticketId || 'Ticket'}] Update: investigation in progress — ${v.issue || ''}`,
      body: (v, tone) =>
        `A quick update on your ticket regarding "${v.issue}".\n\nOur engineer is actively investigating the issue. ${tone === 'detailed' ? 'We have analysed the sign-in logs and message traces and are currently narrowing down the root cause. ' : ''}We expect to give you a next update by ${v.date || '<date>'}.\n\nNo action is needed from your side at this moment.`,
    },
    nl: {
      subject: (v) => `[${v.ticketId || 'Ticket'}] Update: onderzoek loopt — ${v.issue || ''}`,
      body: (v, tone) =>
        `Een korte update over je ticket met betrekking tot "${v.issue}".\n\nOnze engineer onderzoekt het probleem actief. ${tone === 'detailed' ? 'We hebben de aanmeldlogboeken en message traces geanalyseerd en zijn de oorzaak aan het afbakenen. ' : ''}We verwachten je uiterlijk ${v.date || '<datum>'} een volgende update te geven.\n\nEr is op dit moment geen actie van jouw kant nodig.`,
    },
  },
  resolved: {
    en: {
      subject: (v) => `[${v.ticketId || 'Ticket'}] Resolved — ${v.issue || ''}`,
      body: (v, tone) =>
        `Good news: the issue "${v.issue}" has been resolved.\n\n${tone === 'detailed' ? 'Summary of the fix:\n- Root cause: <root cause>\n- Action taken: <action>\n- Verified: <verification>\n\n' : ''}Could you confirm everything works as expected on your side? If we do not hear from you within 3 business days, we will close the ticket.${tone === 'friendly' ? '\n\nThanks for your patience!' : ''}`,
    },
    nl: {
      subject: (v) => `[${v.ticketId || 'Ticket'}] Opgelost — ${v.issue || ''}`,
      body: (v, tone) =>
        `Goed nieuws: het probleem "${v.issue}" is opgelost.\n\n${tone === 'detailed' ? 'Samenvatting van de oplossing:\n- Oorzaak: <oorzaak>\n- Uitgevoerde actie: <actie>\n- Geverifieerd: <verificatie>\n\n' : ''}Kun je bevestigen dat alles aan jouw kant weer naar verwachting werkt? Als we binnen 3 werkdagen niets horen, sluiten we het ticket.${tone === 'friendly' ? '\n\nBedankt voor je geduld!' : ''}`,
    },
  },
  escalation: {
    en: {
      subject: (v) => `[${v.ticketId || 'Ticket'}] Update: escalated to specialist team — ${v.issue || ''}`,
      body: (v) =>
        `An update on your ticket regarding "${v.issue}".\n\nThe issue requires specialist attention and has been escalated to our specialist team${v.date ? ` on ${v.date}` : ''}. They have all the investigation details and will continue with priority.\n\nWe remain your point of contact and will keep you informed of every development.`,
    },
    nl: {
      subject: (v) => `[${v.ticketId || 'Ticket'}] Update: geëscaleerd naar specialistisch team — ${v.issue || ''}`,
      body: (v) =>
        `Een update over je ticket met betrekking tot "${v.issue}".\n\nHet probleem vereist specialistische aandacht en is${v.date ? ` op ${v.date}` : ''} geëscaleerd naar ons specialistische team. Zij beschikken over alle onderzoeksgegevens en pakken dit met prioriteit op.\n\nWij blijven je aanspreekpunt en houden je van elke ontwikkeling op de hoogte.`,
    },
  },
  'migration-announce': {
    en: {
      subject: (v) => `Important: your mailbox moves to the new environment on ${v.date || '<date>'}`,
      body: (v, tone) =>
        `We are improving the IT environment of ${v.customer || 'your organization'}. As part of this, your mailbox and data will be migrated to the new Microsoft 365 environment.\n\nWhat you need to know:\n- Migration date: ${v.date || '<date>'}\n- Your email address stays the same\n- During the migration window (evening/weekend) email may be briefly unavailable\n\nWhat we ask of you:\n- Before the migration: close Outlook at the end of the day\n- After the migration: follow the start-up guide we will send you (signing in again + setting up your phone)\n${tone === 'detailed' ? '\nWhat does NOT move automatically:\n- Email signatures (copy them before the migration!)\n- Links to old Teams meetings — recurring meetings will be recreated\n\n' : ''}We will send a reminder shortly before the migration.`,
    },
    nl: {
      subject: (v) => `Belangrijk: je mailbox verhuist op ${v.date || '<datum>'} naar de nieuwe omgeving`,
      body: (v, tone) =>
        `We verbeteren de IT-omgeving van ${v.customer || 'jouw organisatie'}. Als onderdeel hiervan worden je mailbox en gegevens gemigreerd naar de nieuwe Microsoft 365-omgeving.\n\nWat je moet weten:\n- Migratiedatum: ${v.date || '<datum>'}\n- Je e-mailadres blijft hetzelfde\n- Tijdens het migratievenster (avond/weekend) kan e-mail kort niet beschikbaar zijn\n\nWat we van je vragen:\n- Vóór de migratie: sluit Outlook aan het einde van de dag af\n- Na de migratie: volg de opstartgids die we je toesturen (opnieuw aanmelden + telefoon instellen)\n${tone === 'detailed' ? '\nWat NIET automatisch meegaat:\n- E-mailhandtekeningen (kopieer deze vóór de migratie!)\n- Links naar oude Teams-vergaderingen — terugkerende vergaderingen worden opnieuw aangemaakt\n\n' : ''}Kort voor de migratie sturen we nog een herinnering.`,
    },
  },
  'migration-reminder': {
    en: {
      subject: (v) => `Reminder: mailbox migration this ${v.date || '<date>'} — action needed`,
      body: (v) =>
        `A reminder: on ${v.date || '<date>'} your mailbox moves to the new Microsoft 365 environment.\n\nYour checklist for the migration day:\n1. Copy your email signature to a Word document\n2. Close Outlook at the end of the day\n3. After the go-live message: restart your computer and sign in with the instructions provided\n4. Set up your phone again with the new account settings\n\nThe service desk has extra staffing on the days after the migration to help you quickly.`,
    },
    nl: {
      subject: (v) => `Herinnering: mailboxmigratie op ${v.date || '<datum>'} — actie nodig`,
      body: (v) =>
        `Een herinnering: op ${v.date || '<datum>'} verhuist je mailbox naar de nieuwe Microsoft 365-omgeving.\n\nJouw checklist voor de migratiedag:\n1. Kopieer je e-mailhandtekening naar een Word-document\n2. Sluit Outlook aan het einde van de dag af\n3. Na het go-live bericht: herstart je computer en meld je aan volgens de instructies\n4. Stel je telefoon opnieuw in met de nieuwe accountgegevens\n\nDe servicedesk is de dagen na de migratie extra bemand om je snel te helpen.`,
    },
  },
  'cutover-done': {
    en: {
      subject: () => `Your mailbox has been migrated — you can start working`,
      body: (v, tone) =>
        `The migration to the new Microsoft 365 environment is complete. Your mailbox is now live in the new environment.\n\nGetting started:\n1. Restart your computer\n2. Open Outlook — if it asks for an account, follow the setup guide\n3. Reconfigure email on your phone (remove the old account, add the new one)\n4. Recreate your email signature\n\nKnown points of attention:\n- Old Teams meeting links no longer work — recurring meetings are being recreated\n- Check that your shared mailboxes are visible (report it if not)\n${tone === 'detailed' ? '\nAll your historical mail, calendar items and contacts have been migrated. If you miss anything, contact the service desk with the folder name and approximate date.\n' : ''}\nQuestions or issues? The service desk is ready for you${v.ticketId ? ` (reference: ${v.ticketId})` : ''}.`,
    },
    nl: {
      subject: () => `Je mailbox is gemigreerd — je kunt aan de slag`,
      body: (v, tone) =>
        `De migratie naar de nieuwe Microsoft 365-omgeving is afgerond. Je mailbox is nu live in de nieuwe omgeving.\n\nAan de slag:\n1. Herstart je computer\n2. Open Outlook — vraagt deze om een account, volg dan de instructiegids\n3. Stel e-mail op je telefoon opnieuw in (oude account verwijderen, nieuwe toevoegen)\n4. Maak je e-mailhandtekening opnieuw aan\n\nBekende aandachtspunten:\n- Oude Teams-vergaderlinks werken niet meer — terugkerende vergaderingen worden opnieuw aangemaakt\n- Controleer of je gedeelde mailboxen zichtbaar zijn (meld het als dit niet zo is)\n${tone === 'detailed' ? '\nAl je historische mail, agenda-items en contacten zijn gemigreerd. Mis je iets, neem dan contact op met de servicedesk met de mapnaam en geschatte datum.\n' : ''}\nVragen of problemen? De servicedesk staat voor je klaar${v.ticketId ? ` (referentie: ${v.ticketId})` : ''}.`,
    },
  },
  'post-migration': {
    en: {
      subject: () => `How is the new environment working for you? — extra support this week`,
      body: (v) =>
        `Last week your mailbox was migrated to the new Microsoft 365 environment. We hope everything is working smoothly.\n\nThis week we offer extended support (hypercare):\n- Priority handling of migration-related tickets\n- Help with phone setup, signatures and shared mailboxes\n\nMost common quick fixes:\n- Outlook slow/asking for password: restart once more\n- Missing shared mailbox: report it, we restore the permission\n- Phone not syncing: remove and re-add the account\n\nAnything not working as expected? Let us know — we will fix it quickly.`,
    },
    nl: {
      subject: () => `Hoe bevalt de nieuwe omgeving? — extra ondersteuning deze week`,
      body: (v) =>
        `Vorige week is je mailbox gemigreerd naar de nieuwe Microsoft 365-omgeving. We hopen dat alles soepel werkt.\n\nDeze week bieden we verlengde ondersteuning (hypercare):\n- Prioriteit voor migratie-gerelateerde tickets\n- Hulp bij telefooninstellingen, handtekeningen en gedeelde mailboxen\n\nMeest voorkomende snelle oplossingen:\n- Outlook traag/vraagt om wachtwoord: nog één keer herstarten\n- Gedeelde mailbox ontbreekt: meld het, wij herstellen de rechten\n- Telefoon synchroniseert niet: account verwijderen en opnieuw toevoegen\n\nWerkt iets niet zoals verwacht? Laat het ons weten — we lossen het snel op.`,
    },
  },
  'security-recommendation': {
    en: {
      subject: (v) => `Security recommendation for ${v.customer || 'your organization'}: ${v.issue || '<topic>'}`,
      body: (v, tone) =>
        `During our regular security review of your Microsoft 365 environment we identified an improvement opportunity:\n\nFinding: ${v.issue || '<finding>'}\n\nRisk if not addressed: this configuration leaves the environment more vulnerable to account compromise and data exposure.\n\nOur recommendation: <recommended setting/change>\nImpact for users: <impact, e.g. one-time MFA registration>\nEffort: <estimate>\n\n${tone === 'detailed' ? 'We propose to plan this change in a maintenance window, communicate it to affected users in advance, and verify the result together afterwards. A rollback plan is part of the change.\n\n' : ''}Shall we schedule this? We are happy to walk you through the details.`,
    },
    nl: {
      subject: (v) => `Security-aanbeveling voor ${v.customer || 'jouw organisatie'}: ${v.issue || '<onderwerp>'}`,
      body: (v, tone) =>
        `Tijdens onze periodieke security-review van jullie Microsoft 365-omgeving hebben we een verbeterpunt geconstateerd:\n\nBevinding: ${v.issue || '<bevinding>'}\n\nRisico bij niet oppakken: deze configuratie maakt de omgeving kwetsbaarder voor accountovername en datalekken.\n\nOns advies: <aanbevolen instelling/wijziging>\nImpact voor gebruikers: <impact, bijv. eenmalige MFA-registratie>\nInspanning: <inschatting>\n\n${tone === 'detailed' ? 'We stellen voor deze wijziging in een onderhoudsvenster te plannen, vooraf te communiceren naar betrokken gebruikers en het resultaat samen te verifiëren. Een rollback-plan maakt deel uit van de change.\n\n' : ''}Zullen we dit inplannen? We lichten de details graag toe.`,
    },
  },
};

export function generateMail(type: MailType, lang: Lang, tone: Tone, vars: MailVars): { subject: string; body: string } {
  const tpl = bodies[type][lang];
  const subject = tpl.subject(vars);
  const body = `${greetings[lang][tone](vars)}\n\n${tpl.body(vars, tone)}\n\n${closings[lang][tone](vars)}`;
  return { subject, body };
}
