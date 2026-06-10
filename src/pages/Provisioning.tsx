import { useMemo, useRef, useState } from 'react';
import { UserPlus, PlayCircle, Rocket, CheckCircle2, XCircle, AlertTriangle, Trash2, Plus, ChevronDown, ChevronRight } from 'lucide-react';
import { Badge, Button, Card, CodeBlock, Field, PageHeader, Section, Select, TextOutput } from '../components/ui';
import { PSConsole, ConsoleLine } from '../components/PSConsole';
import { AIHelper } from '../components/AIHelper';
import { uid, useLocalStorage } from '../store/useLocalStorage';
import { getSession } from '../services/auth';
import { getSSOSettings } from '../store/settings';
import { currentAccount } from '../services/sso';
import { AccountType, GroupCatalogueEntry, ProvisioningRequest, ProvisioningStatus } from '../types';
import { accountTypeInfo, buildRequestJson, seedCatalogue, validateRequest } from '../data/provisioning';
import { addToCloudGroup, buildOnPremScript, createMemberUser, duplicateCheck, inviteGuest, verifyManager } from '../services/graphProvisioning';

const ts = () => new Date().toLocaleTimeString('en-GB');

function emptyRequest(operator: string): ProvisioningRequest {
  return {
    requestId: `REQ-${uid().toUpperCase()}`,
    createdAt: new Date().toISOString(),
    requestedBy: operator,
    businessOwner: '',
    managerUpn: '',
    ticketNumber: '',
    accountType: 'Internal',
    firstName: '',
    lastName: '',
    displayName: '',
    mail: '',
    upnPrefix: '',
    upnDomain: 'contoso.com',
    department: '',
    jobTitle: '',
    startDate: '',
    endDate: '',
    reviewDate: '',
    groups: [],
    status: 'Draft',
    statusMessage: '',
  };
}

const statusColor: Record<ProvisioningStatus, string> = {
  Draft: 'gray', Validated: 'blue', AlreadyExists: 'orange', ValidationError: 'red',
  Created: 'green', Invited: 'green', PendingSync: 'orange', ApprovalPending: 'purple', Failed: 'red',
};

export default function Provisioning() {
  const operator = getSession()?.email ?? '';
  const [req, setReq] = useState<ProvisioningRequest>(() => emptyRequest(operator));
  const [catalogue, setCatalogue] = useLocalStorage<GroupCatalogueEntry[]>('group-catalogue', seedCatalogue);
  const [log, setLog] = useLocalStorage<ProvisioningRequest[]>('provisioning-requests', []);
  const [consoleLines, setConsoleLines] = useState<ConsoleLine[] | null>(null);
  const [consoleDone, setConsoleDone] = useState(false);
  const [testPassed, setTestPassed] = useState(false);
  const [statusJson, setStatusJson] = useState('');
  const [prodConfirmed, setProdConfirmed] = useState(false);
  const [prodBusy, setProdBusy] = useState(false);
  const [prodOutput, setProdOutput] = useState<{ title: string; text: string }[]>([]);
  const [showCatalogue, setShowCatalogue] = useState(false);
  const [newEntry, setNewEntry] = useState<GroupCatalogueEntry>({ id: '', displayName: '', source: 'Cloud', allowedAccountTypes: ['Internal'], approvalRequired: false, expiryRequired: false });
  const liveRef = useRef(false);

  const ssoReady = getSSOSettings().enabled && !!getSSOSettings().clientId;
  const set = <K extends keyof ProvisioningRequest>(k: K, v: ProvisioningRequest[K]) => {
    setReq((r) => {
      const next = { ...r, [k]: v };
      if (k === 'firstName' || k === 'lastName') {
        next.displayName = `${next.firstName} ${next.lastName}`.trim();
        next.upnPrefix = next.upnPrefix || `${next.firstName}.${next.lastName}`.toLowerCase().replace(/[^a-z0-9.]/g, '');
      }
      return next;
    });
    setTestPassed(false);
    setProdConfirmed(false);
  };

  const allowedGroups = useMemo(() => catalogue.filter((g) => g.allowedAccountTypes.includes(req.accountType)), [catalogue, req.accountType]);
  const selectedGroups = catalogue.filter((g) => req.groups.includes(g.id));
  const upn = `${req.upnPrefix || '<prefix>'}@${req.upnDomain}`;

  const pushLog = (status: ProvisioningStatus, message: string, objectId?: string) => {
    const entry = { ...req, status, statusMessage: message, targetObjectId: objectId, createdAt: new Date().toISOString() };
    setReq(entry);
    setLog((prev) => [entry, ...prev.filter((p) => p.requestId !== entry.requestId)].slice(0, 100));
  };

  // ---------- TEST RUN: validation + duplicate check, animated console ----------
  const runTest = async () => {
    setProdOutput([]);
    setProdConfirmed(false);
    setConsoleDone(false);
    const issues = validateRequest(req, catalogue);
    const blocking = issues.filter((i) => i.blocking);
    const lines: ConsoleLine[] = [
      { text: `Invoke-UserProvisioning -RequestId "${req.ticketNumber || req.requestId}" -AccountType ${req.accountType} -Mode Test`, type: 'cmd', delay: 500 },
      { text: '', delay: 150 },
      { text: `[${ts()}] Step 1/4 — Field validation (runbook §4) ...`, type: 'header', delay: 400 },
    ];
    if (issues.length === 0) {
      lines.push({ text: `[${ts()}] [ PASS ] All required fields valid for account type ${req.accountType}.`, type: 'ok', delay: 350 });
    } else {
      for (const i of issues) {
        lines.push({ text: `[${ts()}] [ ${i.blocking ? 'FAIL' : 'WARN'} ] ${i.field}: ${i.message}`, type: i.blocking ? 'err' : 'warn', delay: 300 });
      }
    }

    let dupLine: ConsoleLine[] = [];
    let dupFound = false;
    lines.push({ text: '', delay: 150 });
    lines.push({ text: `[${ts()}] Step 2/4 — Duplicate check (UPN, mail, proxyAddresses, otherMails) ...`, type: 'header', delay: 450 });

    liveRef.current = false;
    if (blocking.length === 0 && ssoReady && (await currentAccount())) {
      try {
        const dup = await duplicateCheck(upn, req.mail);
        liveRef.current = true;
        if (dup) {
          dupFound = true;
          dupLine = [
            { text: `[${ts()}] [ FAIL ] Account already exists in Entra ID (live Graph check).`, type: 'err', delay: 350 },
            { text: `           └─ matchedAttribute: ${dup.matchedAttribute} | matchedValue: ${dup.matchedValue}`, type: 'err', delay: 200 },
            { text: `           └─ objectId: ${dup.objectId} (${dup.displayName}, ${dup.userType ?? 'user'})`, type: 'err', delay: 200 },
          ];
          setStatusJson(JSON.stringify({ status: 'AlreadyExists', message: 'Account already exists in On-Premises AD or Microsoft Entra ID.', matchedAttribute: dup.matchedAttribute, matchedValue: dup.matchedValue, objectId: dup.objectId, nextAction: 'Review existing account or update request.' }, null, 2));
        } else {
          dupLine = [{ text: `[${ts()}] [ PASS ] No existing object found on UPN/mail/proxyAddresses (live Graph check).`, type: 'ok', delay: 350 }];
        }
      } catch (e) {
        dupLine = [{ text: `[${ts()}] [ WARN ] Live Graph duplicate check unavailable: ${e instanceof Error ? e.message : e}`, type: 'warn', delay: 350 },
          { text: `           └─ Sign in via Microsoft SSO with provisioning scopes to enable live checks.`, type: 'warn', delay: 200 }];
      }
    } else if (blocking.length === 0) {
      dupLine = [{ text: `[${ts()}] [ WARN ] Simulated — connect Microsoft 365 SSO (Settings) for the live Entra ID duplicate check.`, type: 'warn', delay: 350 }];
    } else {
      dupLine = [{ text: `[${ts()}] [ SKIP ] Duplicate check skipped — fix validation errors first.`, type: 'warn', delay: 300 }];
    }
    lines.push(...dupLine);

    const info = accountTypeInfo[req.accountType];
    lines.push({ text: '', delay: 150 });
    lines.push({ text: `[${ts()}] Step 3/4 — Routing decision (runbook §3/§6) ...`, type: 'header', delay: 400 });
    lines.push({ text: `[${ts()}] Account type ${req.accountType} → ${info.route}`, type: 'info', delay: 400 });
    lines.push({ text: `[${ts()}] Policy rule: ${info.rule}`, type: 'info', delay: 350 });

    lines.push({ text: '', delay: 150 });
    lines.push({ text: `[${ts()}] Step 4/4 — Group catalogue check (${selectedGroups.length} selected) ...`, type: 'header', delay: 400 });
    if (!selectedGroups.length) lines.push({ text: `[${ts()}] [ WARN ] No groups selected.`, type: 'warn', delay: 250 });
    for (const g of selectedGroups) {
      const route = g.source === 'Cloud' ? 'Graph group assignment' : 'on-prem AD assignment (synced/on-prem managed)';
      lines.push({ text: `[${ts()}] [ ${g.approvalRequired ? 'HOLD' : 'PASS'} ] ${g.displayName} (${g.source}) → ${route}${g.approvalRequired ? ' — APPROVAL REQUIRED' : ''}`, type: g.approvalRequired ? 'warn' : 'ok', delay: 300 });
    }

    const passed = blocking.length === 0 && !dupFound;
    lines.push({ text: '', delay: 250 });
    lines.push({ text: `[${ts()}] ${'='.repeat(58)}`, type: 'header', delay: 200 });
    lines.push({
      text: passed
        ? `[${ts()}] TEST PASSED — request is ready for production execution.`
        : `[${ts()}] TEST FAILED — ${dupFound ? 'duplicate found' : blocking.length + ' validation error(s)'}; production stays locked.`,
      type: passed ? 'ok' : 'err', delay: 350,
    });

    if (!dupFound) {
      setStatusJson(JSON.stringify({
        status: passed ? 'Validated' : 'ValidationError',
        message: passed ? 'Request validated — unique in directory, routing resolved.' : 'Validation errors must be resolved.',
        accountType: req.accountType,
        userPrincipalName: upn,
        issues: issues.map((i) => `${i.field}: ${i.message}`),
      }, null, 2));
    }
    setTestPassed(passed);
    setConsoleLines(lines);
    pushLog(passed ? 'Validated' : dupFound ? 'AlreadyExists' : 'ValidationError', passed ? 'Test run passed' : dupFound ? 'Duplicate found in directory' : 'Validation errors');
  };

  // ---------- PRODUCTION ----------
  const runProduction = async () => {
    setProdBusy(true);
    setProdOutput([]);
    const out: { title: string; text: string }[] = [];
    try {
      if (req.accountType === 'Internal') {
        // Hybrid rule: AD-first, never cloud-only. Emit the runbook script + PendingSync.
        out.push({ title: 'On-Premises AD provisioning script (run via delegated service account)', text: buildOnPremScript(req, selectedGroups) });
        out.push({ title: 'Status', text: JSON.stringify({ status: 'PendingSync', message: 'User created on-prem - waiting for Entra sync.', requestId: req.ticketNumber || req.requestId, nextAction: 'Verify object after sync, then assign cloud groups.' }, null, 2) });
        pushLog('PendingSync', 'AD-first script issued — waiting for Entra Connect sync');
      } else if (req.accountType === 'ExternalMember') {
        if (req.managerUpn) await verifyManager(req.managerUpn);
        const created = await createMemberUser(req);
        const results: string[] = [`User created: ${created.upn} (objectId ${created.id})`, `Initial password (share via secure channel, change forced at first sign-in): ${created.initialPassword}`];
        for (const g of selectedGroups) {
          if (g.approvalRequired) { results.push(`Group "${g.displayName}": APPROVAL PENDING — not assigned automatically.`); continue; }
          if (g.source === 'Cloud') {
            try { await addToCloudGroup(created.id, g.id); results.push(`Group "${g.displayName}": assigned via Graph.`); }
            catch (e) { results.push(`Group "${g.displayName}": FAILED — ${e instanceof Error ? e.message : e}`); }
          } else {
            results.push(`Group "${g.displayName}" (${g.source}): assign on-prem — Add-ADGroupMember -Identity "${g.id}" -Members "${req.upnPrefix}"`);
          }
        }
        out.push({ title: 'Execution result', text: results.join('\n') });
        out.push({ title: 'Status', text: JSON.stringify({ status: 'Created', objectId: created.id, userPrincipalName: created.upn, requestId: req.ticketNumber || req.requestId }, null, 2) });
        pushLog(selectedGroups.some((g) => g.approvalRequired) ? 'ApprovalPending' : 'Created', `Member created: ${created.upn}`, created.id);
      } else {
        const invited = await inviteGuest(req, 'https://myapps.microsoft.com');
        const results: string[] = [`Invitation sent to ${req.mail} (guest objectId ${invited.id})`, `Redeem URL: ${invited.inviteRedeemUrl}`];
        for (const g of selectedGroups) {
          if (g.approvalRequired) { results.push(`Group "${g.displayName}": APPROVAL PENDING.`); continue; }
          if (g.source === 'Cloud') {
            try { await addToCloudGroup(invited.id, g.id); results.push(`Group "${g.displayName}": assigned.`); }
            catch (e) { results.push(`Group "${g.displayName}": FAILED — ${e instanceof Error ? e.message : e}`); }
          }
        }
        out.push({ title: 'Execution result', text: results.join('\n') });
        out.push({ title: 'Status', text: JSON.stringify({ status: 'Invited', objectId: invited.id, invitedUserEmailAddress: req.mail, requestId: req.ticketNumber || req.requestId }, null, 2) });
        pushLog('Invited', `Guest invited: ${req.mail}`, invited.id);
      }
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      out.push({ title: 'Execution FAILED', text: `${msg}\n\nNo secrets/tokens are shown. Typical causes: missing delegated scopes (User.ReadWrite.All / User.Invite.All / Group.ReadWrite.All on the SSO app registration + admin consent), insufficient operator role, or policy block.` });
      pushLog('Failed', msg);
    } finally {
      setProdBusy(false);
      setProdOutput(out);
    }
  };

  const newRequest = () => {
    setReq(emptyRequest(operator));
    setConsoleLines(null);
    setStatusJson('');
    setTestPassed(false);
    setProdOutput([]);
    setProdConfirmed(false);
  };

  return (
    <div>
      <PageHeader title="User Provisioning" subtitle="Enterprise account provisioning per runbook: GUI request → validation → duplicate check → routed execution (hybrid AD-first, Graph member creation, or B2B invitation) with approved group catalogue and audit log." icon={<UserPlus size={20} />} />

      {!ssoReady && (
        <Card className="mb-5 p-4 border-amber-200 dark:border-amber-800 text-sm text-amber-700 dark:text-amber-300 flex gap-2">
          <AlertTriangle size={16} className="mt-0.5 shrink-0" />
          <span>Live Entra ID execution needs Microsoft 365 SSO (Settings → Microsoft SSO) with delegated scopes <code>User.ReadWrite.All</code>, <code>User.Invite.All</code>, <code>Group.ReadWrite.All</code> (admin consent). Without it, the module still validates, simulates and generates all scripts. The frontend never stores app secrets — per the runbook security model.</span>
        </Card>
      )}

      <div className="grid gap-6 xl:grid-cols-[460px_1fr]">
        {/* ============ FORM ============ */}
        <div className="space-y-4 self-start">
          <Card className="p-5 space-y-3">
            <Section title="Requester">
              <div className="grid grid-cols-2 gap-3">
                <Field label="Requester (operator)" value={req.requestedBy} onChange={(v) => set('requestedBy', v)} />
                <Field label="Ticket / request number" value={req.ticketNumber} onChange={(v) => set('ticketNumber', v)} placeholder="RITM0012345" />
                <Field label="Business owner" value={req.businessOwner} onChange={(v) => set('businessOwner', v)} />
                <Field label="Manager UPN" value={req.managerUpn} onChange={(v) => set('managerUpn', v)} placeholder="manager@contoso.com" />
              </div>
            </Section>
          </Card>

          <Card className="p-5">
            <Section title="Account type (determines the provisioning route)">
              <div className="space-y-2">
                {(Object.keys(accountTypeInfo) as AccountType[]).map((t) => (
                  <label key={t} className={`block cursor-pointer rounded-lg border p-3 ${req.accountType === t ? 'border-blue-500 bg-blue-50 dark:bg-blue-900/20' : 'border-slate-200 dark:border-slate-700'}`}>
                    <div className="flex items-center gap-2">
                      <input type="radio" checked={req.accountType === t} onChange={() => { set('accountType', t); set('groups', []); }} className="text-blue-600" />
                      <span className="font-semibold text-sm text-slate-700 dark:text-slate-200">{accountTypeInfo[t].label}</span>
                    </div>
                    <div className="mt-1 pl-6 text-xs text-slate-500 dark:text-slate-400">{accountTypeInfo[t].route}</div>
                    <div className="pl-6 text-xs text-amber-600 dark:text-amber-400">{accountTypeInfo[t].rule}</div>
                  </label>
                ))}
              </div>
            </Section>
          </Card>

          <Card className="p-5 space-y-3">
            <Section title="Personal details">
              <div className="grid grid-cols-2 gap-3">
                <Field label="First name" value={req.firstName} onChange={(v) => set('firstName', v)} />
                <Field label="Last name" value={req.lastName} onChange={(v) => set('lastName', v)} />
                <Field label="Display name" value={req.displayName} onChange={(v) => set('displayName', v)} />
                <Field label={req.accountType === 'Guest' ? 'External email (invitation)' : 'Mail address'} value={req.mail} onChange={(v) => set('mail', v)} placeholder={req.accountType === 'Guest' ? 'person@partner.com' : 'first.last@contoso.com'} />
                <Field label="UPN prefix" value={req.upnPrefix} onChange={(v) => set('upnPrefix', v)} />
                <Field label="UPN domain" value={req.upnDomain} onChange={(v) => set('upnDomain', v)} />
                <Field label="Department" value={req.department} onChange={(v) => set('department', v)} />
                <Field label="Job title" value={req.jobTitle} onChange={(v) => set('jobTitle', v)} />
              </div>
              {req.accountType !== 'Guest' && <p className="mt-2 text-xs text-slate-400">Resulting UPN: <span className="font-mono text-blue-500">{upn}</span></p>}
            </Section>
          </Card>

          <Card className="p-5">
            <Section title="Lifecycle">
              <div className="grid grid-cols-3 gap-3">
                <Field label="Start date" type="date" value={req.startDate} onChange={(v) => set('startDate', v)} />
                <Field label={`End date${req.accountType !== 'Internal' ? ' (required)' : ''}`} type="date" value={req.endDate} onChange={(v) => set('endDate', v)} />
                <Field label="Review date" type="date" value={req.reviewDate} onChange={(v) => set('reviewDate', v)} />
              </div>
            </Section>
          </Card>

          <Card className="p-5">
            <Section title={`Groups — approved catalogue (${allowedGroups.length} allowed for ${req.accountType})`}>
              <div className="space-y-2">
                {allowedGroups.map((g) => (
                  <label key={g.id} className="flex cursor-pointer items-start gap-2.5 rounded-lg border border-slate-200 dark:border-slate-700 p-2.5 hover:bg-slate-50 dark:hover:bg-slate-700/40">
                    <input type="checkbox" checked={req.groups.includes(g.id)}
                      onChange={() => set('groups', req.groups.includes(g.id) ? req.groups.filter((x) => x !== g.id) : [...req.groups, g.id])}
                      className="mt-0.5 h-4 w-4 rounded border-slate-300 text-blue-600" />
                    <span className="flex-1">
                      <span className="flex flex-wrap items-center gap-1.5 text-sm font-medium text-slate-700 dark:text-slate-200">
                        {g.displayName}
                        <Badge color={g.source === 'Cloud' ? 'blue' : g.source === 'OnPrem' ? 'orange' : 'purple'}>{g.source}</Badge>
                        {g.approvalRequired && <Badge color="red">approval</Badge>}
                        {g.expiryRequired && <Badge color="orange">expiry</Badge>}
                      </span>
                      {g.description && <span className="block text-xs text-slate-400">{g.description}</span>}
                    </span>
                  </label>
                ))}
                {!allowedGroups.length && <p className="text-sm text-slate-400">No catalogue groups allow account type {req.accountType} — add them below.</p>}
              </div>
              <button onClick={() => setShowCatalogue(!showCatalogue)} className="mt-3 flex items-center gap-1 text-xs font-semibold text-blue-500 hover:underline">
                {showCatalogue ? <ChevronDown size={13} /> : <ChevronRight size={13} />} Manage catalogue
              </button>
              {showCatalogue && (
                <div className="mt-3 space-y-2 rounded-lg border border-slate-200 dark:border-slate-700 p-3">
                  {catalogue.map((g) => (
                    <div key={g.id} className="flex items-center gap-2 text-xs text-slate-600 dark:text-slate-300">
                      <span className="flex-1 truncate">{g.displayName} <span className="text-slate-400">({g.source} · {g.allowedAccountTypes.join('/')})</span></span>
                      <button onClick={() => setCatalogue((prev) => prev.filter((x) => x.id !== g.id))} className="text-slate-400 hover:text-red-500"><Trash2 size={13} /></button>
                    </div>
                  ))}
                  <div className="grid grid-cols-2 gap-2 pt-2">
                    <Field label="Display name" value={newEntry.displayName} onChange={(v) => setNewEntry({ ...newEntry, displayName: v })} />
                    <Field label="Group ID (GUID) or DN" value={newEntry.id} onChange={(v) => setNewEntry({ ...newEntry, id: v })} />
                    <Select label="Source" value={newEntry.source} onChange={(v) => setNewEntry({ ...newEntry, source: v as GroupCatalogueEntry['source'] })} options={['Cloud', 'OnPrem', 'Synced'].map((s) => ({ value: s, label: s }))} />
                    <div className="flex items-end gap-3 pb-1 text-xs text-slate-600 dark:text-slate-300">
                      {(['Internal', 'ExternalMember', 'Guest'] as AccountType[]).map((t) => (
                        <label key={t} className="flex items-center gap-1">
                          <input type="checkbox" checked={newEntry.allowedAccountTypes.includes(t)}
                            onChange={() => setNewEntry({ ...newEntry, allowedAccountTypes: newEntry.allowedAccountTypes.includes(t) ? newEntry.allowedAccountTypes.filter((x) => x !== t) : [...newEntry.allowedAccountTypes, t] })} />
                          {t === 'ExternalMember' ? 'Ext' : t === 'Internal' ? 'Int' : 'Guest'}
                        </label>
                      ))}
                    </div>
                  </div>
                  <div className="flex items-center gap-3 text-xs text-slate-600 dark:text-slate-300">
                    <label className="flex items-center gap-1"><input type="checkbox" checked={newEntry.approvalRequired} onChange={(e) => setNewEntry({ ...newEntry, approvalRequired: e.target.checked })} /> approval required</label>
                    <label className="flex items-center gap-1"><input type="checkbox" checked={newEntry.expiryRequired} onChange={(e) => setNewEntry({ ...newEntry, expiryRequired: e.target.checked })} /> expiry required</label>
                    <Button variant="secondary" className="ml-auto !px-2 !py-1 text-xs" onClick={() => {
                      if (!newEntry.id || !newEntry.displayName) return;
                      setCatalogue((prev) => [...prev, newEntry]);
                      setNewEntry({ id: '', displayName: '', source: 'Cloud', allowedAccountTypes: ['Internal'], approvalRequired: false, expiryRequired: false });
                    }}><Plus size={13} /> Add</Button>
                  </div>
                </div>
              )}
            </Section>
          </Card>

          <div className="flex gap-2">
            <Button variant="ai" onClick={runTest} className="flex-1"><PlayCircle size={16} /> Run test (validate + duplicate check)</Button>
            <Button variant="secondary" onClick={newRequest}>New request</Button>
          </div>
        </div>

        {/* ============ EXECUTION SIDE ============ */}
        <div className="space-y-4">
          {consoleLines && (
            <PSConsole lines={consoleLines} onDone={() => setConsoleDone(true)} title={`Windows PowerShell — User Provisioning Runner (${req.ticketNumber || req.requestId})`} />
          )}
          {consoleDone && statusJson && <TextOutput title="Status (GUI/ServiceNow feedback contract)" text={statusJson} />}
          {consoleDone && <TextOutput title="API payload — POST /api/provisioning/users" text={buildRequestJson(req)} />}

          {/* Production gate */}
          <Card className={`p-5 ${testPassed && consoleDone ? 'border-emerald-300 dark:border-emerald-700' : 'opacity-70'}`}>
            <h3 className="font-bold text-slate-800 dark:text-slate-100 flex items-center gap-2"><Rocket size={18} className="text-emerald-500" /> Production execution</h3>
            {!(testPassed && consoleDone) ? (
              <p className="mt-1 text-sm text-slate-500 dark:text-slate-400">Locked — run the test until it passes (validation + duplicate check).</p>
            ) : (
              <>
                <p className="mt-1 text-sm text-slate-500 dark:text-slate-400">
                  {req.accountType === 'Internal' && 'Internal = hybrid AD-first: generates the AD provisioning script and sets status PendingSync (never cloud-only).'}
                  {req.accountType === 'ExternalMember' && 'Creates the member via Microsoft Graph POST /users with your signed-in operator permissions, then assigns approved cloud groups.'}
                  {req.accountType === 'Guest' && 'Sends the B2B invitation via Microsoft Graph POST /invitations, then assigns approved guest groups.'}
                </p>
                <label className="mt-3 flex items-center gap-2 text-sm font-medium text-slate-700 dark:text-slate-200">
                  <input type="checkbox" checked={prodConfirmed} onChange={(e) => setProdConfirmed(e.target.checked)} className="h-4 w-4 rounded border-slate-300 text-emerald-600" />
                  Request approved (ticket {req.ticketNumber || req.requestId}) — execute
                </label>
                {prodConfirmed && (
                  <div className="mt-3">
                    <Button onClick={runProduction} disabled={prodBusy}>
                      {prodBusy ? 'Executing…' : req.accountType === 'Internal' ? 'Generate AD runbook (PendingSync)' : req.accountType === 'Guest' ? 'Send invitation now' : 'Create account now'}
                    </Button>
                  </div>
                )}
              </>
            )}
            {prodOutput.map((o) => (
              <div key={o.title} className="mt-4">
                {o.text.startsWith('#') || o.text.includes('New-ADUser')
                  ? <><div className="mb-1 text-sm font-semibold text-slate-600 dark:text-slate-300">{o.title}</div><CodeBlock code={o.text} /></>
                  : <TextOutput title={o.title} text={o.text} />}
              </div>
            ))}
          </Card>

          {/* Audit log */}
          {log.length > 0 && (
            <Card className="p-5">
              <Section title={`Request log / audit trail (${log.length})`}>
                <div className="overflow-x-auto">
                  <table className="w-full text-sm">
                    <thead>
                      <tr className="border-b border-slate-200 dark:border-slate-700 text-left text-xs uppercase text-slate-400">
                        <th className="py-2 pr-3">Request</th><th className="py-2 pr-3">Operator</th><th className="py-2 pr-3">Type</th><th className="py-2 pr-3">Target</th><th className="py-2 pr-3">Status</th><th className="py-2">Time</th>
                      </tr>
                    </thead>
                    <tbody>
                      {log.slice(0, 15).map((l) => (
                        <tr key={l.requestId + l.createdAt} className="border-b border-slate-100 dark:border-slate-700/50">
                          <td className="py-2 pr-3 font-mono text-xs text-slate-600 dark:text-slate-300">{l.ticketNumber || l.requestId}</td>
                          <td className="py-2 pr-3 text-slate-500 dark:text-slate-300">{l.requestedBy}</td>
                          <td className="py-2 pr-3 text-slate-500 dark:text-slate-300">{l.accountType}</td>
                          <td className="py-2 pr-3 text-slate-500 dark:text-slate-300">{l.upnPrefix ? `${l.upnPrefix}@${l.upnDomain}` : l.mail}</td>
                          <td className="py-2 pr-3"><Badge color={statusColor[l.status]}>{l.status}</Badge></td>
                          <td className="py-2 text-xs text-slate-400">{new Date(l.createdAt).toLocaleString()}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </Section>
            </Card>
          )}

          <AIHelper
            context={`User provisioning request:\n${buildRequestJson(req)}\nStatus: ${req.status} — ${req.statusMessage}`}
            defaultPrompt="Review this provisioning request against best practices and tell me what could go wrong before I execute it."
          />
        </div>
      </div>
    </div>
  );
}
