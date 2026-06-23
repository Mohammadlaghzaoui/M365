import { useEffect, useMemo, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { Layers, Building2, Users2, HardDrive, Smartphone, Monitor, Trash2, ArrowRight, RefreshCw, Plus, ShieldCheck, Server } from 'lucide-react';
import { Badge, Button, Card, PageHeader, ProgressBar } from '../components/ui';
import { tenantIndex, removeTenantResult, loadTenantResult, TenantIndexEntry } from '../services/tenantStore';
import { save } from '../store/useLocalStorage';
import { pullCloud, pushCloud } from '../services/cloudStore';

/** Multi-tenant assessment dashboard — every tenant you connect stays here until you remove it. */
export default function TenantPortfolio() {
  const nav = useNavigate();
  const [tenants, setTenants] = useState<TenantIndexEntry[]>(() => tenantIndex());

  // Refresh from the server on open so changes/deletes from other browsers show up.
  useEffect(() => { pullCloud().then(() => setTenants(tenantIndex())).catch(() => {}); }, []);

  const totals = useMemo(() => tenants.reduce((a, t) => ({
    users: a.users + (t.users || 0),
    devices: a.devices + (t.devices || 0),
    dataGB: a.dataGB + (t.dataGB || 0),
    desktop: a.desktop + (t.desktopUsers || 0),
    mobile: a.mobile + (t.mobileUsers || 0),
  }), { users: 0, devices: 0, dataGB: 0, desktop: 0, mobile: 0 }), [tenants]);

  const open = (t: TenantIndexEntry) => nav(`/tenants/${t.tenantId}`);
  const reconnect = (t: TenantIndexEntry) => {
    const r = loadTenantResult(t.tenantId);
    if (r) save('discovery-result', r);
    nav('/tenants/connect');
  };

  const remove = (t: TenantIndexEntry) => {
    if (!window.confirm(`Remove "${t.displayName}"? This deletes its assessment everywhere you're signed in.`)) return;
    removeTenantResult(t.tenantId);
    setTenants(tenantIndex());
    pushCloud(true).catch(() => {}); // propagate the deletion to your other browsers now
  };

  const fmtData = (gb: number) => (gb >= 1024 ? `${(gb / 1024).toFixed(1)} TB` : `${gb} GB`);
  const relative = (iso: string) => {
    if (!iso) return '—';
    const days = Math.floor((Date.now() - new Date(iso).getTime()) / 86400000);
    return days <= 0 ? 'today' : days === 1 ? 'yesterday' : `${days} days ago`;
  };
  const barColor = (p: number) => (p >= 75 ? 'bg-emerald-600' : p >= 40 ? 'bg-amber-500' : 'bg-rose-500');
  const initials = (n: string) => n.slice(0, 2).toUpperCase();

  return (
    <div>
      <PageHeader title="Source Tenants" subtitle="The Microsoft 365 source tenants you are assessing for migration. Each stays here with its assessment until you remove it." icon={<Layers size={20} />} />

      {/* Portfolio KPIs */}
      <div className="mb-6 grid grid-cols-2 gap-3 lg:grid-cols-5">
        {[
          { icon: Building2, label: 'Source tenants', value: tenants.length, color: 'text-slate-600 dark:text-slate-300' },
          { icon: Users2, label: 'Total users', value: totals.users.toLocaleString(), color: 'text-slate-600 dark:text-slate-300' },
          { icon: Server, label: 'Devices', value: totals.devices.toLocaleString(), color: 'text-slate-600 dark:text-slate-300' },
          { icon: HardDrive, label: 'Total data', value: fmtData(totals.dataGB), color: 'text-slate-600 dark:text-slate-300' },
          { icon: Monitor, label: 'Office / field', value: `${totals.desktop} / ${totals.mobile}`, color: 'text-slate-600 dark:text-slate-300' },
        ].map((k) => (
          <Card key={k.label} className="p-4">
            <div className="flex items-center gap-3">
              <div className={`rounded-lg bg-slate-100 dark:bg-slate-700/50 p-2 ${k.color}`}><k.icon size={18} /></div>
              <div>
                <div className="text-xl font-bold text-slate-800 dark:text-slate-100">{k.value}</div>
                <div className="text-xs text-slate-400">{k.label}</div>
              </div>
            </div>
          </Card>
        ))}
      </div>

      <div className="mb-4 flex items-center justify-between">
        <h2 className="text-sm font-semibold uppercase tracking-wide text-slate-500 dark:text-slate-400">Source tenants</h2>
        <Button onClick={() => nav('/tenants/connect')}><Plus size={15} /> Connect a source tenant</Button>
      </div>

      {tenants.length === 0 ? (
        <Card className="p-10 text-center">
          <Layers size={40} className="mx-auto mb-3 text-slate-300 dark:text-slate-600" />
          <p className="mb-1 font-semibold text-slate-700 dark:text-slate-200">No source tenants yet</p>
          <p className="mb-4 text-sm text-slate-500 dark:text-slate-400">Connect a customer's source tenant with a code — discovery runs here and the tenant stays in this list until you remove it.</p>
          <Button onClick={() => nav('/tenants/connect')}><Plus size={15} /> Connect a source tenant</Button>
        </Card>
      ) : (
        <div className="grid grid-cols-1 gap-4 lg:grid-cols-2 xl:grid-cols-3">
          {tenants.map((t) => (
            <Card key={t.tenantId} className="flex flex-col p-5">
              <div className="mb-3 flex items-start gap-3">
                <div className="flex h-11 w-11 shrink-0 items-center justify-center rounded bg-slate-800 text-sm font-bold text-white">
                  {initials(t.displayName)}
                </div>
                <div className="min-w-0 flex-1">
                  <div className="truncate font-semibold text-slate-800 dark:text-slate-100" title={t.displayName}>{t.displayName}</div>
                  <div className="truncate font-mono text-[11px] text-slate-400" title={t.tenantId}>{t.tenantId}</div>
                </div>
                <Badge color="green">Connected</Badge>
              </div>

              {/* Progress bar — assessment completeness */}
              <div className="mb-3">
                <div className="mb-1 flex items-center justify-between text-xs">
                  <span className="font-medium text-slate-500 dark:text-slate-400">Assessment progress</span>
                  <span className="font-bold text-slate-700 dark:text-slate-200">{t.completeness ?? 0}%</span>
                </div>
                <ProgressBar value={t.completeness ?? 0} color={barColor(t.completeness ?? 0)} />
              </div>

              {/* Stats */}
              <div className="mb-4 grid grid-cols-3 gap-2 text-center">
                {[
                  { icon: Users2, v: (t.users || 0).toLocaleString(), l: 'Users' },
                  { icon: Server, v: (t.devices || 0).toLocaleString(), l: 'Devices' },
                  { icon: HardDrive, v: fmtData(t.dataGB || 0), l: 'Data' },
                ].map((s) => (
                  <div key={s.l} className="rounded-lg bg-slate-50 dark:bg-slate-700/40 px-2 py-2">
                    <s.icon size={14} className="mx-auto mb-0.5 text-slate-400" />
                    <div className="text-sm font-bold text-slate-700 dark:text-slate-200">{s.v}</div>
                    <div className="text-[10px] text-slate-400">{s.l}</div>
                  </div>
                ))}
              </div>

              <div className="mb-4 flex items-center gap-3 text-xs text-slate-500 dark:text-slate-400">
                <span className="inline-flex items-center gap-1"><Monitor size={13} /> {t.desktopUsers ?? 0} office</span>
                <span className="inline-flex items-center gap-1"><Smartphone size={13} /> {t.mobileUsers ?? 0} field</span>
                <span className="ml-auto">Assessed {relative(t.fetchedAt)}</span>
              </div>

              <div className="mt-auto flex items-center gap-2">
                <Button className="flex-1" onClick={() => open(t)}>Open <ArrowRight size={15} /></Button>
                <button onClick={() => reconnect(t)} title="Re-run / re-connect" className="rounded-lg border border-slate-200 dark:border-slate-600 p-2 text-slate-500 hover:bg-slate-100 dark:hover:bg-slate-700"><RefreshCw size={15} /></button>
                <button onClick={() => remove(t)} title="Remove & wipe" className="rounded-lg border border-rose-200 dark:border-rose-800 p-2 text-rose-500 hover:bg-rose-50 dark:hover:bg-rose-900/30"><Trash2 size={15} /></button>
              </div>
            </Card>
          ))}
        </div>
      )}

      <div className="mt-6 flex items-start gap-2 rounded-xl border border-emerald-200 dark:border-emerald-800 bg-emerald-50 dark:bg-emerald-900/20 p-4 text-sm text-emerald-700 dark:text-emerald-300">
        <ShieldCheck size={16} className="mt-0.5 shrink-0" />
        <span><strong>Read-only & private.</strong> Each tenant's assessment is stored separately in this browser only and never leaves your machine. "Remove &amp; wipe" erases that tenant's copy instantly.</span>
      </div>
    </div>
  );
}
