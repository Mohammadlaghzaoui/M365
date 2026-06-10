// Lightweight pure-SVG charts (no dependencies) for the dashboard.

export interface Segment {
  label: string;
  value: number;
  color: string; // hex or tailwind-compatible fill
}

export function DonutChart({ segments, centerLabel, centerValue }: { segments: Segment[]; centerLabel: string; centerValue: string }) {
  const total = Math.max(1, segments.reduce((a, s) => a + s.value, 0));
  const r = 42;
  const c = 2 * Math.PI * r;
  let offset = 0;
  return (
    <div className="flex items-center gap-5">
      <svg viewBox="0 0 110 110" className="h-36 w-36 -rotate-90">
        <circle cx="55" cy="55" r={r} fill="none" stroke="currentColor" className="text-slate-100 dark:text-slate-700" strokeWidth="14" />
        {segments.map((s) => {
          const len = (s.value / total) * c;
          const el = (
            <circle key={s.label} cx="55" cy="55" r={r} fill="none" stroke={s.color} strokeWidth="14"
              strokeDasharray={`${len} ${c - len}`} strokeDashoffset={-offset} strokeLinecap="butt"
              className="transition-all duration-700" />
          );
          offset += len;
          return el;
        })}
      </svg>
      <div>
        <div className="mb-2">
          <div className="text-2xl font-bold text-slate-800 dark:text-slate-100">{centerValue}</div>
          <div className="text-xs text-slate-400">{centerLabel}</div>
        </div>
        <div className="space-y-1">
          {segments.map((s) => (
            <div key={s.label} className="flex items-center gap-2 text-xs text-slate-600 dark:text-slate-300">
              <span className="h-2.5 w-2.5 rounded-sm" style={{ background: s.color }} />
              {s.label} <span className="font-semibold">{s.value}</span>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

export function HBarChart({ data, color = '#2563eb' }: { data: { label: string; value: number; color?: string }[]; color?: string }) {
  const max = Math.max(1, ...data.map((d) => d.value));
  return (
    <div className="space-y-2.5">
      {data.map((d) => (
        <div key={d.label}>
          <div className="mb-0.5 flex justify-between text-xs">
            <span className="text-slate-600 dark:text-slate-300">{d.label}</span>
            <span className="font-semibold text-slate-700 dark:text-slate-200">{d.value}</span>
          </div>
          <div className="h-2.5 w-full overflow-hidden rounded-full bg-slate-100 dark:bg-slate-700">
            <div className="h-full rounded-full transition-all duration-700" style={{ width: `${(d.value / max) * 100}%`, background: d.color ?? color }} />
          </div>
        </div>
      ))}
      {data.length === 0 && <p className="text-xs text-slate-400">No data yet.</p>}
    </div>
  );
}

export function VBarChart({ data, color = '#7c3aed' }: { data: { label: string; value: number }[]; color?: string }) {
  const max = Math.max(1, ...data.map((d) => d.value));
  return (
    <div className="flex h-36 items-end gap-2">
      {data.map((d) => (
        <div key={d.label} className="flex flex-1 flex-col items-center gap-1">
          <span className="text-xs font-semibold text-slate-600 dark:text-slate-300">{d.value}</span>
          <div className="w-full rounded-t-md transition-all duration-700" style={{ height: `${Math.max(4, (d.value / max) * 100)}%`, background: color, opacity: 0.55 + 0.45 * (d.value / max) }} />
          <span className="truncate text-[10px] text-slate-400 max-w-full">{d.label}</span>
        </div>
      ))}
      {data.length === 0 && <p className="text-xs text-slate-400">No data yet.</p>}
    </div>
  );
}
