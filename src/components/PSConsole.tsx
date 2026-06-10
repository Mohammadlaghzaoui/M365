import { useEffect, useRef, useState } from 'react';

/**
 * Animated PowerShell-style console. Plays lines one by one with realistic
 * delays, timestamps, colors and a blinking cursor — used by the migration
 * test runners to show a live execution log.
 */

export interface ConsoleLine {
  text: string;
  type?: 'cmd' | 'info' | 'ok' | 'warn' | 'err' | 'plain' | 'header';
  delay?: number; // ms before this line appears
}

const colors: Record<NonNullable<ConsoleLine['type']>, string> = {
  cmd: 'text-yellow-300',
  info: 'text-cyan-300',
  ok: 'text-emerald-400',
  warn: 'text-amber-300',
  err: 'text-red-400',
  plain: 'text-slate-300',
  header: 'text-violet-300',
};

export function PSConsole({ lines, onDone, title = 'Windows PowerShell — M365 WorkPilot Migration Runner' }: { lines: ConsoleLine[]; onDone?: () => void; title?: string }) {
  const [visible, setVisible] = useState(0);
  const [done, setDone] = useState(false);
  const skipRef = useRef(false);
  const scrollRef = useRef<HTMLDivElement>(null);
  const onDoneRef = useRef(onDone);
  onDoneRef.current = onDone;

  useEffect(() => {
    setVisible(0);
    setDone(false);
    skipRef.current = false;
    let cancelled = false;
    let i = 0;

    const next = () => {
      if (cancelled) return;
      if (i >= lines.length) {
        setDone(true);
        onDoneRef.current?.();
        return;
      }
      i += 1;
      setVisible(i);
      const delay = skipRef.current ? 0 : lines[i - 1]?.delay ?? 120;
      setTimeout(next, delay);
    };
    setTimeout(next, 300);
    return () => { cancelled = true; };
  }, [lines]);

  useEffect(() => {
    scrollRef.current?.scrollTo({ top: scrollRef.current.scrollHeight });
  }, [visible]);

  return (
    <div className="overflow-hidden rounded-xl border border-slate-700 bg-[#012456] shadow-lg">
      {/* Title bar */}
      <div className="flex items-center gap-2 border-b border-slate-700 bg-slate-900 px-3 py-2">
        <span className="h-3 w-3 rounded-full bg-red-500" />
        <span className="h-3 w-3 rounded-full bg-amber-400" />
        <span className="h-3 w-3 rounded-full bg-emerald-500" />
        <span className="ml-2 truncate text-xs font-medium text-slate-300">{title}</span>
        {!done && (
          <button onClick={() => { skipRef.current = true; }} className="ml-auto rounded border border-slate-600 px-2 py-0.5 text-[10px] text-slate-400 hover:text-white">
            Skip animation ≫
          </button>
        )}
      </div>
      {/* Console body */}
      <div ref={scrollRef} className="h-80 overflow-y-auto p-4 font-mono text-[13px] leading-relaxed">
        {lines.slice(0, visible).map((l, i) => (
          <div key={i} className={`whitespace-pre-wrap break-all ${colors[l.type ?? 'plain']}`}>
            {l.type === 'cmd' ? <><span className="text-white">PS C:\WorkPilot&gt; </span>{l.text}</> : l.text}
          </div>
        ))}
        {!done && <span className="inline-block h-4 w-2.5 animate-pulse bg-slate-200 align-middle" />}
        {done && <div className="mt-1 text-white">PS C:\WorkPilot&gt; <span className="inline-block h-4 w-2.5 animate-pulse bg-slate-200 align-middle" /></div>}
      </div>
    </div>
  );
}

const ts = () => new Date().toLocaleTimeString('en-GB');

/** Build a realistic console script from validation results. */
export function buildConsoleScript(opts: {
  command: string;
  connectLines: string[];
  checks: { name: string; status: 'passed' | 'failed' | 'warning'; detail: string }[];
  users?: { email: string; status: 'passed' | 'failed' | 'warning'; exception?: string; detail: string }[];
  summary: string;
  passed: boolean;
}): ConsoleLine[] {
  const out: ConsoleLine[] = [];
  out.push({ text: opts.command, type: 'cmd', delay: 600 });
  out.push({ text: '', delay: 200 });
  for (const c of opts.connectLines) {
    out.push({ text: `[${ts()}] ${c}`, type: 'info', delay: 450 });
  }
  out.push({ text: '', delay: 250 });
  out.push({ text: `[${ts()}] ===== CONFIGURATION VALIDATION (${opts.checks.length} checks) =====`, type: 'header', delay: 350 });
  for (const c of opts.checks) {
    const tag = c.status === 'passed' ? '[ PASS ]' : c.status === 'warning' ? '[ WARN ]' : '[ FAIL ]';
    out.push({
      text: `[${ts()}] ${tag} ${c.name}`,
      type: c.status === 'passed' ? 'ok' : c.status === 'warning' ? 'warn' : 'err',
      delay: 320,
    });
    if (c.status !== 'passed') out.push({ text: `           └─ ${c.detail}`, type: c.status === 'warning' ? 'warn' : 'err', delay: 180 });
  }
  if (opts.users && opts.users.length) {
    out.push({ text: '', delay: 250 });
    out.push({ text: `[${ts()}] ===== SIMULATED BATCH RUN (${opts.users.length} users) =====`, type: 'header', delay: 350 });
    for (const u of opts.users) {
      out.push({ text: `[${ts()}] Processing mailbox ${u.email} ...`, type: 'plain', delay: 380 });
      if (u.status === 'passed') {
        out.push({ text: `           └─ Validation OK — ready for migration`, type: 'ok', delay: 160 });
      } else {
        out.push({ text: `           └─ ${u.exception ? u.exception + ': ' : ''}${u.detail}`, type: u.status === 'warning' ? 'warn' : 'err', delay: 220 });
      }
    }
  }
  out.push({ text: '', delay: 300 });
  out.push({ text: `[${ts()}] ${'='.repeat(60)}`, type: 'header', delay: 250 });
  out.push({ text: `[${ts()}] ${opts.summary}`, type: opts.passed ? 'ok' : 'err', delay: 300 });
  out.push({ text: `[${ts()}] Run log written to migration run history.`, type: 'info', delay: 200 });
  return out;
}
