import { load, save } from '../store/useLocalStorage';

/**
 * Cross-device persistence — fully automatic, linked to the signed-in user.
 *
 * No key, no setup: your source tenants and assessments are saved on your
 * one.com server (api/store.php), namespaced by your login e-mail, and pulled
 * automatically on every browser/device you sign in from. Only assessment data
 * is synced — never the login session or any token.
 */

const PREFIX = 'workpilot:';
const STORE_URL = new URL('api/store.php', document.baseURI).toString();
const EXCLUDE = new Set(['session']);

let currentUser = '';
let hydrated = false; // a browser may only PUSH after a successful PULL

export function setSyncUser(email: string): void {
  currentUser = (email || '').trim().toLowerCase();
  hydrated = false;
}
export function syncEnabled(): boolean { return !!currentUser; }
export function isHydrated(): boolean { return hydrated; }

function headers(extra: Record<string, string> = {}): Record<string, string> {
  return { 'x-user': currentUser || 'shared', ...extra };
}

function syncableKeys(): string[] {
  const out: string[] = [];
  for (let i = 0; i < localStorage.length; i++) {
    const full = localStorage.key(i);
    if (!full || !full.startsWith(PREFIX)) continue;
    const k = full.slice(PREFIX.length);
    if (!EXCLUDE.has(k)) out.push(k);
  }
  return out;
}

/** Pull the server blob and make local match it. Returns number of keys loaded. */
export async function pullCloud(): Promise<number> {
  if (!currentUser) return 0;
  const res = await fetch(STORE_URL, { method: 'GET', headers: headers() });
  if (!res.ok) throw new Error(`Cloud sync GET ${res.status}`);
  const blob = (await res.json()) as Record<string, string>;
  const serverEntries = Object.entries(blob).filter(([k, v]) => !EXCLUDE.has(k) && typeof v === 'string');

  // Completely empty server = first sync or a wiped store: don't delete anything,
  // let this (data-holding) browser push and re-seed the server.
  if (serverEntries.length === 0) { hydrated = true; return 0; }

  // Server has data → it is authoritative. Apply server keys AND remove local
  // synced keys the server no longer has, so deletions propagate to every browser.
  const serverKeys = new Set(serverEntries.map(([k]) => k));
  let n = 0;
  for (const [k, raw] of serverEntries) { localStorage.setItem(PREFIX + k, raw); n++; }
  for (const k of syncableKeys()) { if (!serverKeys.has(k)) localStorage.removeItem(PREFIX + k); }
  hydrated = true;
  return n;
}

/** Push all syncable localStorage keys to the server. Refuses until hydrated (unless forced). */
export async function pushCloud(force = false): Promise<void> {
  if (!currentUser) return;
  if (!hydrated && !force) return;
  const blob: Record<string, string> = {};
  for (const k of syncableKeys()) {
    const raw = localStorage.getItem(PREFIX + k);
    if (raw != null) blob[k] = raw;
  }
  const res = await fetch(STORE_URL, { method: 'POST', headers: headers({ 'content-type': 'application/json' }), body: JSON.stringify(blob) });
  if (!res.ok) throw new Error(`Cloud sync POST ${res.status}`);
}

let timer: ReturnType<typeof setTimeout> | null = null;
/** Debounced push — called after each local save (once this browser is hydrated). */
export function schedulePush(): void {
  if (!currentUser || !hydrated) return;
  if (timer) clearTimeout(timer);
  timer = setTimeout(() => { pushCloud().catch(() => {}); }, 1200);
}

/** Manual pull-then-push for the Settings 'Sync now' button. */
export async function syncNow(): Promise<{ ok: boolean; detail: string }> {
  try {
    const before = (await (await fetch(STORE_URL, { method: 'GET', headers: headers() })).json()) as Record<string, string>;
    const remoteKeys = Object.keys(before || {}).filter((k) => !EXCLUDE.has(k)).length;
    await pullCloud();
    await pushCloud(true);
    return { ok: true, detail: remoteKeys ? `Synced — pulled ${remoteKeys} item(s) from your server.` : 'Synced — this browser is now the source; data will appear on your other browsers.' };
  } catch (e) {
    const m = e instanceof Error ? e.message : String(e);
    if (/404/.test(m)) return { ok: false, detail: 'store.php not found — upload api/store.php to one.com.' };
    return { ok: false, detail: m };
  }
}
