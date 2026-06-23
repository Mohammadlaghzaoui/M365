import { load, save } from '../store/useLocalStorage';

/**
 * Cross-device persistence: syncs the portal's localStorage to a server-side
 * store (api/store.php on one.com) so any browser/device you sign in from sees
 * the same source tenants and assessments.
 *
 * Only assessment/config data is synced — never the login session or any token.
 */

const PREFIX = 'workpilot:';
const STORE_URL = new URL('api/store.php', document.baseURI).toString();

// Keys that must NEVER leave this browser.
const EXCLUDE = new Set(['session', 'cloud-sync']);

export interface CloudSyncConfig {
  enabled: boolean;
  key: string;     // shared secret matching store.php
  user: string;    // namespace (the signed-in e-mail)
}

export function getCloudSync(): CloudSyncConfig {
  return load<CloudSyncConfig>('cloud-sync', { enabled: false, key: '', user: '' });
}
export function saveCloudSync(c: CloudSyncConfig): void {
  save('cloud-sync', c);
}
export function cloudSyncEnabled(): boolean {
  const c = getCloudSync();
  return c.enabled && !!c.key;
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

// A browser must successfully PULL before it is allowed to PUSH, otherwise an
// empty/second browser would overwrite the server and wipe everyone's data.
let hydrated = false;
export function isHydrated(): boolean { return hydrated; }

/** Pull the server blob and hydrate localStorage. Returns number of keys loaded. */
export async function pullCloud(): Promise<number> {
  const c = getCloudSync();
  if (!c.enabled || !c.key) return 0;
  const res = await fetch(STORE_URL, { method: 'GET', headers: { 'x-store-key': c.key, 'x-user': c.user || 'shared' } });
  if (!res.ok) throw new Error(`Cloud sync GET ${res.status}`);
  const blob = (await res.json()) as Record<string, string>;
  const isEmpty = (s: string) => /^(\[\]|\{\}|null|""|)$/.test((s || '').trim());
  let n = 0;
  for (const [k, raw] of Object.entries(blob)) {
    if (EXCLUDE.has(k) || typeof raw !== 'string') continue;
    // Don't overwrite real local data with an empty server value (lets a browser
    // that still has the data re-push it after another browser wiped the server).
    const local = localStorage.getItem(PREFIX + k);
    if (isEmpty(raw) && local && !isEmpty(local)) continue;
    localStorage.setItem(PREFIX + k, raw);
    n++;
  }
  hydrated = true; // a successful read (even of an empty store) makes pushing safe
  return n;
}

/** Push all syncable localStorage keys to the server. Refuses until hydrated (unless forced). */
export async function pushCloud(force = false): Promise<void> {
  const c = getCloudSync();
  if (!c.enabled || !c.key) return;
  if (!hydrated && !force) return; // never clobber the server before a successful pull
  const blob: Record<string, string> = {};
  for (const k of syncableKeys()) {
    const raw = localStorage.getItem(PREFIX + k);
    if (raw != null) blob[k] = raw;
  }
  const res = await fetch(STORE_URL, {
    method: 'POST',
    headers: { 'content-type': 'application/json', 'x-store-key': c.key, 'x-user': c.user || 'shared' },
    body: JSON.stringify(blob),
  });
  if (!res.ok) throw new Error(`Cloud sync POST ${res.status}`);
}

let timer: ReturnType<typeof setTimeout> | null = null;
/** Debounced push — called after each local save (only once this browser is hydrated). */
export function schedulePush(): void {
  if (!cloudSyncEnabled() || !hydrated) return;
  if (timer) clearTimeout(timer);
  timer = setTimeout(() => { pushCloud().catch(() => {}); }, 1200);
}

/** Quick connectivity check for the Settings UI — PULL first (never lose server data), then push the merge. */
export async function testCloudSync(c: CloudSyncConfig): Promise<{ ok: boolean; detail: string }> {
  try {
    const res = await fetch(STORE_URL, { method: 'GET', headers: { 'x-store-key': c.key, 'x-user': c.user || 'shared' } });
    if (res.status === 401) return { ok: false, detail: 'Wrong key — it must match WORKPILOT_STORE_KEY in store.php.' };
    if (res.status === 404) return { ok: false, detail: 'store.php not found — upload public/api/store.php to one.com.' };
    if (!res.ok) return { ok: false, detail: `Server returned ${res.status}.` };
    const remote = (await res.json()) as Record<string, string>;
    const remoteKeys = Object.keys(remote).filter((k) => !EXCLUDE.has(k)).length;
    // Bring the server data into this browser first, then push the merged result.
    await pullCloud();
    await pushCloud(true);
    return { ok: true, detail: remoteKeys ? `Connected — pulled ${remoteKeys} item(s) from the server and synced.` : 'Connected — this browser is now the source; data will sync to others.' };
  } catch (e) {
    return { ok: false, detail: e instanceof Error ? e.message : String(e) };
  }
}
