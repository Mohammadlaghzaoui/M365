import { load, save } from '../store/useLocalStorage';

/**
 * Local authentication for the portal (first version, browser-only).
 * Default account is seeded for sorrento.cloud; the password can be changed in
 * Settings. NOTE: client-side auth is a usability gate, not a security
 * boundary — move to Entra SSO (already integrated) or a backend for real
 * enterprise auth.
 */

export interface Session {
  email: string;
  signedInAt: string;
  via: 'local' | 'microsoft';
}

interface LocalUser {
  email: string;
  hash: string;
}

const DEFAULT_EMAIL = 'admin@sorrento.cloud';

export async function computeHash(text: string): Promise<string> {
  const data = new TextEncoder().encode(text);
  const digest = await crypto.subtle.digest('SHA-256', data);
  return Array.from(new Uint8Array(digest)).map((b) => b.toString(16).padStart(2, '0')).join('');
}

async function getUsers(): Promise<LocalUser[]> {
  const stored = load<LocalUser[]>('local-users', []);
  if (stored.length) return stored;
  const seeded = [{ email: DEFAULT_EMAIL, hash: await computeHash('811018') }];
  save('local-users', seeded);
  return seeded;
}

export async function login(email: string, password: string): Promise<Session> {
  const users = await getUsers();
  const hash = await computeHash(password);
  const user = users.find((u) => u.email.toLowerCase() === email.trim().toLowerCase());
  if (!user || user.hash !== hash) throw new Error('Invalid email or password.');
  const session: Session = { email: user.email, signedInAt: new Date().toISOString(), via: 'local' };
  save('session', session);
  return session;
}

export function microsoftSession(email: string): Session {
  const session: Session = { email, signedInAt: new Date().toISOString(), via: 'microsoft' };
  save('session', session);
  return session;
}

export function getSession(): Session | null {
  return load<Session | null>('session', null);
}

export function logout(): void {
  save('session', null);
}

export async function changePassword(email: string, currentPassword: string, newPassword: string): Promise<void> {
  if (newPassword.length < 6) throw new Error('New password must be at least 6 characters.');
  const users = await getUsers();
  const hash = await computeHash(currentPassword);
  const idx = users.findIndex((u) => u.email.toLowerCase() === email.toLowerCase());
  if (idx < 0 || users[idx].hash !== hash) throw new Error('Current password is incorrect.');
  users[idx] = { ...users[idx], hash: await computeHash(newPassword) };
  save('local-users', users);
}

export async function addUser(email: string, password: string): Promise<void> {
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('Invalid email address.');
  if (password.length < 6) throw new Error('Password must be at least 6 characters.');
  const users = await getUsers();
  if (users.some((u) => u.email.toLowerCase() === email.toLowerCase())) throw new Error('User already exists.');
  users.push({ email, hash: await computeHash(password) });
  save('local-users', users);
}

export function removeUser(email: string): void {
  const users = load<LocalUser[]>('local-users', []);
  if (users.length <= 1) throw new Error('Cannot remove the last user.');
  save('local-users', users.filter((u) => u.email.toLowerCase() !== email.toLowerCase()));
}

export function listUsers(): string[] {
  return load<LocalUser[]>('local-users', []).map((u) => u.email);
}
