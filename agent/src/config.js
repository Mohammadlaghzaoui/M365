import os from 'node:os';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

// Zero-dependency .env loader (reads agent/.env if present).
(() => {
  const dir = path.dirname(fileURLToPath(import.meta.url));
  const envPath = path.join(dir, '..', '.env');
  if (!fs.existsSync(envPath)) return;
  for (const line of fs.readFileSync(envPath, 'utf8').split(/\r?\n/)) {
    const m = line.match(/^\s*([A-Z0-9_]+)\s*=\s*(.*)\s*$/i);
    if (m && process.env[m[1]] === undefined) {
      process.env[m[1]] = m[2].replace(/^["']|["']$/g, '');
    }
  }
})();

const bool = (v, d = false) => (v === undefined ? d : /^(1|true|yes)$/i.test(v));

export const config = {
  port: Number(process.env.PORT ?? 8787),
  apiKey: process.env.AGENT_API_KEY ?? '',
  hostname: os.hostname(),
  allowedOrigins: (process.env.ALLOWED_ORIGINS ?? '')
    .split(',')
    .map((s) => s.trim())
    .filter(Boolean),

  /**
   * RBAC: per-key roles, enforced SERVER-SIDE (the key IS the identity, so a
   * tampered browser cannot escalate). Format:
   *   AGENT_KEYS=longkey1:super_admin,longkey2:engineer,longkey3:read_only
   * AGENT_API_KEY (if set) is implicitly super_admin for backwards compat.
   * Roles: super_admin | architect | engineer | read_only
   */
  keys: (() => {
    const map = new Map();
    for (const pair of (process.env.AGENT_KEYS ?? '').split(',').map((s) => s.trim()).filter(Boolean)) {
      const [key, role] = pair.split(':').map((s) => s.trim());
      if (key && ['super_admin', 'architect', 'engineer', 'read_only'].includes(role)) map.set(key, role);
    }
    if (process.env.AGENT_API_KEY) map.set(process.env.AGENT_API_KEY, 'super_admin');
    return map;
  })(),

  // Microsoft Graph app-only (client credentials)
  tenantId: process.env.TENANT_ID ?? '',
  clientId: process.env.CLIENT_ID ?? '',
  clientSecret: process.env.CLIENT_SECRET ?? '',

  // Exchange Online app-only (certificate) for the Migration Console
  exoAppId: process.env.EXO_APP_ID ?? '',
  exoOrg: process.env.EXO_ORG ?? '',
  exoCertThumbprint: process.env.EXO_CERT_THUMBPRINT ?? '',

  // Safety switches
  allowRawPowerShell: bool(process.env.ALLOW_RAW_POWERSHELL, false),
  powershellExe: process.env.POWERSHELL_EXE ?? (process.platform === 'win32' ? 'powershell.exe' : 'pwsh'),

  // Informational — which on-prem modules the operator says are installed
  declaredModules: (process.env.PS_MODULES ?? '')
    .split(',')
    .map((s) => s.trim())
    .filter(Boolean),
};

if (!config.apiKey) {
  console.warn('  [warn] AGENT_API_KEY is empty — set it (and the same key in the portal) before exposing the agent.');
}
