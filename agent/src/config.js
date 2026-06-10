import os from 'node:os';

const bool = (v, d = false) => (v === undefined ? d : /^(1|true|yes)$/i.test(v));

export const config = {
  port: Number(process.env.PORT ?? 8787),
  apiKey: process.env.AGENT_API_KEY ?? '',
  hostname: os.hostname(),
  allowedOrigins: (process.env.ALLOWED_ORIGINS ?? '')
    .split(',')
    .map((s) => s.trim())
    .filter(Boolean),

  // Microsoft Graph app-only (client credentials)
  tenantId: process.env.TENANT_ID ?? '',
  clientId: process.env.CLIENT_ID ?? '',
  clientSecret: process.env.CLIENT_SECRET ?? '',

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
