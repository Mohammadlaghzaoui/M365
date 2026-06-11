import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

/**
 * Server-side audit log: append-only JSON lines on the agent host.
 * This is the tamper-resistant record — it lives outside the browser and
 * captures every API call that could change anything, plus auth failures.
 */

const dir = path.dirname(fileURLToPath(import.meta.url));
const auditFile = process.env.AUDIT_FILE || path.join(dir, '..', 'audit.log');

export function audit(entry) {
  const line = JSON.stringify({ t: new Date().toISOString(), ...entry });
  try {
    fs.appendFileSync(auditFile, line + '\n');
  } catch (e) {
    console.error('[audit] write failed:', e.message);
  }
}

export function readAudit(limit = 200) {
  try {
    if (!fs.existsSync(auditFile)) return [];
    const lines = fs.readFileSync(auditFile, 'utf8').trim().split('\n');
    return lines.slice(-limit).map((l) => {
      try { return JSON.parse(l); } catch { return { t: '', raw: l }; }
    }).reverse();
  } catch {
    return [];
  }
}
