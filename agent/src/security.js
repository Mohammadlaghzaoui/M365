/**
 * Security helpers — input validation + safe PowerShell literal quoting.
 *
 * The agent NEVER executes client-supplied script text through the structured
 * endpoints. Every value that reaches PowerShell is (a) validated against a
 * strict allowlist pattern and (b) wrapped as a single-quoted PowerShell
 * literal with internal quotes doubled. Single-quoted PS strings have exactly
 * one metacharacter (the single quote), so this fully prevents injection.
 */

export class ValidationError extends Error {}

/** Wrap as a safe single-quoted PowerShell literal. */
export function psLit(value) {
  return `'${String(value ?? '').replace(/'/g, "''")}'`;
}

const PATTERNS = {
  email: /^[^@\s'"`;|&$()]{1,128}@[a-z0-9.-]{1,255}\.[a-z]{2,}$/i,
  upn: /^[^@\s'"`;|&$()]{1,128}@[a-z0-9.-]{1,255}\.[a-z]{2,}$/i,
  domain: /^[a-z0-9.-]{1,255}\.[a-z]{2,}$/i,
  guid: /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i,
  // Names: letters, spaces, hyphen, apostrophe, dot — no shell/PS metachars.
  name: /^[\p{L}\p{M} .'\-]{1,128}$/u,
  // sAMAccountName / UPN prefix
  samlike: /^[a-z0-9._\-]{1,64}$/i,
  // Migration batch / endpoint names
  label: /^[a-z0-9 _\-]{1,128}$/i,
  // Distinguished name / OU path
  dn: /^[a-z0-9 ,=._\-]{1,512}$/i,
};

export function validate(value, kind, field) {
  const v = String(value ?? '').trim();
  const re = PATTERNS[kind];
  if (!re) throw new ValidationError(`Unknown validation kind ${kind}`);
  if (!re.test(v)) throw new ValidationError(`Invalid ${field || kind}: "${v.slice(0, 60)}"`);
  return v;
}

/** Validate + quote in one step for direct interpolation into a PS literal. */
export function psSafe(value, kind, field) {
  return psLit(validate(value, kind, field));
}

export function optional(value, kind, field) {
  if (value === undefined || value === null || String(value).trim() === '') return '';
  return validate(value, kind, field);
}
