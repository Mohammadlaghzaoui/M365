import express from 'express';
import cors from 'cors';
import path from 'node:path';
import fs from 'node:fs';
import https from 'node:https';
import { fileURLToPath } from 'node:url';
import { config } from './config.js';
import { runPowerShell, psAvailable } from './powershell.js';
import { graphRequest, graphConfigured } from './graph.js';
import { jobs, createJob, runJob } from './jobs.js';
import { getBatchStatus, exoConfigured } from './exchange.js';
import { audit, readAudit } from './audit.js';

const app = express();
app.disable('x-powered-by');
app.use(express.json({ limit: '2mb' }));
app.use(cors({ origin: config.allowedOrigins.length ? config.allowedOrigins : true }));

// Basic security headers (defence in depth; the agent is an API + optional SPA host).
app.use((_req, res, next) => {
  res.set('X-Content-Type-Options', 'nosniff');
  res.set('X-Frame-Options', 'DENY');
  res.set('Referrer-Policy', 'no-referrer');
  next();
});

// Simple in-memory rate limiter for failed auth (per IP) — slows key brute force.
const authFails = new Map();
function tooManyFails(ip) {
  const rec = authFails.get(ip);
  if (!rec) return false;
  if (Date.now() > rec.until) { authFails.delete(ip); return false; }
  return rec.count >= 10;
}
function recordFail(ip) {
  const rec = authFails.get(ip) ?? { count: 0, until: 0 };
  rec.count += 1;
  rec.until = Date.now() + 60_000; // 1-minute rolling window
  authFails.set(ip, rec);
}

// ---- Static portal hosting (all-in-one server mode) ----
// If agent/public contains the built portal (npm run build:server in the repo
// root), the agent serves it: one server = portal + API + real execution.
const __dirname = path.dirname(fileURLToPath(import.meta.url));
const publicDir = process.env.PUBLIC_DIR || path.join(__dirname, '..', 'public');
const servesPortal = fs.existsSync(path.join(publicDir, 'index.html'));
if (servesPortal) {
  app.use(express.static(publicDir));
}

const API_PREFIXES = ['/health', '/endpoints', '/provision', '/migrate', '/powershell', '/jobs', '/audit'];
const isApi = (p) => API_PREFIXES.some((x) => p === x || p.startsWith(x + '/'));

// ---- API key auth + RBAC (every API route except /health) ----
// Role is resolved from the KEY server-side — a tampered browser cannot escalate.
const ROLE_RANK = { read_only: 0, engineer: 1, architect: 2, super_admin: 3 };
const minRoleFor = (method, p) => {
  if (method === 'GET') return 'read_only';                       // health, jobs, audit gated below
  if (p.startsWith('/powershell')) return 'super_admin';          // raw PS = highest bar
  if (p === '/migrate/start' || p.startsWith('/provision')) return 'engineer'; // real changes
  if (p === '/migrate/test' || p === '/migrate/status' || p.startsWith('/endpoints')) return 'engineer';
  return 'super_admin';
};

const JOB_PATHS = (p) => p === '/jobs' || p.startsWith('/jobs/');

app.use((req, res, next) => {
  if (!isApi(req.path) || req.path === '/health') return next();
  if (tooManyFails(req.ip)) {
    return res.status(429).json({ error: 'Too many failed attempts. Try again in a minute.' });
  }
  const key = req.header('x-api-key') ?? '';
  const role = config.keys.get(key);
  if (!role) {
    recordFail(req.ip);
    audit({ actor: 'unknown', role: null, method: req.method, path: req.path, result: 'denied:bad-key', ip: req.ip });
    return res.status(401).json({ error: 'Invalid or missing X-API-Key.' });
  }
  // Job logs can contain operational detail — require engineer+ to read them.
  const needed = req.path === '/audit' ? 'architect'
    : (req.method === 'GET' && JOB_PATHS(req.path)) ? 'engineer'
    : minRoleFor(req.method, req.path);
  if (ROLE_RANK[role] < ROLE_RANK[needed]) {
    audit({ actor: key.slice(0, 6) + '…', role, method: req.method, path: req.path, result: `denied:requires-${needed}`, ip: req.ip });
    return res.status(403).json({ error: `Forbidden — this action requires the ${needed} role (your key has ${role}).` });
  }
  req.agentRole = role;
  req.actor = key.slice(0, 6) + '…';
  // Audit every state-changing call (GETs are noise; denials logged above).
  if (req.method !== 'GET') {
    audit({ actor: req.actor, role, method: req.method, path: req.path, result: 'allowed', operator: req.header('x-operator') ?? '', ip: req.ip });
  }
  next();
});

// ---- Audit log (architect/super_admin) ----
app.get('/audit', (req, res) => {
  const key = req.header('x-api-key') ?? '';
  const role = config.keys.get(key);
  if (!role || ROLE_RANK[role] < ROLE_RANK.architect) {
    return res.status(403).json({ error: 'Audit log requires architect or super_admin.' });
  }
  res.json({ entries: readAudit(200) });
});

// ---- Health / capabilities ----
app.get('/health', async (_req, res) => {
  res.json({
    ok: true,
    name: 'WorkPilot Migration Agent',
    version: '1.0.0',
    host: config.hostname,
    capabilities: {
      powershell: await psAvailable(),
      graph: graphConfigured(),
      exchangeOnline: exoConfigured(),
      modules: config.declaredModules,
    },
    time: new Date().toISOString(),
  });
});

// ---- Verify an endpoint / connectivity ----
app.post('/endpoints/verify', async (req, res) => {
  const { type } = req.body ?? {};
  try {
    if (type === 'graph' || type === 'Microsoft 365') {
      const me = await graphRequest('GET', '/organization?$select=displayName,id');
      return res.json({ verified: true, detail: `Graph reachable — tenant ${me.value?.[0]?.displayName ?? 'unknown'}` });
    }
    if (type === 'onprem-ad') {
      const out = await runPowerShell('Get-ADDomain | Select-Object -ExpandProperty DNSRoot');
      return res.json({ verified: true, detail: `On-prem AD reachable — ${out.trim()}` });
    }
    if (type === 'exchange-onprem') {
      const out = await runPowerShell('Get-ExchangeServer | Select-Object -First 1 -ExpandProperty Name');
      return res.json({ verified: true, detail: `Exchange on-prem reachable — ${out.trim()}` });
    }
    return res.status(400).json({ verified: false, error: `Unknown endpoint type "${type}".` });
  } catch (e) {
    res.status(502).json({ verified: false, error: String(e.message ?? e) });
  }
});

// ---- Provision a user (on-prem AD via PowerShell, or cloud via Graph) ----
app.post('/provision/user', async (req, res) => {
  const job = createJob('provision-user', req.body);
  res.status(202).json({ jobId: job.id });
  runJob(job).catch(() => {});
});

// ---- Start a mailbox migration batch ----
app.post('/migrate/start', async (req, res) => {
  const job = createJob('migrate-batch', req.body);
  res.status(202).json({ jobId: job.id });
  runJob(job).catch(() => {});
});

// ---- Test migration (validate endpoint + recipients, surface errors) ----
app.post('/migrate/test', async (req, res) => {
  const job = createJob('migrate-test', req.body);
  res.status(202).json({ jobId: job.id });
  runJob(job).catch(() => {});
});

// ---- Poll live per-user migration statistics for a batch ----
app.post('/migrate/status', async (req, res) => {
  const { batchName } = req.body ?? {};
  if (!batchName) return res.status(400).json({ error: 'batchName required.' });
  try {
    const stats = await getBatchStatus(batchName);
    res.json({ batchName, users: stats });
  } catch (e) {
    res.status(502).json({ error: String(e.message ?? e) });
  }
});

// ---- Run an arbitrary, allow-listed PowerShell task (operator-authored) ----
app.post('/powershell/run', async (req, res) => {
  if (!config.allowRawPowerShell) {
    return res.status(403).json({ error: 'Raw PowerShell execution is disabled. Set ALLOW_RAW_POWERSHELL=true to enable on a trusted agent host.' });
  }
  const job = createJob('powershell', req.body);
  res.status(202).json({ jobId: job.id });
  runJob(job).catch(() => {});
});

// ---- Job status + live log ----
app.get('/jobs/:id', (req, res) => {
  const job = jobs.get(req.params.id);
  if (!job) return res.status(404).json({ error: 'Job not found.' });
  res.json(publicJob(job));
});

// Server-Sent Events stream for live log lines.
app.get('/jobs/:id/stream', (req, res) => {
  const job = jobs.get(req.params.id);
  if (!job) return res.status(404).end();
  res.writeHead(200, {
    'Content-Type': 'text/event-stream',
    'Cache-Control': 'no-cache',
    Connection: 'keep-alive',
  });
  let sent = 0;
  const tick = setInterval(() => {
    for (; sent < job.log.length; sent++) {
      res.write(`data: ${JSON.stringify(job.log[sent])}\n\n`);
    }
    if (job.status === 'completed' || job.status === 'failed') {
      res.write(`event: done\ndata: ${JSON.stringify({ status: job.status })}\n\n`);
      clearInterval(tick);
      res.end();
    }
  }, 400);
  req.on('close', () => clearInterval(tick));
});

app.get('/jobs', (_req, res) => {
  res.json([...jobs.values()].slice(-50).map(publicJob));
});

function publicJob(job) {
  return {
    id: job.id, type: job.type, status: job.status,
    progress: job.progress, log: job.log, result: job.result, error: job.error,
    startedAt: job.startedAt, finishedAt: job.finishedAt,
  };
}

// SPA fallback: anything that's not an API route serves the portal shell.
if (servesPortal) {
  app.get('*', (req, res, next) => {
    if (isApi(req.path)) return next();
    res.sendFile(path.join(publicDir, 'index.html'));
  });
}

// Bind to loopback by default — the agent is exposed deliberately, not by accident.
// Set HOST=0.0.0.0 to listen on all interfaces (only behind HTTPS/VPN).
const host = process.env.HOST || '127.0.0.1';
const networked = host !== '127.0.0.1' && host !== 'localhost';

// Refuse to start exposed to the network with weak/no keys.
const weakKeys = [...config.keys.keys()].filter((k) => k.length < 20 || /^(workpilot-local-key|change-me)/i.test(k));
if (config.keys.size === 0) {
  console.error('  [FATAL] No API key configured. Set AGENT_API_KEY (and optionally AGENT_KEYS).');
  process.exit(1);
}
if (networked && weakKeys.length) {
  console.error('  [FATAL] Refusing to bind to the network with a weak/default API key.');
  console.error('          Use a random key of 20+ characters (e.g. `openssl rand -hex 24`).');
  process.exit(1);
}
if (networked && !config.allowedOrigins.length) {
  console.warn('  [warn] Listening on the network with no ALLOWED_ORIGINS restriction — set it to your portal URL.');
}

// Optional HTTPS (SSL/TLS): set TLS_CERT_FILE + TLS_KEY_FILE (PEM). When present
// the agent serves HTTPS directly — recommended for any networked deployment.
let server = app;
let scheme = 'http';
if (process.env.TLS_CERT_FILE && process.env.TLS_KEY_FILE) {
  try {
    const opts = {
      cert: fs.readFileSync(process.env.TLS_CERT_FILE),
      key: fs.readFileSync(process.env.TLS_KEY_FILE),
      ...(process.env.TLS_CA_FILE ? { ca: fs.readFileSync(process.env.TLS_CA_FILE) } : {}),
    };
    server = https.createServer(opts, app);
    scheme = 'https';
  } catch (e) {
    console.error(`  [FATAL] Could not load TLS certificate/key: ${e.message}`);
    process.exit(1);
  }
} else if (networked) {
  console.warn('  [warn] Listening on the network over plain HTTP. Set TLS_CERT_FILE/TLS_KEY_FILE for SSL, or run behind an HTTPS reverse proxy.');
}

server.listen(config.port, host, () => {
  if (servesPortal) console.log(`\n  Portal UI served from ${publicDir}`);
  console.log(`\n  WorkPilot Migration Agent listening on ${scheme}://${host}:${config.port}`);
  console.log(`  Host: ${config.hostname}  (bind: ${host}, TLS: ${scheme === 'https' ? 'ON' : 'off'})`);
  console.log(`  Keys/roles: ${[...config.keys.values()].join(', ') || 'none'}`);
  console.log(`  PowerShell: ${config.declaredModules.join(', ') || '(detected at runtime)'}`);
  console.log(`  Graph app-auth: ${graphConfigured() ? 'configured' : 'NOT configured'}`);
  console.log(`  Raw PowerShell: ${config.allowRawPowerShell ? 'ENABLED (super_admin only)' : 'disabled'}`);
  console.log(`  Allowed portal origins: ${config.allowedOrigins.join(', ') || '(any)'}\n`);
});
