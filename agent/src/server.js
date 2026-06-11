import express from 'express';
import cors from 'cors';
import path from 'node:path';
import fs from 'node:fs';
import { fileURLToPath } from 'node:url';
import { config } from './config.js';
import { runPowerShell, psAvailable } from './powershell.js';
import { graphRequest, graphConfigured } from './graph.js';
import { jobs, createJob, runJob } from './jobs.js';
import { getBatchStatus, exoConfigured } from './exchange.js';

const app = express();
app.use(express.json({ limit: '5mb' }));
app.use(cors({ origin: config.allowedOrigins.length ? config.allowedOrigins : true }));

// ---- Static portal hosting (all-in-one server mode) ----
// If agent/public contains the built portal (npm run build:server in the repo
// root), the agent serves it: one server = portal + API + real execution.
const __dirname = path.dirname(fileURLToPath(import.meta.url));
const publicDir = process.env.PUBLIC_DIR || path.join(__dirname, '..', 'public');
const servesPortal = fs.existsSync(path.join(publicDir, 'index.html'));
if (servesPortal) {
  app.use(express.static(publicDir));
}

const API_PREFIXES = ['/health', '/endpoints', '/provision', '/migrate', '/powershell', '/jobs'];
const isApi = (p) => API_PREFIXES.some((x) => p === x || p.startsWith(x + '/'));

// ---- API key auth (every API route except /health) ----
app.use((req, res, next) => {
  if (!isApi(req.path) || req.path === '/health') return next();
  const key = req.header('x-api-key');
  if (!config.apiKey || key !== config.apiKey) {
    return res.status(401).json({ error: 'Invalid or missing X-API-Key.' });
  }
  next();
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

app.listen(config.port, () => {
  if (servesPortal) console.log(`\n  Portal UI served from ${publicDir}`);
  console.log(`\n  WorkPilot Migration Agent listening on http://localhost:${config.port}`);
  console.log(`  Host: ${config.hostname}`);
  console.log(`  PowerShell: ${config.declaredModules.join(', ') || '(detected at runtime)'}`);
  console.log(`  Graph app-auth: ${graphConfigured() ? 'configured' : 'NOT configured (set TENANT_ID/CLIENT_ID/CLIENT_SECRET)'}`);
  console.log(`  Allowed portal origins: ${config.allowedOrigins.join(', ') || '(any)'}\n`);
});
