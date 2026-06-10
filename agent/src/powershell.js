import { spawn } from 'node:child_process';
import { config } from './config.js';

let cachedAvailable = null;

export async function psAvailable() {
  if (cachedAvailable !== null) return cachedAvailable;
  try {
    await runPowerShell('$PSVersionTable.PSVersion.ToString()', { timeoutMs: 8000 });
    cachedAvailable = true;
  } catch {
    cachedAvailable = false;
  }
  return cachedAvailable;
}

/**
 * Execute a PowerShell script block and return stdout.
 * onData(line) streams output live for job logs.
 */
export function runPowerShell(script, { timeoutMs = 1000 * 60 * 30, onData } = {}) {
  return new Promise((resolve, reject) => {
    const args = config.powershellExe.includes('pwsh')
      ? ['-NoLogo', '-NonInteractive', '-Command', script]
      : ['-NoLogo', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-Command', script];

    const child = spawn(config.powershellExe, args, { windowsHide: true });
    let stdout = '';
    let stderr = '';

    const timer = setTimeout(() => {
      child.kill();
      reject(new Error(`PowerShell timed out after ${timeoutMs} ms`));
    }, timeoutMs);

    child.stdout.on('data', (b) => {
      const s = b.toString();
      stdout += s;
      if (onData) s.split(/\r?\n/).filter(Boolean).forEach((l) => onData(l));
    });
    child.stderr.on('data', (b) => {
      const s = b.toString();
      stderr += s;
      if (onData) s.split(/\r?\n/).filter(Boolean).forEach((l) => onData(`[stderr] ${l}`));
    });
    child.on('error', (err) => { clearTimeout(timer); reject(err); });
    child.on('close', (code) => {
      clearTimeout(timer);
      if (code === 0) resolve(stdout);
      else reject(new Error(stderr.trim() || `PowerShell exited with code ${code}`));
    });
  });
}
