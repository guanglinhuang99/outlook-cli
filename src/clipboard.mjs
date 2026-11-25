import { spawn } from 'node:child_process';
import { randomUUID } from 'node:crypto';
import { mkdir, open, readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { setTimeout as delay } from 'node:timers/promises';
import { root, runtimeDir, prepareRuntime } from './runtime.mjs';

const directory = path.join(runtimeDir, 'clipboard');
const stateFile = path.join(directory, 'clipboard.json');

async function state() {
  try { return JSON.parse(await readFile(stateFile, 'utf8')); }
  catch (error) { if (error.code === 'ENOENT') return null; throw error; }
}

export async function clipboardStatus() {
  const current = await state();
  if (!current) return { running: false, export_dir: path.join(directory, 'exports') };
  let alive = false;
  if (Number.isInteger(current.pid) && current.pid > 0) {
    try { process.kill(current.pid, 0); alive = true; } catch { }
  }
  return { ...current, running: current.running === true && alive, responsive: Date.now() - Date.parse(current.updated_at) < 15_000 };
}

export async function clipboardCommand(action) {
  if (process.platform !== 'win32') throw new Error('The clipboard bridge requires Windows and classic Outlook.');
  if (!['start', 'stop', 'status'].includes(action)) throw new Error('Usage: outlook-local clipboard start|stop|status');
  if (action === 'status') return clipboardStatus();
  const current = await clipboardStatus();
  if (action === 'stop') {
    if (!current.instance || !/^[a-f0-9-]{36}$/.test(current.instance) || !current.running) return { ...current, already_stopped: true };
    await writeFile(path.join(directory, `${current.instance}.stop`), '', { flag: 'wx' }).catch(error => { if (error.code !== 'EEXIST') throw error; });
    for (let i = 0; i < 100; i++) {
      await delay(150);
      const next = await state();
      if (next?.instance === current.instance && next.running === false) return { ...next, running: false };
    }
    throw new Error('Bridge is still finishing a clipboard operation. Its stop request remains pending.');
  }
  if (current.running) return { ...current, already_running: true };
  await prepareRuntime();
  await mkdir(directory, { recursive: true });
  const instance = randomUUID();
  const logPath = path.join(directory, 'clipboard.log');
  const log = await open(path.join(directory, 'launcher.log'), 'a');
  const powershell = path.join(process.env.SystemRoot ?? 'C:\\Windows', 'System32/WindowsPowerShell/v1.0/powershell.exe');
  const literal = value => `'${value.replaceAll("'", "''")}'`;
  const argumentsLine = ['-NoLogo', '-NoProfile', '-NonInteractive', '-STA', '-ExecutionPolicy', 'Bypass', '-File', path.join(root, 'powershell', 'clipboard.ps1'), '-StateDir', directory, '-Instance', instance]
    .map(value => `"${value}"`).join(' ');
  // Windows PowerShell exits immediately under libuv's DETACHED_PROCESS.
  // Start-Process creates a separate hidden console and survives this launcher.
  const launch = `$ErrorActionPreference = 'Stop'; Start-Process -FilePath ${literal(powershell)} -ArgumentList ${literal(argumentsLine)} -WindowStyle Hidden -RedirectStandardError ${literal(logPath)} -RedirectStandardOutput ${literal(path.join(directory, 'clipboard-output.log'))}`;
  const child = spawn(powershell, ['-NoLogo', '-NoProfile', '-NonInteractive', '-EncodedCommand', Buffer.from(launch, 'utf16le').toString('base64')], {
    cwd: root, windowsHide: true, stdio: ['ignore', log.fd, log.fd],
  });
  let failure;
  child.on('error', error => { failure = error; });
  child.unref();
  await log.close();
  for (let i = 0; i < 100; i++) {
    if (failure) throw failure;
    await delay(150);
    const next = await clipboardStatus();
    if (next.instance === instance && next.running) return next;
    if (child.exitCode !== null && child.exitCode !== 0) break;
  }
  throw new Error(`Clipboard bridge did not start (exit=${child.exitCode}, signal=${child.signalCode}). See ${logPath}`);
}
