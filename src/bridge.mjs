import { spawn } from 'node:child_process';
import { createInterface } from 'node:readline';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import { WindowsFulltextIndex } from './windows-fulltext-index.mjs';
import { searchWindowsFulltext } from './windows-fulltext.mjs';

const workerPath = fileURLToPath(new URL('../powershell/worker.ps1', import.meta.url));
export class OutlookBridge {
  constructor({ timeoutMs = 45000, spawnWorker, fulltextIndex, enableWindowsFulltext = true } = {}) {
    this.timeoutMs = timeoutMs;
    this.fulltextIndex = fulltextIndex;
    this.enableWindowsFulltext = enableWindowsFulltext;
    this.spawnWorker = spawnWorker ?? (() => spawn(path.join(process.env.SystemRoot ?? 'C:\\Windows', 'System32/WindowsPowerShell/v1.0/powershell.exe'),
      ['-NoLogo', '-NoProfile', '-NonInteractive', '-STA', '-ExecutionPolicy', 'Bypass', '-File', workerPath],
      { stdio: ['pipe', 'pipe', 'pipe'], windowsHide: true }));
    this.queue = [];
    this.sequence = 0;
    this.closed = false;
  }
  call(method, params, { timeoutMs = this.timeoutMs } = {}) {
    if (this.closed) return Promise.reject(Object.assign(new Error('Service is stopping'), { code: 'STOPPING' }));
    if (this.enableWindowsFulltext && method === 'search' && params?.search_body) {
      this.fulltextIndex ??= new WindowsFulltextIndex();
      return searchWindowsFulltext(this, params, this.fulltextIndex, {
        inlineColdRead: true, directScan: params.com_body_mode === 'direct',
          deferredIndex: !params.com_body_mode || params.com_body_mode === 'auto'
      });
    }
    if (this.queue.length >= 32) return Promise.reject(Object.assign(new Error('Reader queue is full; retry later'), { code: 'BUSY' }));
    const effectiveTimeout = Number.isFinite(timeoutMs) && timeoutMs > 0 ? timeoutMs : this.timeoutMs;
    return new Promise((resolve, reject) => {
      this.queue.push({ id: ++this.sequence, method, params, resolve, reject, deadline: Date.now() + effectiveTimeout });
      this.pump();
    });
  }
  ensureWorker() {
    if (this.worker) return;
    const worker = this.spawnWorker();
    this.worker = worker;
    const lines = createInterface({ input: worker.stdout });
    lines.on('line', line => {
      if (this.worker !== worker) return;
      try {
        const message = JSON.parse(line);
        if (!this.active || message.id !== this.active.id) throw new Error('Mismatched worker response');
        const active = this.active;
        if (Date.now() >= active.deadline) {
          this.failWorker(worker, 'OUTLOOK_TIMEOUT', 'Reader timed out; retry the request.');
          return;
        }
        clearTimeout(active.timer);
        this.active = null;
        if (message.error) active.reject(Object.assign(new Error(message.error.message), { code: message.error.code }));
        else active.resolve(message.result);
        this.pump();
      } catch {
        this.failWorker(worker, 'WORKER_PROTOCOL', 'Invalid reader worker response');
      }
    });
    // Consume but do not log raw worker diagnostics: they can include mailbox data.
    worker.stderr.resume();
    worker.stdin.on('error', () => this.failWorker(worker, 'WORKER_EXIT', 'Reader worker input closed'));
    worker.on('error', () => this.failWorker(worker, 'WORKER_START', 'Could not start reader worker'));
    worker.on('exit', () => { lines.close(); this.failWorker(worker, 'WORKER_EXIT', 'Reader worker exited; retry the request'); });
  }
  pump() {
    if (this.active || !this.queue.length || this.closed) return;
    const next = this.queue.shift();
    const remaining = next.deadline - Date.now();
    if (remaining <= 0) {
      next.reject(Object.assign(new Error('Reader timed out; retry the request.'), { code: 'OUTLOOK_TIMEOUT' }));
      this.pump();
      return;
    }
    this.active = next;
    try {
      this.ensureWorker();
      const worker = this.worker;
      this.active.timer = setTimeout(() => this.failWorker(worker, 'OUTLOOK_TIMEOUT', 'Reader timed out; retry the request.'), remaining);
      worker.stdin.write(JSON.stringify({ id: this.active.id, method: this.active.method, params: this.active.params }) + '\n');
    } catch {
      const active = this.active;
      this.active = null;
      active.reject(Object.assign(new Error('Could not start reader worker'), { code: 'WORKER_START' }));
      this.pump();
    }
  }
  failWorker(worker, code, message) {
    if (worker !== this.worker) return;
    this.worker = null;
    worker.kill();
    if (this.active) {
      clearTimeout(this.active.timer);
      this.active.reject(Object.assign(new Error(message), { code }));
      this.active = null;
    }
    // Reject queued calls on a worker failure instead of stalling each for another timeout.
    for (const item of this.queue.splice(0)) item.reject(Object.assign(new Error(message), { code }));
  }
  close() {
    this.closed = true;
    this.fulltextIndex?.close?.();
    if (this.worker) this.failWorker(this.worker, 'STOPPING', 'Service stopped');
    for (const item of this.queue.splice(0)) item.reject(new Error('Service stopped'));
  }
}
