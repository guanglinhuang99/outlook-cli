import test from 'node:test';
import assert from 'node:assert/strict';
import { execFile, spawnSync } from 'node:child_process';
import { readFile } from 'node:fs/promises';
import { EventEmitter } from 'node:events';
import { PassThrough, Writable } from 'node:stream';
import { fileURLToPath } from 'node:url';
import { OutlookBridge } from '../src/bridge.mjs';

const runFile = (command, args) => new Promise((resolve, reject) => {
  execFile(command, args, { encoding: 'utf8', maxBuffer: 32 * 1024 * 1024 }, (error, stdout, stderr) => {
    if (error) { error.stdout = stdout; error.stderr = stderr; reject(error); return; }
    resolve({ stdout, stderr });
  });
});

function findPowerShell() {
  const configured = process.env.OUTLOOK_TEST_POWERSHELL;
  const commands = configured ? [configured] : (process.platform === 'win32' ? ['powershell.exe', 'pwsh.exe'] : ['pwsh']);
  for (const command of commands) {
    const probe = spawnSync(command, ['-NoLogo', '-NoProfile', '-NonInteractive', '-Command', 'exit 0'], { stdio: 'ignore' });
    if (probe.status === 0) return command;
  }
  return null;
}

function fakeWorker(handler) {
  const worker = new EventEmitter();
  worker.stdout = new PassThrough();
  worker.stderr = new PassThrough();
  worker.stdin = new Writable({
    write(chunk, encoding, callback) {
      handler(JSON.parse(chunk.toString()), worker);
      callback();
    }
  });
  worker.kill = () => {
    worker.stdout.end();
    worker.stderr.end();
    worker.emit('exit', 0);
  };
  return worker;
}

test('COM bridge applies an internal per-call timeout and counts queue wait', async t => {
  const seen = [];
  const bridge = new OutlookBridge({ timeoutMs: 500, spawnWorker: () => fakeWorker((request, worker) => {
    seen.push(request.method);
    if (request.method === 'slow') {
      setTimeout(() => worker.stdout.write(JSON.stringify({ id: request.id, result: { ok: true } }) + '\n'), 40);
    }
  }) });
  t.after(() => bridge.close());

  const first = bridge.call('slow', {}, { timeoutMs: 200 });
  const second = bridge.call('fulltext_read', { id: 'x' }, { timeoutMs: 10 });
  assert.equal((await first).ok, true);
  await assert.rejects(second, { code: 'OUTLOOK_TIMEOUT' });
  assert.deepEqual(seen, ['slow']);
});

test('COM bridge rejects closed calls before creating the optional full-text index', async () => {
  const bridge = new OutlookBridge();
  bridge.close();
  await assert.rejects(bridge.call('search', { search_body: true }), { code: 'STOPPING' });
  assert.equal(bridge.fulltextIndex, undefined);
});

test('PowerShell COM full-text worker keeps the internal contract scoped and body-only', async () => {
  const source = await readFile(new URL('../powershell/worker.ps1', import.meta.url), 'utf8');
  assert.match(source, /'fulltext_candidates'/);
  assert.match(source, /'fulltext_validate'/);
  assert.match(source, /'fulltext_read'/);
  assert.match(source, /function Resolve-FullText-Scope/);
  assert.match(source, /function FullText-Validate/);
  assert.match(source, /function Folder-Is-Within/);
  assert.match(source, /0x65E20102/);
  assert.match(source, /DownloadState/);
  assert.match(source, /LastModificationTime/);
  assert.match(source, /16000000/);
  assert.match(source, /FULLTEXT_FOLDER_SCOPE: Full-text operations exclude Deleted Items/);
  assert.match(source, /FullText-Result 'truncated' \$null/);
  assert.match(source, /\$null -eq \$before\.ephemeral_guard/);
  assert.match(source, /\$Params\.guard/);
  assert.match(source, /FULLTEXT_SOURCE_CHANGED/);
  assert.match(source, /_fulltext_signature/);
  assert.match(source, /_fulltext_guard/);
  assert.match(source, /_fulltext_offset/);
  assert.match(source, /folder_id = \$Scope\.id/);
  assert.match(source, /validation = @\{ candidates =/);
  assert.match(source, /fulltext_validate accepts at most 50 candidates/);

  const start = source.indexOf('function FullText-Candidates');
  const end = source.indexOf('function FullText-Read');
  assert.ok(start >= 0 && end > start);
  const candidates = source.slice(start, end);
  assert.doesNotMatch(candidates, /\$mail\.Body/);

  const readStart = end;
  const readEnd = source.indexOf('function Iso-Date');
  const reader = source.slice(readStart, readEnd);
  assert.doesNotMatch(reader, /Attachments/);
  assert.doesNotMatch(reader, /\.(?:Save|Move|MarkAsRead)\b/);

  const validateStart = source.indexOf('function FullText-Validate');
  assert.ok(validateStart >= 0 && validateStart < readStart);
  const validator = source.slice(validateStart, readStart);
  assert.doesNotMatch(validator, /\.Body\b/);
});

const powershell = findPowerShell();
test('PowerShell fake-object harness exercises the internal full-text adapter without live COM', { skip: powershell ? false : 'PowerShell unavailable on this host; live Outlook COM is not required for this test' }, async () => {
  const fixture = new URL('./fixtures/windows-fulltext-worker.ps1', import.meta.url);
  const result = await runFile(powershell, ['-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', fileURLToPath(fixture)]);
  const output = JSON.parse(result.stdout.trim());
  assert.equal(output.pass, true, output.message ?? 'fake-object harness failed');
});
