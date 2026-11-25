import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdtemp, rm } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { promisify } from 'node:util';
import { fileURLToPath } from 'node:url';
import test from 'node:test';

const run = promisify(execFile);
test('Windows clipboard descriptor, file-drop and MSG storage preservation', { skip: process.platform !== 'win32' }, async () => {
  const directory = await mkdtemp(path.join(os.tmpdir(), 'outlook-clipboard-test-'));
  try {
    const { stdout } = await run('powershell.exe', ['-NoProfile', '-NonInteractive', '-STA', '-ExecutionPolicy', 'Bypass', '-File', fileURLToPath(new URL('./fixtures/clipboard-selftest.ps1', import.meta.url)), '-OutputDir', directory], { timeout: 30_000, windowsHide: true });
    assert.match(stdout, /clipboard self-test passed/);
  } finally { await rm(directory, { recursive: true, force: true }); }
});
