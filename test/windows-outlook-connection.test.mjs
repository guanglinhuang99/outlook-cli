import test from 'node:test';
import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { fileURLToPath } from 'node:url';

test('PowerShell Outlook lifecycle: attach, activate, reuse, reconnect and fail explicitly', { skip: process.platform !== 'win32' }, async () => {
  const { stdout } = await promisify(execFile)('powershell.exe', [
    '-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File',
    fileURLToPath(new URL('./fixtures/windows-outlook-connection.ps1', import.meta.url))
  ], { timeout: 30000 });
  assert.equal(JSON.parse(stdout.trim()).pass, true);
});
