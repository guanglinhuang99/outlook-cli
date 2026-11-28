import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import test from 'node:test';
import { WINDOWS_MINIMUM_NODE_MAJOR, nodeVersionError } from '../src/node-version.mjs';

test('Windows rejects Node versions below the supported major with an actionable error', () => {
  assert.equal(nodeVersionError('win32', `${WINDOWS_MINIMUM_NODE_MAJOR}.0.0`), null);
  assert.equal(nodeVersionError('darwin', '22.16.0'), null);
  const error = nodeVersionError('win32', '22.16.0');
  assert.equal(error.code, 'NODE_VERSION_UNSUPPORTED');
  assert.match(error.message, new RegExp(`requires Node\\.js ${WINDOWS_MINIMUM_NODE_MAJOR}\\+ \\(this is 22\\.16\\.0\\)`));
});

test('Windows install docs and CI follow the Windows Node requirement', async () => {
  const files = ['README.md', 'INSTALL-WINDOWS.md', 'docs/agent-install-guide.md', 'docs/usage-reference.md'];
  for (const file of files) {
    const text = await readFile(new URL(`../${file}`, import.meta.url), 'utf8');
    for (const line of text.split('\n').filter(line => /Windows|PowerShell/.test(line))) {
      for (const [, major] of line.matchAll(/Node(?:\.js)?\s*\**\s*(\d+)(?:\.\d+)*\+/g)) {
        if (/macOS|Mac |Python/.test(line) && Number(major) < WINDOWS_MINIMUM_NODE_MAJOR) continue; // macOS requirement line
        assert.ok(Number(major) >= WINDOWS_MINIMUM_NODE_MAJOR,
          `${file}: "${line.trim().slice(0, 80)}" names Node ${major}+; Windows needs ${WINDOWS_MINIMUM_NODE_MAJOR}+ (src/node-version.mjs).`);
      }
    }
  }
  const ci = await readFile(new URL('../.github/workflows/ci.yml', import.meta.url), 'utf8');
  assert.match(ci, /exclude:\s*\n\s*- os: windows-latest\s*\n\s*node: '22/, 'CI must not run Windows on unsupported Node 22');
});
