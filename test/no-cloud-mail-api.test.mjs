import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';

const root = fileURLToPath(new URL('..', import.meta.url));
const self = path.relative(root, fileURLToPath(import.meta.url)).split(path.sep).join('/');

// The product reads only mail Outlook already synchronized to this computer; the owner
// cannot use Microsoft Graph, so no Graph command, client, credential or doc may return.
test('no Microsoft Graph code, commands or documentation', async t => {
  let files;
  try {
    files = execFileSync('git', ['ls-files', '-co', '--exclude-standard'], { cwd: root, encoding: 'utf8' }).split('\n');
  } catch {
    return t.skip('git is unavailable to list project files');
  }
  const text = files.filter(file => file && file !== self && /\.(?:mjs|js|py|ps1|cs|cmd|sh|md|txt|json|yml|yaml)$/.test(file));
  const offenders = [];
  for (const file of text) {
    const content = (await readFile(path.join(root, file), 'utf8').catch(() => '')).replace(/object graphs?/gi, '');
    if (/\bgraph\b|graph\.microsoft|OUTLOOK_GRAPH/i.test(content)) offenders.push(file);
  }
  assert.deepEqual(offenders, [], 'Remove Microsoft Graph references; this tool only reads locally synchronized Outlook mail.');
});
