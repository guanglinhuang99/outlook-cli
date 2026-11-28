import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import test from 'node:test';
import { MINIMUM_NODE } from '../src/windows-fulltext-index.mjs';

const files = ['README.md', 'INSTALL-WINDOWS.md', 'docs/agent-install-guide.md', 'docs/usage-reference.md', 'docs/macos.md'];

test('documented and tested Node minimums match the full-text runtime requirement', async () => {
  for (const file of files) {
    const text = await readFile(new URL(`../${file}`, import.meta.url), 'utf8');
    for (const [version] of text.matchAll(/(?<![\d.])22\.\d+(?![\d.])/g)) {
      assert.equal(version, MINIMUM_NODE, `${file} names Node ${version}; Windows full-text search needs ${MINIMUM_NODE} (src/windows-fulltext-index.mjs MINIMUM_NODE). Update the document.`);
    }
  }
  const ci = await readFile(new URL('../.github/workflows/ci.yml', import.meta.url), 'utf8');
  assert.match(ci, new RegExp(`node: \\[[^\\]]*'${MINIMUM_NODE.replace('.', '\\.')}'`), 'CI must test the documented minimum Node version');
});
