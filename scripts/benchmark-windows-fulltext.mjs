// Synthetic COM-style change keys; excludes real Outlook/OST parsing latency.
import assert from 'node:assert/strict';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir, platform } from 'node:os';
import path from 'node:path';
import { performance } from 'node:perf_hooks';
import { WindowsFulltextIndex } from '../src/windows-fulltext-index.mjs';

const directory = await mkdtemp(path.join(tmpdir(), 'outlook-fulltext-bench-'));
const documents = Array.from({ length: 300 }, (_, i) => ({
  id: `synthetic-${i}`, signature: `change-${i}`, text: '邮件测试'.repeat(2500) + (i % 7 === 0 ? '预算审批' : '事项')
}));
let reads = 0;
const run = async (index, query) => {
  const begin = performance.now();
  const before = reads;
  const hits = [];
  for (let offset = 0; offset < documents.length; offset += 50) {
    const candidates = documents.slice(offset, offset + 50);
    const result = await index.evaluate(candidates, {
      query, loadBody: async document => { reads++; return { status: 'ready', text: document.text, signature: document.signature }; }
    });
    assert.equal(result.body_search_incomplete, false);
    hits.push(...result.hits.keys());
  }
  assert.deepEqual(hits.sort(), documents.filter(document => document.text.includes(query)).map(document => document.id).sort());
  return { milliseconds: Number((performance.now() - begin).toFixed(2)), source_body_reads: reads - before, hits: hits.length };
};
try {
  const options = { directory: path.join(directory, 'index') };
  const index = new WindowsFulltextIndex(options);
  const cold = await run(index, '预算');
  assert.equal(cold.source_body_reads, 300);
  const warm = await run(index, '预算');
  assert.equal(warm.source_body_reads, 0);
  const restarted = new WindowsFulltextIndex(options);
  const restart = await run(restarted, '预算审批');
  assert.equal(restart.source_body_reads, 0);
  documents[0].text = '已修改为决算'; documents[0].signature = 'new-change';
  const changed = await run(restarted, '预算');
  assert.equal(changed.source_body_reads, 1);
  console.log(JSON.stringify({ platform: platform(), node: process.version, documents: documents.length,
    synthetic_only: true, cold, warm, restart, changed }, null, 2));
} finally {
  await rm(directory, { recursive: true, force: true });
}
