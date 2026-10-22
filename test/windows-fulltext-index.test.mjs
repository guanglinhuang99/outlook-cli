import test from 'node:test';
import assert from 'node:assert/strict';
import { DatabaseSync } from 'node:sqlite';
import { link, mkdtemp, readFile, rm, stat, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';

import { WindowsFulltextIndex } from '../src/windows-fulltext-index.mjs';

async function fixture(t, options = {}) {
  const root = await mkdtemp(path.join(os.tmpdir(), 'outlook-fulltext-'));
  const directory = path.join(root, 'fulltext');
  t.after(() => rm(root, { recursive: true, force: true }));
  return { root, directory, index: new WindowsFulltextIndex({ directory, ...options }) };
}

function loader(texts, calls = []) {
  return async candidate => {
    calls.push(candidate.id);
    const value = texts[candidate.id];
    if (typeof value === 'object') return value;
    return { status: 'ready', text: value, signature: candidate.signature };
  };
}

test('searches literal Chinese, short queries, Unicode and FTS metacharacters', async t => {
  const { index } = await fixture(t);
  const candidates = [
    { id: 'one', signature: 'v1' },
    { id: 'two', signature: 'v1' },
    { id: 'three', signature: 'v1' }
  ];
  const texts = {
    one: '会议预算复核 Straße a*b "OR" 中文关键字',
    two: '没有目标词',
    three: `${'x'.repeat(400)}尾部独有预算复核`
  };
  for (const [query, expected] of [
    ['算', ['one', 'three']],
    ['预算', ['one', 'three']],
    ['预算复核', ['one', 'three']],
    ['a*b', ['one']],
    ['"OR"', ['one']],
    ['Straße', ['one']],
    ["' OR 1=1", []],
    ['独有', ['three']]
  ]) {
    const result = await index.evaluate(candidates, { query, loadBody: loader(texts) });
    assert.deepEqual([...result.hits.keys()].sort(), expected.sort(), query);
    assert.equal(result.fulltext.errors, 0);
    for (const snippet of result.hits.values()) assert.ok(snippet.length <= 240);
  }
  const nul = await index.evaluate(candidates, { query: '尾\0部', loadBody: loader(texts) });
  assert.equal(nul.hits.size, 0);
});

test('warm reuse, restart, version replacement and cache retention across scopes', async t => {
  const { directory } = await fixture(t);
  const first = new WindowsFulltextIndex({ directory });
  const candidates = [{ id: 'one', signature: 'v1' }, { id: 'two', signature: 'v1' }];
  const firstCalls = [];
  const firstResult = await first.evaluate(candidates, {
    query: '预算',
    loadBody: loader({ one: '预算旧正文', two: '别的正文' }, firstCalls)
  });
  assert.deepEqual([...firstResult.hits.keys()], ['one']);
  assert.equal(firstResult.fulltext.updated, 2);
  assert.deepEqual(firstCalls, ['one', 'two']);

  const restarted = new WindowsFulltextIndex({ directory });
  const warmCalls = [];
  const warm = await restarted.evaluate(candidates, {
    query: '预算',
    loadBody: loader({ one: 'should not read', two: 'should not read' }, warmCalls)
  });
  assert.deepEqual([...warm.hits.keys()], ['one']);
  assert.equal(warm.fulltext.reused, 2);
  assert.equal(warm.fulltext.updated, 0);
  assert.deepEqual(warmCalls, []);

  const changed = await restarted.evaluate([{ id: 'one', signature: 'v2' }], {
    query: '结算',
    loadBody: loader({ one: '已改正文结算确认' })
  });
  assert.deepEqual([...changed.hits.keys()], ['one']);
  assert.deepEqual([...await restarted.evaluate([{ id: 'one', signature: 'v2' }], {
    query: '预算', loadBody: loader({ one: 'should not read' })
  }).then(result => result.hits.keys())], []);

  const scoped = await restarted.evaluate([{ id: 'one', signature: 'v2' }], {
    query: '别的', loadBody: loader({ one: 'should not read' })
  });
  assert.equal(scoped.hits.size, 0);
  assert.equal(scoped.fulltext.reused, 1);

  // A candidate window is a search scope, not the cache owner. Rows for a
  // different page/filter remain reusable when the current scope changes.
  const otherPage = await restarted.evaluate([{ id: 'two', signature: 'v1' }], {
    query: '别的正文',
    loadBody: async () => { throw new Error('retained row should not be loaded'); }
  });
  assert.deepEqual([...otherPage.hits.keys()], ['two']);
  assert.equal(otherPage.fulltext.reused, 1);
  const emptyQuery = await restarted.evaluate([{ id: 'two', signature: 'v1' }], {
    query: '',
    loadBody: async () => { throw new Error('empty query should warm-reuse'); }
  });
  assert.equal(emptyQuery.hits.size, 0);
  assert.equal(emptyQuery.fulltext.reused, 1);
  assert.equal(emptyQuery.fulltext.complete, true);
});

test('concurrent disjoint evaluations retain both cache rows', async t => {
  const { directory } = await fixture(t);
  const left = new WindowsFulltextIndex({ directory });
  const right = new WindowsFulltextIndex({ directory });
  const pause = () => new Promise(resolve => setTimeout(resolve, 5));
  const [leftResult, rightResult] = await Promise.all([
    left.evaluate([{ id: 'left', signature: 'v1' }], {
      query: '左侧正文',
      loadBody: async () => { await pause(); return { status: 'ready', text: '左侧正文', signature: 'v1' }; }
    }),
    right.evaluate([{ id: 'right', signature: 'v1' }], {
      query: '右侧正文',
      loadBody: async () => { await pause(); return { status: 'ready', text: '右侧正文', signature: 'v1' }; }
    })
  ]);
  assert.deepEqual([...leftResult.hits.keys()], ['left']);
  assert.deepEqual([...rightResult.hits.keys()], ['right']);

  const restart = new WindowsFulltextIndex({ directory });
  const [leftWarm, rightWarm] = await Promise.all([
    restart.evaluate([{ id: 'left', signature: 'v1' }], {
      query: '左侧正文', loadBody: async () => { throw new Error('left should be warm'); }
    }),
    restart.evaluate([{ id: 'right', signature: 'v1' }], {
      query: '右侧正文', loadBody: async () => { throw new Error('right should be warm'); }
    })
  ]);
  assert.equal(leftWarm.fulltext.reused, 1);
  assert.equal(rightWarm.fulltext.reused, 1);
  assert.deepEqual([...leftWarm.hits.keys()], ['left']);
  assert.deepEqual([...rightWarm.hits.keys()], ['right']);
});

test('missing bodies retry, invalid signatures and partial deadlines stay incomplete', async t => {
  const { directory } = await fixture(t);
  const index = new WindowsFulltextIndex({ directory });
  const candidate = { id: 'late', signature: 'v1' };
  let calls = 0;
  const missing = await index.evaluate([candidate], {
    query: '目标',
    loadBody: async () => {
      calls += 1;
      return { status: 'missing' };
    }
  });
  assert.equal(missing.fulltext.missing, 1);
  assert.equal(missing.retry_required, false);
  assert.equal(missing.body_search_incomplete, true);
  const recovered = await index.evaluate([candidate], {
    query: '目标',
    loadBody: async current => {
      calls += 1;
      return { status: 'ready', text: '补回来的目标正文', signature: current.signature };
    }
  });
  assert.deepEqual([...recovered.hits.keys()], ['late']);
  assert.equal(calls, 2);

  const bad = await index.evaluate([{ id: 'bad', signature: 'expected' }], {
    query: '正文',
    loadBody: async () => ({ status: 'ready', text: '正文', signature: 'different' })
  });
  assert.equal(bad.fulltext.errors, 1);
  assert.equal(bad.hits.size, 0);

  const callsBeforeDeadline = [];
  const pending = await index.evaluate([
    { id: 'a', signature: 'a' }, { id: 'b', signature: 'b' }
  ], {
    query: '预算', deadline: 0, loadBody: loader({ a: '预算 a', b: '预算 b' }, callsBeforeDeadline)
  });
  assert.equal(pending.fulltext.pending, 2);
  assert.deepEqual(callsBeforeDeadline, []);
});

test('ephemeral candidates are searchable once and never warm-reused', async t => {
  const { index } = await fixture(t);
  const candidate = { id: 'ephemeral', signature: null };
  let calls = 0;
  const loadBody = async current => {
    calls += 1;
    return { status: 'ready', text: '临时预算正文', signature: current.signature };
  };
  const first = await index.evaluate([candidate], { query: '预算', loadBody });
  const second = await index.evaluate([candidate], { query: '预算', loadBody });
  assert.deepEqual([...first.hits.keys()], ['ephemeral']);
  assert.deepEqual([...second.hits.keys()], ['ephemeral']);
  assert.equal(calls, 2);
  assert.equal(first.fulltext.reused, 0);
  assert.equal(second.fulltext.reused, 0);
});

test('a concurrent replacement cannot turn an old candidate into a complete no-match', async t => {
  const { directory, index } = await fixture(t);
  const other = new WindowsFulltextIndex({ directory });
  await index.evaluate([{ id: 'same', signature: 'old' }], { query: '旧预算', loadBody: loader({ same: '旧预算' }) });
  const result = await index.evaluate([{ id: 'same', signature: 'old' }, { id: 'new', signature: 'v1' }], {
    query: '旧预算',
    loadBody: async candidate => {
      await other.evaluate([{ id: 'same', signature: 'new' }], { query: '新预算', loadBody: loader({ same: '新预算' }) });
      return { status: 'ready', text: '其他', signature: candidate.signature };
    }
  });
  assert.equal(result.fulltext.pending, 1);
  assert.equal(result.fulltext.indexed, 1);
  assert.equal(result.retry_required, true);
  assert.equal(result.hits.size, 0);
});

test('a deadline preserves later warm candidates and never persists a late body', async t => {
  let now = 1_800_000_000_000;
  const { directory } = await fixture(t);
  const index = new WindowsFulltextIndex({ directory, now: () => now });
  const warm = { id: 'warm', signature: 'v1' };
  const cold = { id: 'cold', signature: 'v1' };
  await index.evaluate([warm], { query: '预算', loadBody: loader({ warm: '预算' }) });
  const result = await index.evaluate([cold, warm], {
    query: '预算', deadline: now + 10,
    loadBody: async () => { now += 11; return { status: 'ready', text: '预算', signature: 'v1' }; }
  });
  assert.equal(result.fulltext.pending, 1);
  assert.equal(result.fulltext.reused, 1);
  assert.equal(result.fulltext.indexed, 1);
  assert.deepEqual([...result.hits.keys()], ['warm']);
  const calls = [];
  await index.evaluate([cold], { query: '预算', loadBody: loader({ cold: '预算' }, calls) });
  assert.deepEqual(calls, ['cold']);
});

test('corrupt cache fails explicitly and an empty candidate list retains other pages', async t => {
  const { directory, index } = await fixture(t);
  const candidate = { id: 'one', signature: 'v1' };
  await index.evaluate([candidate], { query: '预算', loadBody: loader({ one: '预算' }) });
  assert.equal((await index.evaluate([], { query: '预算', loadBody: loader({}) })).hits.size, 0);
  const retained = await index.evaluate([candidate], { query: '预算', loadBody: async () => { throw Error('unexpected source read'); } });
  assert.equal(retained.fulltext.reused, 1);
  const wrongSchema = new DatabaseSync(path.join(directory, 'fulltext.sqlite3'));
  wrongSchema.exec('PRAGMA user_version=999');
  wrongSchema.close();
  await assert.rejects(new WindowsFulltextIndex({ directory }).evaluate([candidate], {
    query: '预算', loadBody: loader({ one: '预算' })
  }), { code: 'FULLTEXT_CORRUPT' });
  await writeFile(path.join(directory, 'fulltext.sqlite3'), 'not a SQLite database');
  await assert.rejects(new WindowsFulltextIndex({ directory }).evaluate([candidate], {
    query: '预算', loadBody: loader({ one: '预算' })
  }), { code: 'FULLTEXT_CORRUPT' });
});

test('private database paths, source immutability and busy/capacity failures are explicit', async t => {
  const { root, directory } = await fixture(t);
  const index = new WindowsFulltextIndex({ directory });
  const source = path.join(root, 'original-source.dat');
  await writeFile(source, 'source bytes');
  const candidate = { id: 'one', signature: 'v1' };
  await index.evaluate([candidate], {
    query: '正文',
    loadBody: async () => ({ status: 'ready', text: await readFile(source, 'utf8') + ' 正文', signature: 'v1' })
  });
  assert.equal(await readFile(source, 'utf8'), 'source bytes');
  const dbPath = path.join(directory, 'fulltext.sqlite3');
  if (process.platform !== 'win32') {
    assert.equal((await stat(directory)).mode & 0o077, 0);
    assert.equal((await stat(dbPath)).mode & 0o077, 0);
  }

  const hardlink = path.join(root, 'hardlink.sqlite3');
  await link(dbPath, hardlink);
  await assert.rejects(
    () => new WindowsFulltextIndex({ directory }).evaluate([candidate], { query: '正文', loadBody: loader({ one: '正文' }) }),
    error => error.code === 'FULLTEXT_UNSAFE_PATH'
  );
  await rm(hardlink);

  const lock = new DatabaseSync(dbPath, { timeout: 0 });
  lock.exec('BEGIN IMMEDIATE');
  const busy = new WindowsFulltextIndex({ directory, busyTimeoutMs: 0 });
  await assert.rejects(
    () => busy.evaluate([candidate], { query: '正文', loadBody: loader({ one: '正文' }) }),
    error => error.code === 'FULLTEXT_BUSY' || /busy|locked/i.test(error.message)
  );
  lock.exec('ROLLBACK');
  lock.close();

  const capped = new WindowsFulltextIndex({ directory: path.join(root, 'capped'), maxBytes: 8192 });
  await assert.rejects(
    () => capped.evaluate([{ id: 'large', signature: 'v1' }], {
      query: 'x', loadBody: loader({ large: 'x'.repeat(10_000) })
    }),
    error => error.code === 'FULLTEXT_CAP' || /capacity|size|full/i.test(error.message)
  );
});

test('body and batch caps report bounded misses without persisting a partial document', async t => {
  const { directory } = await fixture(t, { bodyCapCodeUnits: 4, singleBatchBytes: 16 });
  const index = new WindowsFulltextIndex({ directory, bodyCapCodeUnits: 4, singleBatchBytes: 16 });
  const result = await index.evaluate([{ id: 'large', signature: 'v1' }], {
    query: 'large', loadBody: loader({ large: 'too large body' })
  });
  assert.equal(result.fulltext.missing, 1);
  assert.equal(result.fulltext.pending, 0);
  assert.equal(result.retry_required, false);
  assert.equal(result.body_search_incomplete, true);
  assert.equal(result.hits.size, 0);
  const retryCalls = [];
  await index.evaluate([{ id: 'large', signature: 'v1' }], {
    query: 'large', loadBody: loader({ large: 'ok' }, retryCalls)
  });
  assert.deepEqual(retryCalls, ['large']);

  const batchIndex = new WindowsFulltextIndex({
    directory: path.join(path.dirname(directory), 'batch'), bodyCapCodeUnits: 100, singleBatchBytes: 6
  });
  const batchCalls = [];
  const batch = await batchIndex.evaluate([
    { id: 'first', signature: 'v1' }, { id: 'second', signature: 'v1' }
  ], {
    query: 'body',
    loadBody: loader({ first: 'first', second: 'second' }, batchCalls)
  });
  assert.equal(batch.fulltext.indexed, 1);
  assert.equal(batch.fulltext.pending, 1);
  assert.equal(batch.retry_required, true);
  assert.deepEqual(batchCalls, ['first', 'second']);
  const second = await batchIndex.evaluate([{ id: 'second', signature: 'v1' }], {
    query: 'second', loadBody: loader({ second: 'second' })
  });
  assert.deepEqual([...second.hits.keys()], ['second']);
});
