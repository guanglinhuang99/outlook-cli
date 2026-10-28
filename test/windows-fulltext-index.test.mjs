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

test('direct scan matches SQLite literal hits and snippets without creating a cache', async t => {
  const { index, directory } = await fixture(t);
  t.after(() => index.close());
  const texts = ['会议预算复核 Straße a*b "OR" 中文关键字 İ 😀', '', `${'x'.repeat(400)}尾部独有预算复核`];
  const candidates = texts.map((text, i) => ({ id: String(i), signature: 'v1' }));
  const bodies = new Map(candidates.map((c, i) => [c.id, { ...c, status: 'unverified', text: texts[i] }]));
  assert.deepEqual(index.directReadBudget(), { max_bytes: 134217728, max_body_bytes: 134217728 });
  const scans = new Map();
  for (const query of ['算', '预算', '预算复核', 'a*b', '"OR"', 'Straße', "' OR 1=1", '独有', 'i\u0307', '😀']) {
    const result = index.evaluateDirect(candidates, bodies, { query });
    assert.equal(result.fulltext.complete, true);
    assert.equal(result.fulltext.evaluated, 3);
    assert.equal(result.fulltext.indexed, 0);
    scans.set(query, result.hits);
  }
  await assert.rejects(stat(directory), { code: 'ENOENT' });
  for (const [query, hits] of scans) {
    const result = await index.evaluate(candidates, { query, loadBody: async c => ({ ...bodies.get(c.id), status: 'ready' }) });
    assert.deepEqual(result.hits, hits, query);
  }
  const rejected = await index.evaluate([{ id: 'unverified', signature: 'v1' }], {
    query: 'hit', loadBody: async c => ({ ...c, status: 'unverified', text: 'hit' })
  });
  assert.equal(rejected.fulltext.errors, 1);
  assert.equal(rejected.hits.size, 0);
});

test('direct scan retains incomplete, retry, byte and time limits', async t => {
  let clock = 1_800_000_000_000;
  const { index } = await fixture(t, { singleBatchBytes: 8, bodyCapCodeUnits: 6, now: () => clock });
  const run = statuses => {
    const candidates = statuses.map((b, i) => ({ id: String(i), signature: 'v1' }));
    return index.evaluateDirect(candidates, new Map(candidates.map((c, i) => [c.id, { ...c, ...statuses[i] }])), { query: 'test' });
  };
  const mixed = run([{ status: 'unverified', text: 'test' }, { status: 'missing' }, { status: 'truncated' },
    { status: 'error' }, { status: 'pending' }, { status: 'ready', text: 'test', signature: 'v2' }]);
  assert.equal(mixed.hits.size, 1);
  assert.equal(mixed.fulltext.missing, 2); assert.equal(mixed.fulltext.errors, 2); assert.equal(mixed.fulltext.pending, 1);
  assert.equal(mixed.body_search_incomplete, true); assert.equal(mixed.retry_required, true);
  assert.equal(run([{ status: 'ready', text: '1234567' }]).fulltext.missing, 1);
  const capped = run(Array.from({ length: 3 }, () => ({ status: 'ready', text: 'test' })));
  assert.equal(capped.fulltext.evaluated, 2); assert.equal(capped.fulltext.pending, 1);
  const c = { id: 'a', signature: 'v1' };
  const expired = index.evaluateDirect([c], new Map([['a', { ...c, status: 'ready', text: 'test' }]]), { query: 'test', deadline: clock });
  assert.equal(expired.fulltext.pending, 1); assert.equal(expired.fulltext.evaluated, 0);
});

test('compact direct no-matches preserve query/version identity, byte caps and cannot enter SQLite', async t => {
  const { index } = await fixture(t, { singleBatchBytes: 8 });
  const candidates = ['a', 'b', 'c'].map(id => ({ id, signature: 'v1' }));
  const body = { status: 'unverified_no_match', signature: 'v1', query: 'test', body_bytes: 4, body_code_units: 4 };
  const result = index.evaluateDirect(candidates, new Map(candidates.map(c => [c.id, body])), { query: 'TEST' });
  assert.equal(result.fulltext.evaluated, 2); assert.equal(result.fulltext.omitted_bodies, 2);
  assert.equal(result.fulltext.pending, 1); assert.equal(result.retry_required, true);
  assert.equal(result.hits.size, 0);
  for (const extra of [{ query: 'other' }, { signature: 'v2' }, { body_bytes: -1 }, { body_code_units: 5 }]) {
    const invalid = index.evaluateDirect(candidates.slice(0, 1), new Map([['a', { ...body, ...extra }]]), { query: 'test' });
    assert.equal(invalid.fulltext.errors, 1); assert.equal(invalid.fulltext.evaluated, 0);
  }
  const rejected = await index.evaluate(candidates.slice(0, 1), { query: 'test', loadBody: async () => body });
  assert.equal(rejected.fulltext.errors, 1); assert.equal(rejected.fulltext.updated, 0);
});

test('inline cold reads require a prepared empty private cache and preserve configured bounds', async t => {
  const { index, directory } = await fixture(t, { singleBatchBytes: 8 });
  assert.deepEqual(await index.prepareColdRead(), { max_bytes: 8, max_body_bytes: 8 });
  assert.equal(await index.prepareColdRead(), null);
  assert.equal(await new WindowsFulltextIndex({ directory }).prepareColdRead(), null);
  assert.ok((await stat(index.databasePath)).size > 0);
  const { index: raced } = await fixture(t);
  const open = raced._openDatabase.bind(raced);
  raced._openDatabase = async () => {
    const db = await open();
    raced._storeReady(db, { id: 'concurrent', signature: 'v1' }, 'text');
    return db;
  };
  assert.equal(await raced.prepareColdRead(), null);
});

test('unsafe cold cache preparation fails before body prefetch can be requested', async t => {
  const { index, directory } = await fixture(t);
  await writeFile(directory, 'not a directory');
  await assert.rejects(index.prepareColdRead());
});

test('bulk FTS failure cannot commit documents without their matching searchable entries', async t => {
  const { index } = await fixture(t);
  const candidates = ['a', 'b'].map(id => ({ id, signature: 'v1' }));
  const loadBody = async () => { throw Error('batch expected'); };
  const options = { query: 'original', loadBody,
    loadBodies: async missing => missing.map(c => ({ ...c, status: 'ready', text: 'original searchable text' })) };
  const first = await index.evaluate(candidates, options);
  assert.equal(first.hits.size, 2);
  const insert = index._storeFtsBatch.bind(index);
  index._storeFtsBatch = () => { throw Error('injected FTS batch failure'); };
  await assert.rejects(index.evaluate(candidates.map(c => ({ ...c, signature: 'v2' })), options));
  index._storeFtsBatch = insert;
  // _prepareCache removes obsolete versions before loading; after failed
  // replacement no new document may exist without its matching FTS entry.
  const db = await index._openDatabase();
  try {
    assert.equal(db.prepare("SELECT count(*) AS n FROM documents WHERE signature='v2'").get().n, 0);
    index._checkIntegrity(db);
  } finally { db.close(); }
  const rebuilt = await index.evaluate(candidates, options);
  assert.equal(rebuilt.hits.size, 2);
  const reused = await index.evaluate(candidates, { ...options, loadBodies: async () => { throw Error('warm bodies should be reused'); } });
  assert.equal(reused.fulltext.reused, 2);
  assert.equal(reused.hits.size, 2);
});

test('bounded batch loading reads only misses, persists versions and rejects malformed batches', async t => {
  const { index } = await fixture(t, { singleBatchBytes: 8 });
  const candidates = ['a', 'b', 'c'].map(id => ({ id, signature: 'v1' }));
  const calls = [];
  const loadBody = () => { throw Error('single reader must not run'); };
  const loadBodies = async (missing, timeout, maxBytes, bodyBytes) => {
    calls.push(missing.map(item => item.id));
    assert.ok(timeout > 0); assert.equal(maxBytes, 8); assert.equal(bodyBytes, 8);
    return missing.map((item, i) => ({ ...item, status: i === 2 ? 'pending' : 'ready', text: 'test' }));
  };
  const first = await index.evaluate(candidates, { query: 'test', loadBody, loadBodies });
  assert.deepEqual([...first.hits.keys()].sort(), ['a', 'b']);
  assert.equal(first.fulltext.pending, 1); assert.equal(first.retry_required, true);
  const second = await index.evaluate(candidates, { query: 'test', loadBody, loadBodies });
  assert.equal(second.fulltext.reused, 2); assert.equal(second.fulltext.updated, 1);
  assert.deepEqual(calls, [['a', 'b', 'c'], ['c']]);
  await index.evaluate(candidates, { query: 'test', loadBody, loadBodies });
  assert.equal(calls.length, 2);
  const changed = candidates.map(item => ({ ...item, signature: 'v2' }));
  const malformed = await index.evaluate(changed, { query: 'test', loadBody,
    loadBodies: async () => changed.map(() => ({ id: 'a', signature: 'v2', status: 'ready', text: 'x' })) });
  assert.equal(malformed.fulltext.errors, 3);
  assert.equal(malformed.hits.size, 0);
});

test('batch mode preserves source signatures, ephemeral bodies, byte caps and late-result deadlines', async t => {
  let clock = 1_800_000_000_000;
  const { index } = await fixture(t, { now: () => clock, singleBatchBytes: 8 });
  const loadBody = () => { throw Error('single reader must not run'); };
  const candidates = [{ id: 'a', signature: null }, { id: 'b', signature: 'v1' }, { id: 'c', signature: 'v1' }];
  const result = await index.evaluate(candidates, { query: 'test', loadBody, loadBodies: async () => [
    { ...candidates[0], status: 'ready', text: 'test' },
    { ...candidates[1], status: 'ready', text: 'test', signature: 'v2' },
    { ...candidates[2], status: 'ready', text: 'testtest' }
  ] });
  assert.deepEqual([...result.hits.keys()], ['a']);
  assert.equal(result.fulltext.errors, 1); assert.equal(result.fulltext.pending, 1);
  const deadline = clock + 100;
  const late = await index.evaluate(candidates, { query: 'test', loadBody, deadline,
    loadBodies: async () => { clock = deadline; return candidates.map(item => ({ ...item, status: 'ready', text: 'test' })); } });
  assert.equal(late.fulltext.pending, 3); assert.equal(late.fulltext.updated, 0);
  assert.equal(late.fulltext.reused, 0); assert.equal(late.hits.size, 0);
});

test('a failed batch write rolls back every document and releases the database lock', async t => {
  const { index } = await fixture(t);
  const candidates = ['a', 'b'].map(id => ({ id, signature: 'v1' }));
  const options = { query: 'test', loadBody: () => { throw Error('single reader must not run'); },
    loadBodies: async items => items.map(item => ({ ...item, status: 'ready', text: 'test' })) };
  const store = index._storeReady.bind(index);
  index._storeReady = (db, candidate, ...args) => {
    store(db, candidate, ...args);
    if (candidate.id === 'b') throw Error('injected write failure');
  };
  await assert.rejects(index.evaluate(candidates, options), { code: 'FULLTEXT_ERROR' });
  index._storeReady = store;
  const result = await index.evaluate(candidates, options);
  assert.equal(result.fulltext.reused, 0); assert.equal(result.fulltext.updated, 2);
});

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
