import { createHash, randomUUID } from 'node:crypto';
import { constants } from 'node:fs';
import { lstat, open, rename, unlink } from 'node:fs/promises';
import { request } from 'node:http';
import path from 'node:path';
import { prepareRuntime } from './runtime.mjs';

const MAX_SCAN = 500;
const MAX_BODY = 4000;
const BATCH_SIZE = 8;
const MAX_DIMENSIONS = 4096;
const MAX_INDEX_BYTES = 64 * 1024 * 1024;
const MAX_RESPONSE_BYTES = 4 * 1024 * 1024;
const STATUSES = ['indexed', 'truncated', 'missing', 'empty', 'unavailable', 'embedding_unavailable'];
const SCOPE_FIELDS = ['folder', 'include_subfolders', 'from_date', 'to_date', 'scan_limit'];
const digestPattern = /^[a-f0-9]{64}$/;
const snapshotPattern = /^[a-f0-9]{24}$/;
const folderPattern = /^(?:inbox|sent|all|[A-Za-z0-9_-]{8,16000})$/;

function fail(code, message) {
  throw Object.assign(new Error(message), { code });
}

function checkedDate(value, name) {
  const match = /^\d{4}-\d{2}-\d{2}$/.test(value);
  if (typeof value !== 'string' || !match || value < '1970-01-01' || value > '9999-12-30' ||
      !Number.isFinite(Date.parse(value)) || new Date(value).toISOString().slice(0, 10) !== value) {
    fail('INVALID_ARGUMENT', `${name} must be a valid YYYY-MM-DD between 1970-01-01 and 9999-12-30`);
  }
  return value;
}

function checkedInteger(value, field, defaultValue, maximum) {
  const result = value === undefined ? defaultValue : value;
  if (!Number.isInteger(result) || result < 1 || result > maximum) {
    fail('INVALID_ARGUMENT', `${field} must be between 1 and ${maximum}`);
  }
  return result;
}

function scopeOptions(params, defaults = {}) {
  const folder = params.folder === undefined ? (defaults.folder ?? 'inbox') : params.folder;
  const include_subfolders = params.include_subfolders === undefined
    ? (defaults.include_subfolders ?? true) : params.include_subfolders;
  if (typeof folder !== 'string' || !folderPattern.test(folder) || typeof include_subfolders !== 'boolean') {
    fail('INVALID_ARGUMENT', 'Expected a verified folder or store selector and a boolean include_subfolders');
  }
  const scan_limit = checkedInteger(params.scan_limit, 'scan_limit', defaults.scan_limit ?? 100, MAX_SCAN);
  const from_date = params.from_date === undefined ? defaults.from_date : checkedDate(params.from_date, 'from_date');
  const to_date = params.to_date === undefined ? defaults.to_date : checkedDate(params.to_date, 'to_date');
  if (from_date && to_date && from_date > to_date) fail('INVALID_ARGUMENT', 'from_date must not be later than to_date');
  return { folder, include_subfolders, scan_limit, ...(from_date ? { from_date } : {}), ...(to_date ? { to_date } : {}) };
}

function originFor(value = 'http://127.0.0.1:11434') {
  let url;
  try { url = new URL(value); } catch { fail('INVALID_EMBEDDING_ORIGIN', 'Set OUTLOOK_OLLAMA_URL to an HTTP loopback origin'); }
  if (url.protocol !== 'http:' || !['127.0.0.1', '[::1]'].includes(url.hostname) ||
      url.username || url.password || url.pathname !== '/' || url.search || url.hash) {
    fail('INVALID_EMBEDDING_ORIGIN', 'Set OUTLOOK_OLLAMA_URL to an HTTP loopback IP origin');
  }
  return url.origin;
}

function loopbackFetch(url, { method, body, signal }) {
  // http.request connects directly to the validated IP without honoring proxy
  // environment variables or following redirects.
  return new Promise((resolve, reject) => {
    const headers = body === undefined ? {} : {
      'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(body)
    };
    const outbound = request(url, { method, headers, signal, agent: false }, response => {
      const chunks = [];
      let size = 0;
      response.on('data', chunk => {
        size += chunk.length;
        if (size > MAX_RESPONSE_BYTES) {
          response.destroy(new Error('Embedding response exceeds the local bound'));
        } else chunks.push(chunk);
      });
      response.once('error', reject);
      response.once('end', () => resolve({
        status: response.statusCode, ok: response.statusCode === 200,
        json: () => JSON.parse(Buffer.concat(chunks).toString('utf8'))
      }));
    });
    outbound.once('error', reject);
    outbound.end(body);
  });
}

function canonicalModel(name) {
  if (typeof name !== 'string' ||
      !/^(?:[a-z0-9][a-z0-9_.-]{0,63}\/)?[a-z0-9][a-z0-9_.-]{0,127}(?::[a-z0-9][a-z0-9_.-]{0,63})?$/i.test(name) ||
      /(?:[-:]cloud)(?:$|:)/i.test(name)) {
    fail('EMBEDDING_MODEL_REQUIRED', 'Set OUTLOOK_EMBEDDING_MODEL to an installed local Ollama model name');
  }
  return name.includes(':') ? name : `${name}:latest`;
}

async function ollamaProvider(env, fetchImpl) {
  const origin = originFor(env.OUTLOOK_OLLAMA_URL);
  const model = canonicalModel(env.OUTLOOK_EMBEDDING_MODEL);
  async function json(endpoint, payload) {
    const method = payload === undefined ? 'GET' : 'POST';
    let result;
    try {
      result = await fetchImpl(`${origin}${endpoint}`, {
        method, signal: AbortSignal.timeout(120_000), redirect: 'error',
        ...(payload === undefined ? {} : { body: JSON.stringify(payload), headers: { 'Content-Type': 'application/json' } })
      });
    } catch {
      fail('EMBEDDING_UNAVAILABLE', 'Could not connect to the loopback Ollama service');
    }
    if (!result?.ok) {
      fail('EMBEDDING_REQUEST_FAILED', `Local Ollama rejected ${endpoint} (HTTP ${result?.status ?? 'unknown'})`);
    }
    try { return await result.json(); } catch { fail('EMBEDDING_BAD_RESPONSE', 'Local Ollama returned invalid JSON'); }
  }
  const status = await json('/api/status');
  if (status?.cloud?.disabled !== true) {
    fail('EMBEDDING_CLOUD_ENABLED', 'The Ollama server must be running with cloud features disabled');
  }
  const tags = await json('/api/tags');
  const info = Array.isArray(tags?.models) ? tags.models.find(entry => {
    const name = entry?.name ?? entry?.model;
    return name === model || `${name}:latest` === model;
  }) : null;
  if (!info || typeof info.digest !== 'string' || !digestPattern.test(info.digest) ||
      info.remote_host || info.remote_model) {
    fail('EMBEDDING_MODEL_UNAVAILABLE', 'The configured local model and its digest were not found in Ollama');
  }
  return {
    identity: { provider: 'ollama-local', model, digest: info.digest },
    async embed(inputs) {
      const result = await json('/api/embed', { model, input: inputs, truncate: false });
      return result?.embeddings;
    }
  };
}

function validVector(vector) {
  if (!Array.isArray(vector) || vector.length < 1 || vector.length > MAX_DIMENSIONS || !vector.every(Number.isFinite)) {
    return false;
  }
  const norm = Math.hypot(...vector);
  return norm > 0 && Number.isFinite(norm);
}

async function embeddingsFor(env, fetchImpl, stateDir, createAppleProvider) {
  const choice = env.OUTLOOK_EMBEDDING_PROVIDER ??
    (env.OUTLOOK_EMBEDDING_MODEL || process.platform !== 'darwin' ? 'ollama' : 'apple');
  let selected;
  if (choice === 'apple') {
    const create = createAppleProvider ?? (await import('./apple-embedding.mjs')).createAppleEmbeddingProvider;
    const apple = await create({ env, stateDir });
    selected = {
      identity: { provider: 'apple-natural-language', model: apple.model, digest: apple.digest },
      embed: inputs => apple.embed(inputs)
    };
  } else if (choice === 'ollama') {
    selected = await ollamaProvider(env, fetchImpl);
  } else fail('INVALID_ARGUMENT', 'OUTLOOK_EMBEDDING_PROVIDER must be apple or ollama');
  if (typeof selected.identity.model !== 'string' || !selected.identity.model ||
      typeof selected.identity.digest !== 'string' || !digestPattern.test(selected.identity.digest)) {
    fail('EMBEDDING_BAD_RESPONSE', 'Local embedding provider returned an invalid model identity');
  }
  async function embed(inputs, allowUnsupported = false) {
    let vectors;
    try {
      vectors = await selected.embed(inputs);
    } catch (error) {
      if (!allowUnsupported || error.code !== 'EMBEDDING_INPUT_UNSUPPORTED') throw error;
      if (inputs.length === 1) return [null];
      const split = Math.floor(inputs.length / 2);
      return [...await embed(inputs.slice(0, split), true), ...await embed(inputs.slice(split), true)];
    }
    if (!Array.isArray(vectors) || vectors.length !== inputs.length ||
        !vectors.every(vector => vector === null || validVector(vector))) {
      fail('EMBEDDING_BAD_RESPONSE', 'Local embedding provider returned invalid vectors');
    }
    if (!allowUnsupported && vectors.some(vector => vector === null)) {
      fail('EMBEDDING_INPUT_UNSUPPORTED', 'The configured model could not embed the query');
    }
    return vectors.map(vector => {
      if (vector === null) return null;
      const magnitude = Math.hypot(...vector);
      return vector.map(value => value / magnitude);
    });
  }
  return {
    identity: selected.identity,
    embed
  };
}

async function readIndex(stateDir, { create = false } = {}) {
  if (typeof stateDir !== 'string' || !path.isAbsolute(stateDir)) {
    fail('INVALID_ARGUMENT', 'A private absolute semantic state directory is required');
  }
  let directory = await lstat(stateDir).catch(error => {
    if (error.code === 'ENOENT') return null;
    throw error;
  });
  if (directory && !directory.isDirectory()) fail('UNSAFE_SEMANTIC_STATE', 'Semantic state directory must be a real directory');
  if (create) {
    await prepareRuntime(stateDir);
    directory = await lstat(stateDir);
  }
  if (!directory) return null;
  if (!directory.isDirectory() || (process.getuid && (directory.uid !== process.getuid() || directory.mode & 0o077))) {
    fail('UNSAFE_SEMANTIC_STATE', 'Semantic state directory must be private and owned by the current user');
  }
  const source = path.join(stateDir, 'index.json');
  const file = await open(source, constants.O_RDONLY | (constants.O_NOFOLLOW ?? 0) | (constants.O_NONBLOCK ?? 0))
    .catch(error => {
      if (error.code === 'ENOENT') return null;
      throw error;
    });
  if (!file) return null;
  try {
    const state = await file.stat();
    if (!state.isFile() || state.nlink !== 1 || state.size > MAX_INDEX_BYTES ||
        (process.getuid && (state.uid !== process.getuid() || state.mode & 0o077))) {
      fail('UNSAFE_SEMANTIC_STATE', 'Semantic index must be a bounded private regular file');
    }
    const data = Buffer.alloc(state.size);
    for (let start = 0; start < data.length;) {
      const { bytesRead } = await file.read(data, start, data.length - start, start);
      if (!bytesRead) fail('SEMANTIC_INDEX_INVALID', 'Semantic index changed while it was read');
      start += bytesRead;
    }
    let index;
    try { index = JSON.parse(data.toString('utf8')); } catch { fail('SEMANTIC_INDEX_INVALID', 'Semantic index JSON is invalid'); }
    if (!index || index.version !== 1 || !snapshotPattern.test(index.snapshot) || !index.coverage || !index.scope ||
        !['apple-natural-language', 'ollama-local'].includes(index.model?.provider) ||
        typeof index.model.model !== 'string' || !digestPattern.test(index.model.digest) ||
        !Array.isArray(index.records) || index.records.length > MAX_SCAN ||
        !index.records.every(row => row && typeof row.id === 'string' && row.id.length <= 16000 &&
          STATUSES.includes(row.status) && (
            ['indexed', 'truncated'].includes(row.status)
              ? digestPattern.test(row.hash) && validVector(row.vector) : row.hash === undefined && row.vector === undefined
          ))) {
      fail('SEMANTIC_INDEX_INVALID', 'Semantic index format is invalid');
    }
    return index;
  } finally {
    await file.close();
  }
}

async function publishIndex(stateDir, index) {
  const json = JSON.stringify(index);
  if (Buffer.byteLength(json) > MAX_INDEX_BYTES) {
    fail('SEMANTIC_INDEX_TOO_LARGE', 'Reduce scan_limit to keep the semantic index within its local bound');
  }
  const scratch = path.join(stateDir, `index-${randomUUID()}.tmp`);
  const file = await open(scratch, 'wx', 0o600);
  try {
    try { await file.writeFile(json); await file.sync(); } finally { await file.close(); }
    await rename(scratch, path.join(stateDir, 'index.json'));
  } finally {
    await unlink(scratch).catch(error => { if (error.code !== 'ENOENT') throw error; });
  }
}

async function pageFor(call, scope, limit, cursor) {
  const { scan_limit: _cap, ...filters } = scope;
  const page = await call('context_bundle', { ...filters, limit, mode: 'metadata', ...(cursor ? { cursor } : {}) });
  if (!page || page.backend !== 'mac-cache' || !snapshotPattern.test(page.snapshot) || !Array.isArray(page.items) ||
      !page.items.every(row => row && typeof row.id === 'string' && row.id.length <= 16000) ||
      !Number.isSafeInteger(page.total_matched) || page.total_matched < page.items.length) {
    fail('SEMANTIC_SOURCE_INVALID', 'The macOS cache service returned an invalid metadata page');
  }
  return page;
}

async function scan(call, scope) {
  const ids = new Set();
  const limit = Math.min(100, scope.scan_limit);
  let first;
  let cursor;
  let complete;
  do {
    const page = await pageFor(call, scope, limit, cursor);
    first ??= page;
    if ((cursor && (!page.items.length || page.next_cursor === cursor)) || page.snapshot !== first.snapshot ||
        page.total_matched !== first.total_matched) {
      fail('SEMANTIC_SNAPSHOT_CHANGED', 'The metadata snapshot changed during the semantic scan');
    }
    for (const { id } of page.items) {
      if (ids.size >= scope.scan_limit) break;
      if (ids.has(id)) fail('SEMANTIC_SOURCE_INVALID', 'Semantic scan returned a duplicate message ID');
      ids.add(id);
    }
    complete = page.complete === true;
    cursor = page.next_cursor;
    if (!complete && !cursor) fail('SEMANTIC_SOURCE_INVALID', 'Semantic scan stopped before the next metadata page');
  } while (!complete && ids.size < scope.scan_limit);
  return {
    snapshot: first.snapshot,
    ids: [...ids],
    coverage: {
      matched_in_local_metadata: first.total_matched, scanned: ids.size,
      scope_traversal_complete: complete && ids.size === first.total_matched,
      date_filter_incomplete: first.date_filter_incomplete ?? null,
      received_time_unknown: first.received_time_unknown ?? null,
      mailbox_complete: false
    }
  };
}

async function bodies(call, ids, snapshot) {
  const bundle = await call('messages_bundle', {
    ids, max_chars: MAX_BODY, text_budget_chars: BATCH_SIZE * MAX_BODY
  });
  if (!bundle || bundle.snapshot !== snapshot || !Array.isArray(bundle.items)) {
    fail('SEMANTIC_SNAPSHOT_CHANGED', 'The message bundle returned a different snapshot');
  }
  const items = new Map(bundle.items.map(item => [item.id, item]));
  return ids.map(id => {
    const row = items.get(id);
    const body = row?.body;
    if (!body) return { id, status: 'unavailable' };
    if (!body.available || typeof body.text !== 'string') {
      return { id, status: body.status === 'not_cached' ? 'missing' : 'unavailable' };
    }
    if (body.text.length > MAX_BODY) fail('SEMANTIC_SOURCE_INVALID', 'The message bundle exceeded the body limit');
    if (!body.text.trim()) return { id, status: 'empty' };
    const status = body.truncated || body.source_truncated ? 'truncated' : 'indexed';
    const text = `${row.subject ?? ''}\n\n${body.text}`;
    const hash = createHash('sha256').update(JSON.stringify([status, body.total_chars ?? null, text])).digest('hex');
    return { id, status, text, hash, row };
  });
}

function sameModel(left, right) {
  return left?.provider === right.provider && left.model === right.model && left.digest === right.digest;
}

function report(index) {
  if (!index) return { status: 'not_indexed' };
  const statuses = Object.fromEntries(STATUSES.map(status => [status, 0]));
  for (const record of index.records) statuses[record.status]++;
  return {
    status: 'indexed', provider: index.model.provider, model: index.model, scope: index.scope,
    snapshot: index.snapshot, indexed_at: index.indexed_at, max_body_chars: MAX_BODY,
    coverage: { ...index.coverage, ...statuses, body_coverage_complete: statuses.indexed === index.records.length },
    snapshot_verified: false
  };
}

async function synchronize(call, embeddings, stateDir, scope) {
  const previous = await readIndex(stateDir, { create: true });
  const reusable = new Map(sameModel(previous?.model, embeddings.identity)
    ? previous.records.filter(row => row.vector).map(row => [row.hash, row.vector]) : []);
  const { snapshot, ids, coverage } = await scan(call, scope);
  const records = [];
  const unsupported = new Set();
  let reused = 0;
  for (let start = 0; start < ids.length; start += BATCH_SIZE) {
    const batch = await bodies(call, ids.slice(start, start + BATCH_SIZE), snapshot);
    const pending = new Map(batch.filter(row => row.hash && !reusable.has(row.hash) && !unsupported.has(row.hash))
      .map(row => [row.hash, row.text]));
    for (const row of batch) if (row.hash && reusable.has(row.hash)) reused++;
    if (pending.size) {
      const vectors = await embeddings.embed([...pending.values()], true);
      [...pending.keys()].forEach((hash, offset) => {
        if (vectors[offset]) reusable.set(hash, vectors[offset]);
        else unsupported.add(hash);
      });
    }
    records.push(...batch.map(({ id, status, hash }) => {
      if (!hash) return { id, status };
      if (unsupported.has(hash)) return { id, status: 'embedding_unavailable' };
      return { id, status, hash, vector: reusable.get(hash) };
    }));
  }
  const dimensions = new Set(records.filter(row => row.vector).map(row => row.vector.length));
  if (dimensions.size > 1) fail('EMBEDDING_DIMENSION_CHANGED', 'Model embedding dimensions changed during the semantic scan');
  if ((await pageFor(call, scope, 1)).snapshot !== snapshot) {
    fail('SEMANTIC_SNAPSHOT_CHANGED', 'The metadata snapshot changed before the semantic index was ready; run sync again');
  }
  const index = { version: 1, model: embeddings.identity, indexed_at: new Date().toISOString(), scope, snapshot, coverage, records };
  await publishIndex(stateDir, index);
  return { ...report(index), snapshot_verified: true, reused_embeddings: reused };
}

async function search(call, embeddings, index, query, limit) {
  if (!sameModel(index.model, embeddings.identity)) {
    fail('SEMANTIC_MODEL_CHANGED', 'The configured local model differs from the index model; run semantic sync again');
  }
  if ((await pageFor(call, index.scope, 1)).snapshot !== index.snapshot) {
    fail('SEMANTIC_INDEX_STALE', 'The local mailbox snapshot changed; run semantic sync again');
  }
  const indexed = index.records.filter(row => row.vector);
  if (!indexed.length) return { ...report(index), snapshot_verified: true, items: [], skipped_unverifiable: 0 };
  const [queryVector] = await embeddings.embed([query]);
  if (indexed.some(row => row.vector.length !== queryVector.length)) {
    fail('EMBEDDING_DIMENSION_CHANGED', 'Model embedding dimensions differ from the index; run semantic sync again');
  }
  const ranked = indexed.map(row => ({
    ...row, score: row.vector.reduce((total, value, offset) => total + value * queryVector[offset], 0)
  })).sort((left, right) => right.score - left.score);
  const items = [];
  let skipped_unverifiable = 0;
  for (let start = 0; start < ranked.length && items.length < limit; start += BATCH_SIZE) {
    const group = ranked.slice(start, start + BATCH_SIZE);
    const verified = await bodies(call, group.map(row => row.id), index.snapshot);
    for (let offset = 0; offset < group.length && items.length < limit; offset++) {
      const candidate = group[offset];
      const current = verified[offset];
      if (!current.hash || current.hash !== candidate.hash) {
        skipped_unverifiable++;
        continue;
      }
      const { id, subject, sender, received_at, folder_path } = current.row;
      items.push({ id, subject, sender, received_at, folder_path, score: candidate.score,
        body_coverage: candidate.status, verified_current: true });
    }
  }
  if ((await pageFor(call, index.scope, 1)).snapshot !== index.snapshot) {
    fail('SEMANTIC_INDEX_STALE', 'The local mailbox snapshot changed during verification; run semantic sync again');
  }
  return { ...report(index), snapshot_verified: true, items, skipped_unverifiable };
}

export async function runSemanticWorkflow(action, params = {}, {
  call, fetchImpl = loopbackFetch, stateDir, env = process.env, createAppleProvider
} = {}) {
  const allowed = action === 'sync' ? SCOPE_FIELDS : action === 'search' ? [...SCOPE_FIELDS, 'query', 'limit'] :
    action === 'status' ? [] : null;
  if (!allowed || !params || typeof params !== 'object' || Array.isArray(params) ||
      Object.keys(params).some(key => !allowed.includes(key))) {
    fail('INVALID_ARGUMENT', 'Expected semantic sync, search, or status and supported JSON parameters');
  }
  if (action === 'status') return report(await readIndex(stateDir));
  if (typeof call !== 'function') fail('INVALID_ARGUMENT', 'The macOS Outlook bundle service is required');
  if (action === 'sync') {
    const scope = scopeOptions(params);
    return synchronize(call, await embeddingsFor(env, fetchImpl, stateDir, createAppleProvider), stateDir, scope);
  }
  const query = typeof params.query === 'string' ? params.query.trim() : '';
  if (!query || query.length > 1000) fail('INVALID_ARGUMENT', 'query must contain 1 through 1000 characters');
  const limit = checkedInteger(params.limit, 'limit', 10, 50);
  const index = await readIndex(stateDir);
  if (!index) fail('SEMANTIC_INDEX_REQUIRED', 'Run semantic sync for this mailbox scope first');
  if (JSON.stringify(scopeOptions(params, index.scope)) !== JSON.stringify(index.scope)) {
    fail('SEMANTIC_SCOPE_CHANGED', 'The requested mailbox scope differs from the index; run semantic sync for this scope');
  }
  return search(call, await embeddingsFor(env, fetchImpl, stateDir, createAppleProvider), index, query, limit);
}
