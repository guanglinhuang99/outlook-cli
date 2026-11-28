import { createHash, randomUUID } from 'node:crypto';
import { constants } from 'node:fs';
import { link, lstat, open, readdir, realpath, unlink } from 'node:fs/promises';
import path from 'node:path';
import { setTimeout as delay } from 'node:timers/promises';
import { createAppleEmbeddingProvider } from './apple-embedding.mjs';
import { readAttachmentText as extractAttachmentText } from './attachment-workflows.mjs';
import { assertSafeMacPath } from './mac-paths.mjs';
import { call as serviceCall, prepareRuntime, readState as serviceState } from './runtime.mjs';
import { withFullDetail } from './tools.mjs';

const BACKEND = 'mac-cache';
const VERIFIED_SCOPES = new Set(['verified_inbox_sent_cache', 'verified_account_cache']);
const FORMAT = 'outlook-cache-derived-evidence-v1';
const MAX_TEXT = 100000;
const MAX_RECORDS = 2500;
const MAX_SEGMENTS = 256;
const MAX_INDEX_BYTES = 64 * 1024 * 1024;
const MAX_SEGMENT_BYTES = 5 * 1024 * 1024;
const MAX_ATTACHMENT_BYTES = 10 * 1024 * 1024;
const MAX_ATTACHMENT_BATCH_BYTES = 50 * 1024 * 1024;
const snapshotPattern = /^[a-f0-9]{24}$/;
const digestPattern = /^[a-f0-9]{64}$/;
const archivePattern = /^oc1_[a-f0-9]{64}$/;
const object = value => Boolean(value && typeof value === 'object' && !Array.isArray(value));
const sha = value => createHash('sha256').update(value).digest('hex');
const json = value => JSON.stringify(value);

function fail(code, message) {
  throw Object.assign(new Error(message), { code });
}

function shape(params, fields) {
  if (!object(params) || Object.keys(params).some(key => !fields.includes(key))) {
    fail('INVALID_ARGUMENT', 'Expected a JSON object with fields supported by this knowledge action');
  }
}

function number(value, field, fallback, maximum = MAX_TEXT) {
  const result = value === undefined ? fallback : value;
  if (!Number.isSafeInteger(result) || result < 1 || result > maximum) {
    fail('INVALID_ARGUMENT', `${field} must be an integer from 1 to ${maximum}`);
  }
  return result;
}

function take(text, maximum) {
  const prefix = text.slice(0, maximum);
  return /[\ud800-\udbff]$/.test(prefix) ? prefix.slice(0, -1) : prefix;
}

function sourceSnapshot(id) {
  if (typeof id !== 'string' || !/^[A-Za-z0-9_-]{1,16000}$/.test(id)) {
    fail('BACKEND_ID_MISMATCH', 'Expected a mac-cache message ID');
  }
  let value;
  try {
    const decoded = Buffer.from(id, 'base64url');
    if (decoded.toString('base64url') === id) value = JSON.parse(decoded.toString('utf8'));
  } catch {}
  if (value?.v !== 3 || value.backend !== BACKEND || value.kind !== 'message' ||
      typeof value.key !== 'string' || typeof value.snapshot !== 'string' || !snapshotPattern.test(value.snapshot)) {
    fail('BACKEND_ID_MISMATCH', 'Expected a mac-cache message ID');
  }
  return value.snapshot;
}

function sourceError(code) {
  return typeof code === 'string' && /^[A-Z][A-Z0-9_]{0,79}$/.test(code) ? code : 'SOURCE_READ_FAILED';
}

function archiveScopes(records) {
  const data_scopes = [...new Set(records.map(record => record.source.data_scope))].sort();
  const data_scope = data_scopes.length > 1 ? 'mixed_verified_cache' : data_scopes[0] ?? null;
  return { data_scope, data_scopes };
}

async function assemble(params, { call, readAttachmentText, readState }) {
  if (!Array.isArray(params.ids) || !params.ids.length || params.ids.length > 50) {
    fail('INVALID_ARGUMENT', 'ids must contain 1 to 50 mac-cache message IDs');
  }
  const ids = [...new Set(params.ids)];
  const snapshot = sourceSnapshot(ids[0]);
  if (ids.some(id => sourceSnapshot(id) !== snapshot)) {
    fail('SNAPSHOT_MISMATCH', 'All knowledge message IDs must belong to the same cache snapshot');
  }
  const requests = params.attachment_requests === undefined ? [] : params.attachment_requests;
  if (!Array.isArray(requests) || requests.length > 25) {
    fail('INVALID_ARGUMENT', 'attachment_requests must contain at most 25 {id,index} objects');
  }
  const attachmentKeys = new Set();
  for (const request of requests) {
    shape(request, ['id', 'index']);
    if (!ids.includes(request.id) || !Number.isSafeInteger(request.index) ||
        request.index < 1 || request.index > 10000) {
      fail('INVALID_ARGUMENT', 'Attachment requests must target selected IDs and indexes from 1 to 10000');
    }
    const key = [request.id, request.index].join(':');
    if (attachmentKeys.has(key)) fail('INVALID_ARGUMENT', 'Duplicate attachment request');
    attachmentKeys.add(key);
  }

  const bodyMaximum = number(params.body_max_chars, 'body_max_chars', 12000);
  const attachmentMaximum = number(params.attachment_text_max_chars, 'attachment_text_max_chars', 12000);
  const budget = number(params.text_budget_chars, 'text_budget_chars', 60000);
  let used = 0;
  const messages = new Map();
  const errors = [];
  let dataScope, capturedAt, coverage;
  for (let offset = 0; offset < ids.length; offset += 25) {
    const selected = ids.slice(offset, offset + 25);
    const remaining = budget - used;
    const result = await call('messages_bundle', {
      ids: selected, max_chars: Math.max(1, Math.min(bodyMaximum, remaining)), backend: BACKEND,
      text_budget_chars: Math.max(1000, remaining)
    });
    if (result?.backend !== BACKEND || !VERIFIED_SCOPES.has(result.data_scope) || result.snapshot !== snapshot ||
        (dataScope !== undefined && result.data_scope !== dataScope) ||
        !Array.isArray(result.items) || !Array.isArray(result.errors)) {
      fail('CACHE_INTEGRITY_ERROR', 'Message bundle did not return a consistent verified cache snapshot and scope');
    }
    if (offset === 0) {
      dataScope = result.data_scope;
      capturedAt = result.captured_at ?? null;
      coverage = result.coverage ?? null;
    }
    if (capturedAt !== (result.captured_at ?? null)) {
      fail('CACHE_INTEGRITY_ERROR', 'Message bundle capture timestamp changed');
    }
    for (const item of result.items) {
      if (!item || !selected.includes(item.id) || messages.has(item.id)) {
        fail('CACHE_INTEGRITY_ERROR', 'Message bundle returned an unexpected or duplicate message');
      }
      const original = item.body?.available === true && typeof item.body.text === 'string' ? item.body.text : null;
      const text = original === null || used >= budget ? null : take(original, Math.min(bodyMaximum, budget - used));
      used += text?.length ?? 0;
      const { body, ...metadata } = item;
      const error = result.errors.find(entry => entry?.id === item.id);
      const truncated = body?.truncated === true || body?.source_truncated === true ||
        (original !== null && original.length !== text?.length);
      const bodyOutput = {
        format: 'text', available: text !== null, text, truncated,
        status: text === null && (original !== null || error?.code === 'TEXT_BUDGET_EXHAUSTED') ?
          'text_budget_exhausted' : body?.status ?? (error ? 'body_read_failed' : 'unknown'),
        code: error ? sourceError(error.code) : null,
        total_chars: body?.total_chars ?? null,
        derived_from: body?.derived_from ?? null,
        text_sha256: text === null ? null : sha(text)
      };
      messages.set(item.id, { ...metadata, body: bodyOutput });
    }
    for (const error of result.errors) {
      if (!selected.includes(error?.id)) fail('CACHE_INTEGRITY_ERROR', 'Message bundle returned an unexpected failure');
      errors.push({ id: error.id, code: sourceError(error.code) });
    }
    for (const id of selected) {
      if (!messages.has(id) && !errors.some(error => error.id === id)) errors.push({ id, code: 'SOURCE_OMITTED' });
    }
  }

  const attachmentCall = async (method, input, options) => {
    const result = await call(method, input, options);
    if (result?.backend !== BACKEND || result.snapshot !== snapshot || result.data_scope !== dataScope) {
      fail('ATTACHMENT_INTEGRITY_ERROR', 'Attachment RPC did not return the selected cache snapshot and scope');
    }
    return result;
  };
  const attachments = [];
  let attachmentBytes = 0;
  for (const { id, index } of requests) {
    const item = messages.get(id);
    const listed = Array.isArray(item?.attachments) ? item.attachments.find(entry => entry?.index === index) : null;
    const base = { id, index, filename: listed?.filename ?? null, size_bytes: listed?.size_bytes ?? null,
      sha256: null, available: false, text: null, text_sha256: null, truncated: false, format: null,
      status: listed?.status ?? 'unknown', verified_metadata: Boolean(listed) };
    if (!listed || used >= budget) {
      attachments.push({ ...base, code: !item ? 'MESSAGE_UNAVAILABLE' :
        !listed ? 'ATTACHMENT_NOT_VERIFIED' : 'TEXT_BUDGET_EXHAUSTED' });
      continue;
    }
    // Unknown sizes reserve the transport's full per-file bound.
    const charged = Number.isSafeInteger(listed.size_bytes) && listed.size_bytes >= 0 ?
      Math.min(listed.size_bytes, MAX_ATTACHMENT_BYTES) : MAX_ATTACHMENT_BYTES;
    if (attachmentBytes + charged > MAX_ATTACHMENT_BATCH_BYTES) {
      attachments.push({ ...base, code: 'ATTACHMENT_BYTE_BUDGET_EXHAUSTED' });
      continue;
    }
    attachmentBytes += charged;
    try {
      const maximum = Math.min(attachmentMaximum, budget - used);
      const result = await readAttachmentText({ id, index, max_chars: maximum, backend: BACKEND },
        { call: attachmentCall, readState });
      if (result?.backend !== BACKEND || result.data_scope !== dataScope || result.id !== id || result.index !== index ||
          typeof result.text !== 'string' || result.text.length > maximum || typeof result.sha256 !== 'string' ||
          !digestPattern.test(result.sha256) || !Number.isSafeInteger(result.size_bytes) ||
          result.size_bytes < 0 || result.size_bytes > MAX_ATTACHMENT_BYTES ||
          (Number.isSafeInteger(listed.size_bytes) && listed.size_bytes !== result.size_bytes)) {
        fail('ATTACHMENT_INTEGRITY_ERROR', 'Attachment extraction did not match the selected cache item');
      }
      used += result.text.length;
      attachments.push({ ...base, size_bytes: result.size_bytes, filename: result.filename ?? base.filename,
        sha256: result.sha256, available: true, format: result.format, text: result.text,
        text_sha256: sha(result.text), truncated: result.truncated === true, code: null });
    } catch (error) {
      attachments.push({ ...base, code: sourceError(error?.code) });
    }
  }
  const items = ids.flatMap(id => messages.has(id) ? [messages.get(id)] : []);
  return {
    backend: BACKEND, data_scope: dataScope, date_basis: 'received_at', snapshot, captured_at: capturedAt, coverage,
    items, attachments, errors, text_budget_chars: budget, text_chars: used,
    complete: errors.length === 0 && items.every(item => item.body.available && !item.body.truncated) &&
      attachments.every(entry => entry.available && !entry.truncated)
  };
}

function evidence(bundle, kind, item, attachment) {
  const { body, attachments, ...parent } = item;
  const metadata = kind === 'mail' ? { ...parent, attachments: attachments ?? null } : parent;
  const details = kind === 'mail' ? body : attachment;
  const source = {
    backend: BACKEND, data_scope: bundle.data_scope, date_basis: bundle.date_basis,
    snapshot: bundle.snapshot, captured_at: bundle.captured_at,
    message_id: item.id, internet_message_id: metadata.internet_message_id ?? null,
    original_rfc822_available: false,
    ...(attachment ? { attachment_index: attachment.index, attachment_bytes_sha256: attachment.sha256 } : {})
  };
  const data = {
    kind, source, metadata,
    content: {
      format: details.format, available: details.available, text: details.text,
      text_sha256: details.text_sha256, truncated: details.truncated, status: details.status,
      code: details.code ?? null,
      ...(attachment ? { filename: details.filename, size_bytes: details.size_bytes } :
        { derived_from: details.derived_from, total_chars: details.total_chars })
    }
  };
  return { archive_id: 'oc1_' + sha(json(data)), ...data };
}

async function privateDirectory(stateDir, create, assertSafePath) {
  if (typeof stateDir !== 'string' || !path.isAbsolute(stateDir) || stateDir.includes('\0')) {
    fail('INVALID_ARGUMENT', 'A private absolute knowledge state directory is required');
  }
  if (create) {
    try { await prepareRuntime(stateDir); }
    catch (error) {
      if (error.code !== 'INVALID_ARGUMENT' ||
          error.message !== 'runtime state directory is not managed by outlook-local service') throw error;
      // A simultaneous first sync may still be publishing the runtime marker.
      await delay(25);
      await prepareRuntime(stateDir);
    }
  }
  const info = await lstat(stateDir).catch(error => {
    if (error.code === 'ENOENT') return null;
    throw error;
  });
  if (!info) return null;
  if (!info.isDirectory() || (process.getuid && (info.uid !== process.getuid() || (info.mode & 0o077)))) {
    fail('UNSAFE_KNOWLEDGE_STATE', 'Knowledge state directory must be private and owned by the current user');
  }
  return process.platform === 'darwin' ? assertSafePath(stateDir) : realpath(stateDir);
}

const segmentPath = (dir, slot) => path.join(dir, 'segment-' + String(slot).padStart(3, '0') + '.json');

function safeArchiveRecord(record) {
  if (!object(record) || typeof record.archive_id !== 'string' || !archivePattern.test(record.archive_id) ||
      !['mail', 'attachment'].includes(record.kind) ||
      record.source?.backend !== BACKEND || !VERIFIED_SCOPES.has(record.source.data_scope) ||
      !snapshotPattern.test(record.source?.snapshot) || record.source?.original_rfc822_available !== false ||
      !object(record.metadata) || !object(record.content) ||
      record.content.available !== (typeof record.content.text === 'string')) return false;
  const { archive_id, ...payload } = record;
  return archive_id === 'oc1_' + sha(json(payload)) &&
    (record.content.text === null ? record.content.text_sha256 === null :
      typeof record.content.text === 'string' && sha(record.content.text) === record.content.text_sha256);
}

async function loadSegment(file) {
  const handle = await open(file, constants.O_RDONLY | (constants.O_NOFOLLOW ?? 0) | (constants.O_NONBLOCK ?? 0))
    .catch(error => {
      if (error.code === 'ENOENT') return null;
      if (error.code === 'ELOOP') fail('UNSAFE_KNOWLEDGE_STATE', 'Knowledge segment must not be a symbolic link');
      throw error;
    });
  if (!handle) return null;
  try {
    const info = await handle.stat();
    if (!info.isFile() || info.nlink < 1 || info.nlink > 2 || info.size > MAX_SEGMENT_BYTES ||
        (process.getuid && (info.uid !== process.getuid() || (info.mode & 0o077)))) {
      fail('UNSAFE_KNOWLEDGE_STATE', 'Knowledge segment must be a private, bounded regular file');
    }
    const raw = Buffer.alloc(info.size + 1);
    let size = 0;
    while (size < raw.length) {
      const { bytesRead } = await handle.read(raw, size, raw.length - size);
      if (bytesRead === 0) break;
      size += bytesRead;
    }
    let parsed;
    if (size === info.size) {
      try { parsed = JSON.parse(raw.subarray(0, size).toString('utf8')); } catch {}
    }
    if (info.nlink === 2) {
      // A crash after atomic publication can leave the matching private staging
      // name. It is the only extra link accepted for a published segment.
      const stagedInfo = typeof parsed?.pending === 'string' &&
          /^\.outlook-evidence-[a-f0-9-]{36}$/.test(parsed.pending) ?
        await lstat(path.join(path.dirname(file), parsed.pending)).catch(error => {
          if (error.code === 'ENOENT') return null;
          throw error;
        }) : null;
      if ((!stagedInfo || stagedInfo.dev !== info.dev || stagedInfo.ino !== info.ino) &&
          (await handle.stat()).nlink !== 1) {
        fail('UNSAFE_KNOWLEDGE_STATE', 'Knowledge segment has an unexpected additional hard link');
      }
    }
    if (parsed?.version !== 1 || !Array.isArray(parsed.entries) ||
        parsed.digest !== sha(json(parsed.entries)) || !parsed.entries.every(entry => safeArchiveRecord(entry?.record))) {
      fail('KNOWLEDGE_INTEGRITY_ERROR', 'Knowledge segment checksum or evidence was invalid');
    }
    return { entries: parsed.entries, size };
  } finally {
    await handle.close();
  }
}

async function readArchive(directory) {
  const archive = { records: new Map(), slot: 0, bytes: 0 };
  if (!directory) return archive;
  const files = (await readdir(directory)).filter(name => /^segment-\d{3}\.json$/.test(name)).sort();
  if (files.length > MAX_SEGMENTS) fail('KNOWLEDGE_CAPACITY', 'Knowledge archive exceeds the segment limit');
  for (const name of files) {
    const file = segmentPath(directory, archive.slot);
    if (path.basename(file) !== name) fail('KNOWLEDGE_INTEGRITY_ERROR', 'Knowledge archive has a missing segment');
    const segment = await loadSegment(file);
    if (!segment) fail('KNOWLEDGE_INTEGRITY_ERROR', 'Knowledge archive segment disappeared');
    archive.slot += 1;
    archive.bytes += segment.size;
    if (archive.bytes > MAX_INDEX_BYTES) fail('KNOWLEDGE_CAPACITY', 'Knowledge archive exceeds the storage bound');
    for (const entry of segment.entries) archive.records.set(entry.record.archive_id, entry);
    if (archive.records.size > MAX_RECORDS) fail('KNOWLEDGE_CAPACITY', 'Knowledge archive exceeds the evidence limit');
  }
  return archive;
}

async function staged(directory, data, publish, pending = '.outlook-evidence-' + randomUUID()) {
  const file = path.join(directory, pending);
  const handle = await open(file, 'wx', 0o600);
  try {
    try {
      await handle.writeFile(data);
      await handle.sync();
    } finally { await handle.close(); }
    return await publish(file);
  } finally {
    await unlink(file).catch(() => {});
  }
}

async function appendArchive(directory, entries) {
  // Slots are ordered and published with exclusive links; a losing writer
  // reloads the winner's segment and recalculates both capacity and duplicates.
  for (;;) {
    const archive = await readArchive(directory);
    const changes = entries.filter(entry => {
      const existing = archive.records.get(entry.record.archive_id);
      return !existing || (entry.embedding && json(entry.embedding) !== json(existing.embedding));
    });
    if (!changes.length) return { stored: 0, count: archive.records.size };
    const count = archive.records.size +
      changes.filter(entry => !archive.records.has(entry.record.archive_id)).length;
    const pending = '.outlook-evidence-' + randomUUID();
    const data = json({ version: 1, pending, entries: changes, digest: sha(json(changes)) });
    const bytes = Buffer.byteLength(data);
    if (count > MAX_RECORDS || archive.slot >= MAX_SEGMENTS || bytes > MAX_SEGMENT_BYTES ||
        archive.bytes + bytes > MAX_INDEX_BYTES) {
      fail('KNOWLEDGE_CAPACITY', 'Knowledge archive capacity reached; select less evidence or use a new service state directory');
    }
    try {
      await staged(directory, data, file => link(file, segmentPath(directory, archive.slot)), pending);
      return { stored: changes.length, count };
    } catch (error) {
      if (error.code !== 'EEXIST') throw error;
    }
  }
}

function searchable(record) {
  return [record.metadata.subject, record.metadata.sender?.name, record.metadata.sender?.smtp_address,
    record.content.filename, record.content.text].filter(value => typeof value === 'string').join('\n');
}

async function providerFor({ createAppleProvider, env }) {
  const provider = await createAppleProvider({ env });
  if (typeof provider?.model !== 'string' || !provider.model || typeof provider.digest !== 'string' ||
      !digestPattern.test(provider.digest) || typeof provider.embed !== 'function') {
    fail('EMBEDDING_BAD_RESPONSE', 'Local Apple embedding provider returned an invalid model identity');
  }
  return provider;
}

async function vectorsFor(provider, texts) {
  const results = await provider.embed(texts);
  if (!Array.isArray(results) || results.length !== texts.length) {
    fail('EMBEDDING_BAD_RESPONSE', 'Local Apple embedding provider returned invalid vectors');
  }
  return results.map(vector => {
    if (vector === null) return null;
    if (!Array.isArray(vector) || !vector.length || vector.length > 4096 ||
        !vector.every(value => typeof value === 'number' && Number.isFinite(value))) {
      fail('EMBEDDING_BAD_RESPONSE', 'Local Apple embedding provider returned an invalid vector');
    }
    const norm = Math.hypot(...vector);
    if (!norm || !Number.isFinite(norm)) fail('EMBEDDING_BAD_RESPONSE', 'Local Apple embedding provider returned an invalid vector');
    return vector.map(value => value / norm);
  });
}

async function attachVectors(entries, options) {
  const provider = await providerFor(options);
  const candidates = entries.filter(entry => searchable(entry.record).trim().length > 0);
  let embedded = 0;
  for (let offset = 0; offset < candidates.length; offset += 8) {
    const batch = candidates.slice(offset, offset + 8);
    // Keep long extracted text in evidence while bounding each model input.
    const vectors = await vectorsFor(provider, batch.map(entry => take(searchable(entry.record), 6000)));
    batch.forEach((entry, index) => {
      const vector = vectors[index];
      if (vector) {
        entry.embedding = { model: provider.model, digest: provider.digest, vector };
        embedded += 1;
      }
    });
  }
  return { model: provider.model, digest: provider.digest, embedded, without_vector: entries.length - embedded,
    input_max_chars: 6000 };
}

function summary(record) {
  return {
    archive_id: record.archive_id, kind: record.kind, source: record.source,
    subject: record.metadata.subject ?? null, received_at: record.metadata.received_at ?? null,
    filename: record.content.filename ?? null, text_available: record.content.available,
    text_sha256: record.content.text_sha256, text_status: record.content.status,
    text_truncated: record.content.truncated, text_code: record.content.code
  };
}

async function search(archive, params, options) {
  const { query, kind = 'all', mode = 'literal' } = params;
  if (typeof query !== 'string' || query.length > 500 || !query.trim() ||
      !['all', 'mail', 'attachment'].includes(kind) || !['literal', 'semantic', 'hybrid'].includes(mode)) {
    fail('INVALID_ARGUMENT', 'Provide a query of 1 to 500 characters, a valid kind and a valid search mode');
  }
  const limit = number(params.limit, 'limit', 20, 100);
  const candidates = [...archive.records.values()].filter(entry => kind === 'all' || entry.record.kind === kind);
  let provider, queryVector, semanticError;
  if (mode !== 'literal' && candidates.some(entry => entry.embedding)) {
    try {
      provider = await providerFor(options);
      if (candidates.some(({ embedding }) => embedding?.model === provider.model && embedding.digest === provider.digest)) {
        [queryVector] = await vectorsFor(provider, [query.trim()]);
        if (!queryVector) fail('EMBEDDING_INPUT_UNSUPPORTED', 'The local Apple model could not embed the search query');
      }
    } catch (error) {
      if (mode !== 'hybrid' || !['EMBEDDING_PROVIDER_UNAVAILABLE', 'EMBEDDING_INPUT_UNSUPPORTED'].includes(error.code)) {
        throw error;
      }
      semanticError = error.code;
    }
  }
  const needle = query.trim().toLowerCase();
  const hits = [];
  let embedded = 0;
  for (const { record, embedding } of candidates) {
    const document = searchable(record);
    const position = document.toLowerCase().indexOf(needle);
    const literal = position !== -1;
    let semantic = null;
    if (queryVector && embedding?.model === provider.model && embedding.digest === provider.digest &&
        Array.isArray(embedding.vector) && embedding.vector.length === queryVector.length &&
        embedding.vector.every(Number.isFinite)) {
      embedded += 1;
      const similarity = embedding.vector.reduce((sum, value, index) => sum + value * queryVector[index], 0);
      semantic = Math.max(-1, Math.min(1, similarity));
    }
    if ((mode === 'literal' && literal) || (mode === 'semantic' && semantic !== null) ||
        (mode === 'hybrid' && (literal || semantic !== null))) {
      const start = mode !== 'semantic' && literal ? Math.max(0, position - 60) : 0;
      const score = mode === 'literal' ? 1 : (mode === 'hybrid' && literal ? 2 : 0) + (semantic ?? 0);
      hits.push({ ...summary(record), literal_match: mode === 'semantic' ? null : literal,
        semantic_score: semantic, score, snippet: take(document.slice(start), 240) });
    }
  }
  hits.sort((a, b) => b.score - a.score || (b.received_at ?? '').localeCompare(a.received_at ?? '') ||
    a.archive_id.localeCompare(b.archive_id));
  return {
    mode, kind, ...archiveScopes(candidates.map(entry => entry.record)), scanned: candidates.length, matches: hits.length,
    semantic_coverage: mode === 'literal' ? null : { eligible: candidates.length, indexed: embedded,
      complete: embedded === candidates.length, ...(semanticError ? { error: semanticError } : {}) },
    items: hits.slice(0, limit)
  };
}

function getRecord(archive, id) {
  if (typeof id !== 'string' || !archivePattern.test(id)) fail('INVALID_ARGUMENT', 'Expected a knowledge archive ID');
  const entry = archive.records.get(id);
  if (!entry) fail('KNOWLEDGE_NOT_FOUND', 'Evidence ID is absent from this local knowledge archive');
  return entry.record;
}

async function exportRecords(archive, params, { assertSafePath, readState, platform }) {
  if (!Array.isArray(params.ids) || !params.ids.length || params.ids.length > 50 ||
      typeof params.out_dir !== 'string' || !params.out_dir.trim() || params.out_dir.includes('\0')) {
    fail('INVALID_ARGUMENT', 'Provide 1 to 50 archive IDs and an explicit existing output directory');
  }
  const records = [...new Set(params.ids)].map(id => getRecord(archive, id));
  const input = path.resolve(params.out_dir);
  let protectedRoots;
  if (platform === 'darwin') {
    await assertSafePath(input);
    const state = await readState({ requireMacProtectedRoots: true }).catch(error => {
      if (error.code === 'SERVICE_UNAVAILABLE' || error.code === 'SERVICE_NOT_RUNNING') return null;
      throw error;
    });
    protectedRoots = state?.protected_mail_roots ?? [];
  }
  const guard = async () => {
    const resolved = await realpath(input);
    if (platform === 'darwin' && await assertSafePath(resolved, { protectedRoots }) !== resolved) {
      fail('OUTPUT_DIR_CHANGED', 'Evidence export directory was redirected');
    }
    const info = await lstat(resolved, { bigint: true });
    if (!info.isDirectory()) fail('INVALID_ARGUMENT', 'Evidence export requires an existing directory');
    return { resolved, info };
  };
  const { resolved, info } = await guard();
  const manifest = records.map(record => ({ archive_id: record.archive_id, source: record.source,
    text_sha256: record.content.text_sha256, text_status: record.content.status }));
  const scopes = archiveScopes(records);
  const payload = { format: FORMAT, original_rfc822_available: false, ...scopes, manifest, records };
  const data = json({ ...payload, payload_sha256: sha(json(payload)) }) + '\n';
  const digest = sha(data);
  const target = path.join(resolved, 'outlook-evidence-' + digest + '.json');
  await staged(resolved, data, async file => {
    const current = await guard();
    if (current.resolved !== resolved || current.info.dev !== info.dev || current.info.ino !== info.ino) {
      fail('OUTPUT_DIR_CHANGED', 'Evidence export directory changed during publication');
    }
    await link(file, target).catch(error => {
      if (error.code === 'EEXIST') fail('KNOWLEDGE_EXPORT_EXISTS', 'An evidence export already exists at the generated filename');
      throw error;
    });
  });
  return { path: target, format: FORMAT, original_rfc822_available: false,
    ...scopes, records: records.length, sha256: digest };
}

const bundleFields = ['ids', 'attachment_requests', 'body_max_chars', 'attachment_text_max_chars', 'text_budget_chars'];

export async function runKnowledgeWorkflow(action, params = {}, {
  stateDir, call = (method, input, options) => serviceCall(method, withFullDetail(method, input), options), readAttachmentText = extractAttachmentText, readState = serviceState,
  createAppleProvider = createAppleEmbeddingProvider, env = process.env, assertSafePath = assertSafeMacPath,
  platform = process.platform
} = {}) {
  if (action === 'bundle') {
    shape(params, bundleFields);
    return assemble(params, { call, readAttachmentText, readState });
  }
  if (action === 'sync' || action === 'index') {
    shape(params, [...bundleFields, 'semantic']);
    if (params.semantic !== undefined && typeof params.semantic !== 'boolean') {
      fail('INVALID_ARGUMENT', 'semantic must be boolean');
    }
    const bundle = await assemble(params, { call, readAttachmentText, readState });
    const byId = new Map(bundle.items.map(item => [item.id, item]));
    const entries = [...bundle.items.map(item => evidence(bundle, 'mail', item)),
      ...bundle.attachments.filter(item => item.verified_metadata)
        .map(item => evidence(bundle, 'attachment', byId.get(item.id), item))]
      .map(record => ({ record }));
    const embedding = params.semantic && entries.length ? await attachVectors(entries, { createAppleProvider, env }) : null;
    const directory = await privateDirectory(stateDir, entries.length > 0, assertSafePath);
    const update = entries.length ? await appendArchive(directory, entries) :
      { stored: 0, count: (await readArchive(directory)).records.size };
    return { ...update, backend: BACKEND, data_scope: bundle.data_scope, snapshot: bundle.snapshot,
      captured_at: bundle.captured_at, coverage: bundle.coverage, evidence: entries.map(entry => summary(entry.record)),
      errors: bundle.errors, attachment_errors: bundle.attachments.filter(item => item.code)
        .map(({ id, index, code }) => ({ id, index, code })),
      complete: bundle.complete, semantic: embedding };
  }
  const fields = { status: [], search: ['query', 'kind', 'mode', 'limit'], read: ['archive_id'],
    source: ['archive_id'], export: ['ids', 'out_dir'] };
  if (!Object.hasOwn(fields, action)) fail('INVALID_ARGUMENT', 'Unknown knowledge action');
  shape(params, fields[action]);
  const directory = await privateDirectory(stateDir, false, assertSafePath);
  const archive = await readArchive(directory);
  if (action === 'status') {
    const entries = [...archive.records.values()];
    return { backend: BACKEND, ...archiveScopes(entries.map(entry => entry.record)),
      count: entries.length, index_bytes: archive.bytes,
      text_available: entries.filter(entry => entry.record.content.available).length,
      embedded: entries.filter(entry => entry.embedding).length,
      snapshots: [...new Set(entries.map(entry => entry.record.source.snapshot))].sort(),
      capacity: { records: MAX_RECORDS, segments: MAX_SEGMENTS, bytes: MAX_INDEX_BYTES } };
  }
  if (action === 'search') return search(archive, params, { createAppleProvider, env });
  if (action === 'export') return exportRecords(archive, params, { assertSafePath, readState, platform });
  const record = getRecord(archive, params.archive_id);
  return action === 'read' ? record : summary(record);
}
