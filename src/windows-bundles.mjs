import { createHash } from 'node:crypto';
import { localError } from './offline.mjs';
import { formatDay, parseDay, singaporeDay, singaporeMidnight, singaporeTimezone } from './singapore-dates.mjs';

const folderLimit = 300;
const pageCallLimit = 8;
const cursorLimit = 4096;
const cursorMismatch = 'cursor belongs to another folder or query, or the folder list changed; start again without cursor';

function readCursor(encoded, expected) {
  if (!encoded) return null;
  let cursor;
  try {
    if (encoded.length <= cursorLimit && /^[A-Za-z0-9_-]+$/.test(encoded)) {
      cursor = JSON.parse(Buffer.from(encoded, 'base64url').toString('utf8'));
    }
  } catch {
    throw localError('INVALID_ARGUMENT', cursorMismatch);
  }
  if (cursor?.v !== 1 || cursor.m !== expected.m || cursor.q !== expected.q ||
      cursor.f !== expected.f || cursor.t !== expected.t ||
      typeof cursor.fid !== 'string' || !cursor.fid ||
      !Number.isSafeInteger(cursor.o) || cursor.o < 0) {
    throw localError('INVALID_ARGUMENT', cursorMismatch);
  }
  return cursor;
}

function writeCursor(fields, folder, offset) {
  const encoded = Buffer.from(JSON.stringify({ v: 1, ...fields, fid: folder.id, o: offset })).toString('base64url');
  if (encoded.length > cursorLimit) throw localError('INVALID_ARGUMENT', 'Folder ID is too long for a bundle cursor');
  return encoded;
}

async function expandFolders(folder, includeSubfolders, call) {
  const folders = [];
  const seen = new Set();
  let truncated = false;
  function add(entry) {
    if (seen.has(entry.id)) return;
    if (folders.length === folderLimit) {
      truncated = true;
      return;
    }
    seen.add(entry.id);
    folders.push(entry);
  }

  for (const root of folder === 'all' ? ['inbox', 'sent'] : [folder]) {
    const start = folders.length;
    const first = await call('folders', { folder: root, offset: 0, limit: includeSubfolders ? 200 : 1 });
    add(first.parent);
    if (!includeSubfolders || folders.length === start) continue;

    for (let index = start; index < folders.length && !truncated; index += 1) {
      const parent = folders[index];
      if (index !== start && parent.child_count === 0) continue;
      if (index !== start && folders.length === folderLimit && parent.child_count > 0) {
        truncated = true;
        break;
      }
      let page = index === start ? first : await call('folders', { folder: parent.id, offset: 0, limit: 200 });
      for (;;) {
        for (const child of page.folders) add(child);
        if (truncated || page.next_offset == null) break;
        if (folders.length === folderLimit) {
          truncated = true;
          break;
        }
        page = await call('folders', { folder: parent.id, offset: page.next_offset, limit: 200 });
      }
    }
    if (truncated) break;
  }
  return { folders, truncated };
}

async function readBodies(items, result, params, call, keepMetadata) {
  for (const item of items) {
    const { id } = item;
    const remaining = params.text_budget_chars - result.body_text_chars;
    if (remaining <= 0) {
      result.errors.push({ id, code: 'TEXT_BUDGET_EXHAUSTED', message: 'Bundle text budget exhausted' });
    } else {
      try {
        const read = await call('read', {
          id, format: 'text', body_offset: 0, max_chars: Math.min(params.max_chars, remaining)
        });
        const { body, to, cc, attachments } = read;
        result.body_text_chars += body?.text?.length ?? 0;
        result.body_complete &&= typeof body?.text === 'string' && body.available !== false &&
          !body.truncated && !body.source_truncated && body.next_offset == null;
        if (keepMetadata) {
          result.items.push({ ...item, body, to, cc, attachments });
        } else {
          const { backend, data_scope, warnings, limitations, date_basis, ...message } = read;
          result.items.push(message);
        }
        continue;
      } catch (error) {
        result.errors.push({ id, code: error.code ?? 'BODY_READ_FAILED', message: error.message });
      }
    }
    result.body_complete = false;
    if (keepMetadata) result.items.push(item);
  }
}

export async function runWindowsBundle(method, params, call) {
  const batch = method === 'messages_bundle';
  const today = method === 'today_bundle';
  const mode = batch ? 'context' : params.mode;
  const warnings = new Map();
  const result = {
    backend: null, data_scope: null, items: [], errors: [], body_text_chars: 0,
    body_complete: mode === 'metadata' ? null : true,
    mode, folder: batch ? null : params.folder, folders_scanned: 0, folders_truncated: false,
    retry_required: false, body_search_incomplete: false, person_search_incomplete: false,
    complete: true, date_basis: 'received_at', timezone: singaporeTimezone
  };

  async function invoke(operation, input) {
    const response = await call(operation, input);
    for (const warning of response.warnings ?? []) {
      if (!warnings.has(warning.code)) warnings.set(warning.code, warning);
    }
    for (const key of ['backend', 'data_scope']) {
      if (response[key] == null) continue;
      if (result[key] !== null && result[key] !== response[key]) {
        warnings.set('MIXED_BACKENDS', {
          code: 'MIXED_BACKENDS', message: 'Bundle operations returned different backends or data scopes.'
        });
      }
      result[key] = response[key];
    }
    return response;
  }

  if (batch) {
    const items = [...new Set(params.ids)].map(id => ({ id }));
    await readBodies(items, result, params, invoke, false);
    return { ...result, warnings: [...warnings.values()] };
  }
  if (params.field !== 'all') throw localError('INVALID_ARGUMENT', 'Windows bundles only support field=all');
  if (params.has_cached_attachments !== undefined) {
    throw localError('INVALID_ARGUMENT', 'Windows bundles do not support has_cached_attachments');
  }

  const { folder, include_subfolders, query, person, person_role, unread, search_body } = params;
  const q = createHash('sha256').update(JSON.stringify([
    folder, include_subfolders, query, person, person_role, unread, search_body
  ])).digest('base64url');
  const fields = { m: method, d: null, f: params.from_date ?? null, t: params.to_date ?? null, q };
  const cursor = readCursor(params.cursor, fields);
  if (today) {
    const date = cursor ? cursor.d : params.date ?? singaporeDay(Date.now());
    parseDay(date);
    if (cursor && params.date && params.date !== date) {
      throw localError('INVALID_ARGUMENT', 'date does not match the cursor date');
    }
    fields.d = date;
    result.date = date;
  } else if (cursor && cursor.d !== null) {
    throw localError('INVALID_ARGUMENT', cursorMismatch);
  }
  const first = fields.d ?? fields.f;
  const last = fields.d ?? fields.t;
  const after = first ? singaporeMidnight(first) : undefined;
  const before = last ? singaporeMidnight(formatDay(new Date(parseDay(last).getTime() + 86400000)), true) : undefined;

  const { folders, truncated } = await expandFolders(folder, include_subfolders, async (operation, input) => {
    try {
      return await invoke(operation, input);
    } catch (error) {
      if (cursor && error.code === 'INVALID_ID') throw localError('INVALID_ARGUMENT', cursorMismatch);
      throw error;
    }
  });
  result.folders_truncated = truncated;
  let index = cursor ? folders.findIndex(entry => entry.id === cursor.fid) : 0;
  if (index < 0) throw localError('INVALID_ARGUMENT', cursorMismatch);
  let offset = cursor?.o ?? 0;
  const selected = [];
  const search = Boolean(query || person);
  for (; index < folders.length && selected.length < params.limit && result.folders_scanned < pageCallLimit; index += 1) {
    const { id, name, path } = folders[index];
    const input = { folder: id, offset, limit: params.limit - selected.length, scan_limit: 500, unread, after, before };
    const page = await invoke(search ? 'search' : 'messages',
      search ? { ...input, query, person, person_role, search_body } : input);
    result.folders_scanned += 1;
    result.body_search_incomplete ||= Boolean(page.body_search_incomplete);
    result.person_search_incomplete ||= Boolean(page.person_search_incomplete);
    if (page.retry_required) {
      // The same folder position is repeated; returning its partial hits now would duplicate them.
      result.retry_required = true;
      break;
    }
    selected.push(...page.messages.map(item => ({ ...item, folder: { id, name, path } })));
    if (page.next_offset != null) {
      offset = page.next_offset;
      break;
    }
    offset = 0;
  }
  result.next_cursor = index < folders.length ? writeCursor(fields, folders[index], offset) : null;
  result.complete = result.next_cursor === null && !result.retry_required;
  if (mode === 'metadata') result.items = selected;
  else {
    await readBodies(selected, result, params, invoke, true);
    if (result.retry_required) result.body_complete = false;
  }
  return { ...result, warnings: [...warnings.values()] };
}
