import { createHash, randomUUID } from 'node:crypto';
import { constants } from 'node:fs';
import { lstat, open, rename, unlink } from 'node:fs/promises';
import path from 'node:path';
import { prepareRuntime } from './runtime.mjs';

const MAX_SCAN = 5000;
const MAX_OUTPUT = 100;
const BODY_BATCH = 25;
const BODY_BUDGET = 60000;
const CHECKPOINT_BYTES = 1024 * 1024;
const SNAPSHOT = /^[a-f0-9]{24}$/;
const DIGEST = /^[a-f0-9]{64}$/;
const OPAQUE = /^[A-Za-z0-9_-]{8,16000}$/;
const DATE = /^\d{4}-\d{2}-\d{2}$/;
const ALIASES = new Set(['all', 'inbox', 'sent', 'drafts', 'deleted', 'junk', 'outbox']);
const PRESETS = new Set(['actionable', 'risk', 'unread', 'folder-summary']);
const SCOPE_FIELDS = ['folder', 'include_subfolders', 'from_date', 'to_date', 'scan_limit'];
const ACTION_FIELDS = {
  conversation: ['id', 'limit', 'subject_fallback', 'include_body', 'max_chars'],
  'delta-bundle': ['since'],
  'task-bundle': ['preset', 'limit']
};

function fail(code, message) {
  throw Object.assign(new Error(message), { code });
}

function digest(value) {
  return createHash('sha256').update(value).digest('hex');
}

function validDay(value) {
  return typeof value === 'string' && DATE.test(value) && value >= '0001-01-01' &&
    Number.isFinite(Date.parse(value)) && new Date(value).toISOString().slice(0, 10) === value;
}

function boundedInteger(value, name, fallback, ceiling) {
  if (value === undefined) return fallback;
  if (!Number.isInteger(value) || value < 1 || value > ceiling) {
    fail('INVALID_ARGUMENT', `${name} must be an integer in 1..${ceiling}`);
  }
  return value;
}

function boolean(value, name, fallback) {
  if (value === undefined) return fallback;
  if (typeof value !== 'boolean') fail('INVALID_ARGUMENT', `${name} must be boolean`);
  return value;
}

function localId(value, kind) {
  if (typeof value !== 'string' || !OPAQUE.test(value)) return null;
  let parsed;
  try {
    const buffer = Buffer.from(value, 'base64url');
    if (buffer.toString('base64url') !== value) return null;
    parsed = JSON.parse(buffer.toString('utf8'));
  } catch {
    return null;
  }
  if (parsed?.v !== 3 || parsed.backend !== 'mac-cache' || parsed.kind !== kind ||
      typeof parsed.key !== 'string' || !parsed.key || !SNAPSHOT.test(parsed.snapshot ?? '')) return null;
  return parsed;
}

function scopeFor(action, params) {
  const folder = params.folder === undefined ? (action === 'conversation' ? 'all' : 'inbox') : params.folder;
  const folderId = ALIASES.has(folder) ? null : localId(folder, 'folder');
  if (!ALIASES.has(folder) && !folderId) fail('INVALID_ARGUMENT', 'Expected a folder alias or a mac-cache folder/store ID');
  const include_subfolders = boolean(params.include_subfolders, 'include_subfolders', true);
  const { from_date, to_date } = params;
  for (const [name, value] of [['from_date', from_date], ['to_date', to_date]]) {
    if (value !== undefined && !validDay(value)) fail('INVALID_ARGUMENT', `${name} must be a valid YYYY-MM-DD date`);
  }
  if (from_date && to_date && from_date > to_date) {
    fail('INVALID_ARGUMENT', 'from_date must not be later than to_date');
  }
  return {
    folderId,
    scanLimit: boundedInteger(params.scan_limit, 'scan_limit', MAX_SCAN, MAX_SCAN),
    query: { folder, include_subfolders, ...(from_date && { from_date }), ...(to_date && { to_date }) },
    // The hash binds a checkpoint to its full selection without persisting mailbox paths.
    digest: digest(JSON.stringify([folderId ? ['id', folderId.key] : ['alias', folder],
      include_subfolders, from_date ?? null, to_date ?? null]))
  };
}

async function scan(scope, call) {
  const limit = Math.min(100, scope.scanLimit);
  const rows = [];
  const seenRows = new Set();
  const seenCursors = new Set();
  const reasons = new Set();
  let first, cursor = null, traversalComplete = false;
  do {
    const page = await call('context_bundle', {
      ...scope.query, mode: 'metadata', limit, ...(cursor && { cursor })
    });
    if (page?.backend !== 'mac-cache') {
      fail('UNSUPPORTED_BACKEND', 'This workflow requires the macOS Outlook local cache');
    }
    if (!SNAPSHOT.test(page.snapshot ?? '') || !Array.isArray(page.items) || page.items.length > limit ||
        !Number.isSafeInteger(page.total_matched) || page.total_matched < 0 || page.date_basis !== 'received_at' ||
        !validDay(page.scope_from_date) || page.include_subfolders !== scope.query.include_subfolders ||
        !Number.isSafeInteger(page.received_time_unknown) || page.received_time_unknown < 0 ||
        typeof page.date_filter_incomplete !== 'boolean') {
      fail('BAD_RESPONSE', 'The local service returned invalid scan metadata');
    }
    if (first && page.snapshot !== first.snapshot) {
      reasons.add('snapshot_changed_during_scan');
      break;
    }
    first ??= page;
    if (page.total_matched !== first.total_matched || page.scope_from_date !== first.scope_from_date ||
        page.data_scope !== first.data_scope ||
        page.received_time_unknown !== first.received_time_unknown ||
        page.date_filter_incomplete !== first.date_filter_incomplete) {
      reasons.add('scan_metadata_changed');
      break;
    }
    if (page.retry_required || (Array.isArray(page.errors) && page.errors.length)) {
      reasons.add('local_page_incomplete');
      break;
    }
    for (const row of page.items) {
      if (rows.length === scope.scanLimit) break;
      const identity = localId(row?.id, 'message');
      if (!identity || identity.snapshot !== first.snapshot || seenRows.has(identity.key)) {
        reasons.add('invalid_or_duplicate_message_identity');
        continue;
      }
      seenRows.add(identity.key);
      rows.push({ row, key: identity.key });
    }
    cursor = page.next_cursor;
    if (!cursor && page.complete === true && rows.length === first.total_matched) {
      traversalComplete = !reasons.size;
      break;
    }
    if (rows.length === scope.scanLimit) {
      reasons.add('scan_limit_reached');
      break;
    }
    if (typeof cursor !== 'string' || cursor.length > 4000 || seenCursors.has(cursor) || !page.items.length) {
      reasons.add('incomplete_pagination');
      break;
    }
    seenCursors.add(cursor);
    // Each successful page must advance within the caller's bounded scan.
    if (seenCursors.size >= Math.ceil(scope.scanLimit / limit) || reasons.size) {
      reasons.add('incomplete_pagination');
      break;
    }
  } while (cursor);

  if (first.date_filter_incomplete === true) reasons.add('unknown_received_time_excluded');
  return {
    rows,
    coverage: {
      backend: first.backend,
      snapshot: first.snapshot,
      captured_at: first.captured_at ?? null,
      data_scope: first.data_scope ?? 'verified_inbox_sent_cache',
      local_scope_from_date: first.scope_from_date,
      date_basis: first.date_basis,
      ...scope.query,
      scanned: rows.length,
      total_in_local_scope: first.total_matched,
      scan_limit: scope.scanLimit,
      scope_traversal_complete: traversalComplete,
      complete: traversalComplete && reasons.size === 0,
      date_filter_incomplete: first.date_filter_incomplete === true,
      received_time_unknown: first.received_time_unknown ?? null,
      mailbox_complete: false,
      incomplete_reasons: [...reasons]
    }
  };
}

async function currentDeltaScope(scope, snapshot, dataScope, call) {
  const { folder } = scope.query;
  let head, keys;
  if (folder !== 'all' && ALIASES.has(folder)) {
    head = await call('folders', { folder, limit: 1 });
    const resolved = localId(head?.parent?.id, 'folder');
    if (resolved?.snapshot === snapshot) keys = [resolved.key];
  } else {
    head = await call('stores', {});
    if (scope.folderId?.snapshot === snapshot) {
      keys = [scope.folderId.key];
    } else if (folder === 'all' && Array.isArray(head?.stores)) {
      const roots = head.stores.map(store => localId(store?.root_id, 'folder'));
      if (roots.every(root => root?.snapshot === snapshot)) keys = [...new Set(roots.map(root => root.key))].sort();
    }
  }
  if (head?.backend !== 'mac-cache' || head.snapshot !== snapshot ||
      (head.data_scope && head.data_scope !== dataScope)) return { reason: 'snapshot_no_longer_current' };
  if (!keys) return { reason: 'local_scope_identity_unverified' };
  return { sourceDigest: digest(JSON.stringify([dataScope, keys])) };
}

function exactMid(value) {
  return typeof value === 'string' && /^<[^<>\s]+@[^<>\s]+>$/.test(value) ? value : null;
}

function relatedIds(message) {
  return new Set(Array.isArray(message.related_internet_message_ids)
    ? message.related_internet_message_ids.map(exactMid).filter(Boolean) : []);
}

function normalizedSubject(subject) {
  if (typeof subject !== 'string') return '';
  return subject.normalize('NFKC')
    .replace(/^(?:\s*(?:re|fw|fwd|aw|wg|回复|答复|转发)\s*:\s*)+/iu, '')
    .trim().replace(/\s+/gu, ' ').toLowerCase();
}

function append(map, value, entry) {
  if (!value) return;
  const existing = map.get(value);
  if (existing) existing.push(entry);
  else map.set(value, [entry]);
}

function conversationMatches(anchor, rows, subjectFallback) {
  const byMid = new Map();
  const byRelated = new Map();
  for (const entry of rows.some(({ key }) => key === anchor.key) ? rows : [anchor, ...rows]) {
    append(byMid, exactMid(entry.row.internet_message_id), entry);
    for (const mid of relatedIds(entry.row)) append(byRelated, mid, entry);
  }
  const bases = new Map([[anchor.key, 'anchor']]);
  const queue = [anchor];
  const expanded = new Set();
  function visit(entry, basis) {
    if (bases.has(entry.key)) return;
    bases.set(entry.key, basis);
    queue.push(entry);
  }

  for (const entry of queue) {
    const id = exactMid(entry.row.internet_message_id);
    if (id && byMid.get(id).length === 1) {
      for (const reply of byRelated.get(id) ?? []) visit(reply, 'exact_message_id_reference');
    }
    for (const reference of relatedIds(entry.row)) {
      if (expanded.has(reference)) continue;
      expanded.add(reference);
      const owners = byMid.get(reference) ?? [];
      if (owners.length === 1) visit(owners[0], 'exact_message_id_reference');
      if (!owners.length) {
        for (const reply of byRelated.get(reference) ?? []) visit(reply, 'exact_shared_message_id_reference');
      }
    }
  }

  // Cache duplicates with the same InternetMessageId are exact copies by ID, but
  // ambiguous owners never contribute follow-on reference edges.
  for (const entry of queue) {
    const mid = exactMid(entry.row.internet_message_id);
    for (const copy of byMid.get(mid) ?? []) {
      if (!bases.has(copy.key)) bases.set(copy.key, 'exact_internet_message_id');
    }
  }
  if (subjectFallback) {
    const subject = normalizedSubject(anchor.row.subject);
    if (subject.length >= 4) {
      for (const { key, row } of rows) {
        if (bases.has(key) || normalizedSubject(row.subject) !== subject) continue;
        // A same-subject message with an explicit, different reference chain
        // should stay outside the fallback group.
        if (relatedIds(row).size === 0) bases.set(key, 'normalized_subject_fallback');
      }
    }
  }
  return [
    { ...anchor.row, match_basis: 'anchor' },
    ...rows.filter(({ key }) => key !== anchor.key && bases.has(key))
      .map(({ key, row }) => ({ ...row, match_basis: bases.get(key) }))
      .sort((left, right) => Number(left.match_basis === 'normalized_subject_fallback') -
        Number(right.match_basis === 'normalized_subject_fallback'))
  ];
}

async function conversation(params, scope, call) {
  const identity = localId(params.id, 'message');
  if (!identity) fail('INVALID_ARGUMENT', 'A mac-cache message ID is required');
  const subjectFallback = boolean(params.subject_fallback, 'subject_fallback', true);
  const includeBody = boolean(params.include_body, 'include_body', false);
  const maxChars = boundedInteger(params.max_chars, 'max_chars', 4000, 12000);
  const limit = boundedInteger(params.limit, 'limit', 25, MAX_OUTPUT);
  const { rows, coverage } = await scan(scope, call);
  let anchor = rows.find(({ key }) => key === identity.key);
  const anchorInScope = Boolean(anchor);
  if (!anchor) {
    if (coverage.snapshot !== identity.snapshot) {
      fail('SNAPSHOT_MISMATCH', 'The anchor was outside the current scan; list it again on the current snapshot');
    }
    const { body, attachments, ...row } = await call('read', { id: params.id, max_chars: 1 });
    if (row?.id !== params.id) fail('BAD_RESPONSE', 'The local service did not return the requested anchor');
    anchor = { key: identity.key, row };
  }
  const matches = conversationMatches(anchor, rows, subjectFallback);
  let items = matches.slice(0, limit);
  let body = null;
  if (includeBody) {
    // A single bounded batch avoids reading bodies of other candidates or making
    // each selected item consume the full command-level allowance.
    const selected = items.slice(0, BODY_BATCH);
    const response = await call('messages_bundle', {
      ids: selected.map(row => row.id), max_chars: maxChars, text_budget_chars: BODY_BUDGET
    });
    if (response?.backend !== 'mac-cache' || response.snapshot !== coverage.snapshot ||
        !Array.isArray(response.items) || !Array.isArray(response.errors)) {
      fail('BAD_RESPONSE', 'The local service returned bodies for a different conversation snapshot');
    }
    const byId = new Map(response.items.map(row => [row.id, row.body]));
    items = items.map(row => byId.get(row.id) ? { ...row, body: byId.get(row.id) } : row);
    body = {
      requested: items.length, attempted: selected.length, complete: response.body_complete === true &&
        selected.length === items.length && selected.every(row => byId.get(row.id)?.available === true),
      errors: response.errors
    };
  }
  return {
    ...coverage,
    anchor_id: anchor.row.id,
    anchor_in_scan: anchorInScope,
    subject_fallback: subjectFallback,
    matched_in_scan: matches.length - Number(!anchorInScope),
    items,
    truncated: matches.length > limit,
    body
  };
}

function canonical(value) {
  if (Array.isArray(value)) return value.map(canonical);
  if (value === null || typeof value !== 'object') return value ?? null;
  return Object.fromEntries(Object.keys(value).sort().map(key => [key, canonical(value[key])]));
}

function deltaEntry({ key, row }) {
  const stable_id = digest(`mac-cache:message:${key}`);
  const folderKey = localId(row.folder_id, 'folder')?.key ?? null;
  const metadata = canonical({
    internet_message_id: row.internet_message_id, related_internet_message_ids: [...relatedIds(row)].sort(),
    subject: row.subject, subject_status: row.subject_status, sender: row.sender, preview: row.preview,
    sent_at: row.sent_at, received_at: row.received_at, received_time_status: row.received_time_status,
    unread: row.unread, folder_key: folderKey, folder_path: row.folder_path, folder_status: row.folder_status,
    participants: row.participants, participants_status: row.participants_status, attachment_count: row.attachment_count
  });
  return {
    stable_id, fingerprint: digest(JSON.stringify(metadata)),
    item: {
      stable_id, id: row.id, internet_message_id: row.internet_message_id ?? null, subject: row.subject ?? null,
      sender: row.sender ?? null, sent_at: row.sent_at ?? null, received_at: row.received_at ?? null,
      unread: row.unread ?? null, folder_path: row.folder_path ?? null
    }
  };
}

function checkpointDirectory(stateDir) {
  if (typeof stateDir !== 'string' || !path.isAbsolute(stateDir)) {
    fail('INVALID_ARGUMENT', 'A private absolute checkpoint state directory is required');
  }
  return stateDir;
}

async function checkPrivateDirectory(directory) {
  const dir = await lstat(directory);
  if (!dir.isDirectory() || dir.isSymbolicLink() || dir.uid !== process.getuid?.() || (dir.mode & 0o077) !== 0) {
    fail('UNSAFE_CHECKPOINT_STATE', 'The checkpoint directory must be private and owned by the current user');
  }
}

async function readCheckpoint(directory, token) {
  const filename = path.join(directory, `checkpoint-${token}.json`);
  let handle;
  try {
    await checkPrivateDirectory(directory);
    handle = await open(filename, constants.O_RDONLY | (constants.O_NOFOLLOW ?? 0) | (constants.O_NONBLOCK ?? 0));
    const stats = await handle.stat();
    if (!stats.isFile() || stats.uid !== process.getuid?.() || (stats.mode & 0o077) !== 0 || stats.nlink !== 1 ||
        stats.size < 1 || stats.size > CHECKPOINT_BYTES) {
      fail('UNSAFE_CHECKPOINT_STATE', 'The checkpoint must be a private regular file');
    }
    const content = Buffer.alloc(stats.size + 1);
    let offset = 0;
    while (offset < content.length) {
      const { bytesRead } = await handle.read(content, offset, content.length - offset, offset);
      if (!bytesRead) break;
      offset += bytesRead;
    }
    const bytes = content.subarray(0, offset);
    if (offset !== stats.size || digest(bytes) !== token) {
      fail('INVALID_CHECKPOINT', 'The checkpoint content does not match its token');
    }
    let data;
    try { data = JSON.parse(bytes); } catch { fail('INVALID_CHECKPOINT', 'The checkpoint content is invalid'); }
    if (data?.version !== 1 || !DIGEST.test(data.scope ?? '') || !DIGEST.test(data.source_scope ?? '') ||
        !SNAPSHOT.test(data.snapshot ?? '') ||
        !validDay(data.scope_from_date) || !Array.isArray(data.rows) || data.rows.length > MAX_SCAN ||
        !data.rows.every(row => Array.isArray(row) && row.length === 2 &&
          row.every(hash => typeof hash === 'string' && DIGEST.test(hash))) ||
        new Set(data.rows.map(([id]) => id)).size !== data.rows.length) {
      fail('INVALID_CHECKPOINT', 'The checkpoint schema is invalid');
    }
    return data;
  } catch (error) {
    if (error.code === 'ENOENT') fail('CHECKPOINT_NOT_FOUND', 'The local delta checkpoint is unavailable');
    if (error.code === 'ELOOP') fail('UNSAFE_CHECKPOINT_STATE', 'The checkpoint must be a private regular file');
    throw error;
  } finally {
    await handle?.close();
  }
}

async function publishCheckpoint(directory, data) {
  await checkPrivateDirectory(directory).catch(error => { if (error.code !== 'ENOENT') throw error; });
  await prepareRuntime(directory);
  await checkPrivateDirectory(directory);
  const bytes = Buffer.from(JSON.stringify(data));
  const token = digest(bytes);
  const target = path.join(directory, `checkpoint-${token}.json`);
  const stage = path.join(directory, `.checkpoint-${randomUUID()}.tmp`);
  const handle = await open(stage, 'wx', 0o600);
  try {
    try {
      await handle.writeFile(bytes);
      await handle.sync();
    } finally {
      await handle.close();
    }
    // Each target is content-addressed. Concurrent publishers of one token
    // write identical bytes and this rename exposes only a fully written file.
    await rename(stage, target);
  } finally {
    await unlink(stage).catch(error => { if (error.code !== 'ENOENT') throw error; });
  }
  return token;
}

async function delta(params, scope, call, stateDir) {
  if (params.since !== undefined && (typeof params.since !== 'string' || !DIGEST.test(params.since))) {
    fail('INVALID_ARGUMENT', 'since must be a local delta checkpoint token');
  }
  const directory = checkpointDirectory(stateDir);
  const previous = params.since ? await readCheckpoint(directory, params.since) : null;
  if (previous && previous.scope !== scope.digest) {
    fail('CHECKPOINT_SCOPE_MISMATCH', 'Use the same folder, subfolder and received-date filters as the checkpoint');
  }
  const { rows, coverage } = await scan(scope, call);
  let sourceDigest;
  if (coverage.complete) {
    const current = await currentDeltaScope(scope, coverage.snapshot, coverage.data_scope, call);
    sourceDigest = current.sourceDigest;
    if (current.reason) {
      coverage.incomplete_reasons.push(current.reason);
      coverage.complete = false;
    } else if (previous && previous.source_scope !== sourceDigest) {
      fail('CHECKPOINT_SCOPE_MISMATCH', 'The selected Outlook cache account or folder root changed since this checkpoint');
    }
  }
  const priorRows = new Map(previous?.rows);
  const currentRows = new Map(rows.map(entry => {
    const current = deltaEntry(entry);
    return [current.stable_id, current];
  }));
  const added = [];
  const changed = [];
  for (const [id, entry] of currentRows) {
    const prior = priorRows.get(id);
    if (prior === undefined) added.push(entry.item);
    else if (prior !== entry.fingerprint) changed.push(entry.item);
  }
  const sameHorizon = previous?.scope_from_date === coverage.local_scope_from_date;
  const canReportDisappearance = Boolean(previous) && coverage.complete && sameHorizon;
  const removed = canReportDisappearance ? [...priorRows.keys()]
    .filter(id => !currentRows.has(id)).map(stable_id => ({ stable_id })) : [];
  const checkpoint = coverage.complete ? await publishCheckpoint(directory, {
    version: 1, scope: scope.digest, source_scope: sourceDigest, snapshot: coverage.snapshot,
    scope_from_date: coverage.local_scope_from_date,
    rows: [...currentRows.values()].map(({ stable_id, fingerprint }) => [stable_id, fingerprint])
      .sort(([left], [right]) => left.localeCompare(right))
  }) : null;
  return {
    ...coverage,
    baseline: previous === null,
    previous_checkpoint: params.since ?? null,
    checkpoint,
    comparison_basis: 'local_metadata',
    added,
    changed,
    removed_from_local_scope: removed,
    removed_from_local_scope_complete: canReportDisappearance,
    removal_status: !previous ? 'baseline' : !coverage.complete ? 'incomplete_local_scan'
      : !sameHorizon ? 'cache_horizon_changed' : 'compared_local_scope',
    server_deletions_verified: false
  };
}

const TASK_RULES = {
  actionable: [
    ['action_requested', /\b(?:action required|action requested|response required|please (?:approve|review|confirm|respond|reply|submit|sign))\b|请(?:您|各位|大家)?\s*(?:审批|审核|确认|回复|反馈|提交|签署|处理)/iu],
    ['explicit_deadline', /\b(?:deadline|due date)\s*[:：]|\b(?:overdue|past due)\b|(?:回复|提交|反馈|审批)截止|(?:逾期|即将到期)(?:提醒|通知)/iu]
  ],
  risk: [
    ['limit_breach', /\b(?:(?:risk|compliance|regulatory|investment|exposure) limit breach|breach of (?:the )?(?:risk |investment |regulatory )?limit)\b|(?:投资|风险|合规|监管|持仓|集中度|敞口)(?:限额|限制|比例)(?:超限|超标|突破)|(?:超限|超标)(?:预警|告警|通知)/iu],
    ['risk_alert', /\b(?:risk alert|compliance alert|compliance violation|fraud alert|security incident|credit downgrade|payment default)\b|(?:风险|合规)(?:预警|告警|事件)|违规(?:告警|通报)|(?:信用|评级)下调|(?:付款|支付|债券)违约/iu]
  ]
};

function taskReasons(row, preset) {
  if (preset === 'unread') return row.unread === true ? [{ rule: 'verified_unread', field: 'unread' }] : [];
  const reasons = [];
  for (const [rule, expression] of TASK_RULES[preset]) {
    const field = ['subject', 'preview'].find(name => typeof row[name] === 'string' &&
      expression.test(row[name].normalize('NFKC')));
    if (field) reasons.push({ rule, field });
  }
  return reasons;
}

function unreadCounts(rows) {
  let knownUnread = 0, knownRead = 0;
  for (const { row } of rows) {
    if (row.unread === true) knownUnread += 1;
    if (row.unread === false) knownRead += 1;
  }
  return { known_unread: knownUnread, known_read: knownRead, unknown_unread: rows.length - knownUnread - knownRead };
}

function folderSummaries(rows) {
  const folders = new Map();
  for (const entry of rows) {
    const { row } = entry;
    const folder = localId(row.folder_id, 'folder');
    append(folders, folder?.key ?? 'unknown', entry);
  }
  return [...folders.values()].map(entries => {
    const row = entries[0].row;
    return {
      folder_id: row.folder_id ?? null, folder_path: row.folder_path ?? null, folder_status: row.folder_status ?? null,
      count: entries.length, ...unreadCounts(entries),
      known_cached_attachments: entries.filter(({ row }) => row.attachment_count > 0).length,
      latest_received_at: entries.find(({ row }) => row.received_at)?.row.received_at ?? null,
      examples: entries.slice(0, 3).map(({ row }) => ({ id: row.id, subject: row.subject ?? null }))
    };
  }).sort((left, right) => right.count - left.count || (left.folder_path ?? '').localeCompare(right.folder_path ?? ''));
}

async function task(params, scope, call) {
  const preset = params.preset ?? 'actionable';
  if (!PRESETS.has(preset)) fail('INVALID_ARGUMENT', 'preset must be actionable, risk, unread or folder-summary');
  const limit = boundedInteger(params.limit, 'limit', 25, MAX_OUTPUT);
  const { rows, coverage } = await scan(scope, call);
  const counts = unreadCounts(rows);
  if (preset === 'folder-summary') {
    const folders = folderSummaries(rows);
    return { ...coverage, preset, ...counts, total_folders_in_scan: folders.length,
      folders: folders.slice(0, limit), truncated: folders.length > limit };
  }
  const matched = rows.flatMap(({ row }) => {
    const reasons = taskReasons(row, preset);
    return reasons.length ? [{ ...row, reasons }] : [];
  });
  const unknown = preset === 'unread' ? counts.unknown_unread : 0;
  return {
    ...coverage, preset, ...counts,
    selection_basis: preset === 'unread' ? 'verified_unread_state' : 'explicit_phrases_in_subject_or_preview',
    matched: matched.length, unmatched: rows.length - matched.length - unknown, unknown,
    items: matched.slice(0, limit), truncated: matched.length > limit
  };
}

export async function runReadParity(action, params = {}, { call, stateDir } = {}) {
  if (!Object.hasOwn(ACTION_FIELDS, action) || !params || typeof params !== 'object' || Array.isArray(params) ||
      typeof call !== 'function') fail('INVALID_ARGUMENT', 'Invalid local mail workflow request');
  const fields = new Set([...SCOPE_FIELDS, ...ACTION_FIELDS[action]]);
  if (Object.keys(params).some(key => !fields.has(key))) fail('INVALID_ARGUMENT', 'Unknown local mail workflow parameter');
  const scope = scopeFor(action, params);
  if (action === 'conversation') return conversation(params, scope, call);
  if (action === 'delta-bundle') return delta(params, scope, call, stateDir);
  return task(params, scope, call);
}
