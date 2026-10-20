import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { readFile, stat } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { convert } from 'html-to-text';

const execute = promisify(execFile);
const root = fileURLToPath(new URL('..', import.meta.url));
const aliases = {
  inbox: ['inbox', '收件箱'], sent: ['sent items', 'sent', '已发送邮件'],
  deleted: ['deleted items', 'trash', '已删除邮件'], drafts: ['drafts', '草稿'],
  junk: ['junk email', 'junk e-mail', '垃圾邮件'], outbox: ['outbox', '发件箱']
};
export const localError = (code, message) => Object.assign(new Error(message), { code });
export function encodeLocalId(kind, store, nid) {
  return Buffer.from(JSON.stringify({ v: 2, backend: 'local', kind, store, nid })).toString('base64url');
}
export function decodeId(value) {
  try { return JSON.parse(Buffer.from(value, 'base64url').toString('utf8')); } catch { return null; }
}
function utc(value) {
  if (!value) return null;
  const normalized = value.replace(' ', 'T');
  const millis = Date.parse(/(?:Z|[+-]\d\d:\d\d)$/.test(normalized) ? normalized : normalized + 'Z');
  return Number.isFinite(millis) ? new Date(millis).toISOString() : null;
}
function sender(name, address) {
  return { name: name ?? null, address: address ?? null, address_type: address?.includes('@') ? 'SMTP' : address ? 'EX' : null, smtp_address: address?.includes('@') ? address : null };
}
export function asText(message) {
  if (message.body === null || message.body_format === 'plain') return message.body;
  if (message.body_format !== 'html') throw localError('OFFLINE_BODY_FORMAT', 'Only RTF is cached; use COM to convert this message');
  return convert(message.body, { wordwrap: false, limits: { maxInputLength: 16000000, maxDepth: 5000, maxChildNodes: 1000000 }, selectors: [{ selector: 'img', format: 'skip' }, { selector: 'a', options: { ignoreHref: true } }] });
}

export class LocalFileBridge {
  constructor({ binary = process.env.OUTLOOK_LOCAL_READER ?? path.join(root, 'runtime/ost-mcp/ost-mcp.exe'), files, run, timeoutMs = 45000 } = {}) {
    this.binary = binary;
    this.files = files;
    this.run = run ?? (async (args, timeout) => (await execute(this.binary, args, { encoding: 'utf8', windowsHide: true, timeout, maxBuffer: 64 * 1024 * 1024 })).stdout);
    this.timeoutMs = timeoutMs;
    this.queue = [];
    this.active = false;
    this.tail = Promise.resolve();
    this.closed = false;
  }
  call(method, params) {
    if (this.closed) return Promise.reject(localError('STOPPING', 'Service stopped'));
    if (this.queue.length >= 32) return Promise.reject(localError('BUSY', 'Local reader queue is full'));
    const ticket = {};
    this.queue.push(ticket);
    const result = this.tail.then(async () => {
      this.queue.splice(this.queue.indexOf(ticket), 1);
      if (this.closed) throw localError('STOPPING', 'Service stopped');
      this.active = true;
      this.deadline = Date.now() + this.timeoutMs;
      try { return await this.perform(method, params); } finally { this.active = false; }
    });
    this.tail = result.catch(() => {});
    return result;
  }
  close() { this.closed = true; }
  async invoke(args, json = true) {
    const remaining = this.deadline - Date.now();
    if (remaining <= 0) throw localError('OFFLINE_TIMEOUT', 'Local file scan exceeded its deadline; reduce scan_limit');
    let text;
    try { text = await this.run(args, remaining); }
    catch (error) {
      if (error.code === 'ENOENT') throw localError('OFFLINE_READER_MISSING', 'Install the bundled OST/PST reader using setup-reader.ps1 or reinstall the application');
      if (error.code === 'EPERM' || error.code === 'EACCES') throw localError('OFFLINE_READER_START', 'Windows or the execution sandbox blocked the parser process. Start the service in the normal desktop user session.');
      if (error.killed) throw localError('OFFLINE_TIMEOUT', 'Local file reader timed out; reduce scan_limit');
      // Do not echo reader stderr, which may contain paths or message content.
      const diagnostic = String(error.stderr ?? '');
      if (/encrypt|crypt|unsupported/i.test(diagnostic)) throw localError('OFFLINE_UNSUPPORTED', 'This OST/PST variant or encryption is not supported by the local parser');
      throw localError('OFFLINE_READ_FAILED', 'Cannot parse this local file; it may be locked, changing, damaged or unsupported. COM remains available when Outlook is open.');
    }
    if (!json) return text;
    try { return JSON.parse(text); } catch { throw localError('OFFLINE_PROTOCOL', 'Local parser returned invalid JSON'); }
  }
  async discover() {
    let paths = this.files;
    if (paths === undefined && process.env.OUTLOOK_LOCAL_FILES) {
      try { paths = JSON.parse(process.env.OUTLOOK_LOCAL_FILES); } catch { throw localError('INVALID_ARGUMENT', 'OUTLOOK_LOCAL_FILES must be a JSON array of absolute OST/PST paths'); }
    }
    if (paths === undefined) {
      try { paths = JSON.parse(await readFile(path.join(root, 'local-stores.json'), 'utf8')); }
      catch (error) { if (error.code !== 'ENOENT') throw localError('INVALID_ARGUMENT', 'local-stores.json must contain a JSON array of absolute OST/PST paths'); }
    }
    if (paths === undefined) {
      const listing = await this.invoke(['--list'], false);
      paths = listing.split(/\r?\n/).map(line => line.match(/\s([A-Za-z]:\\.*\.(?:ost|pst))\s*$/i)?.[1]).filter(Boolean);
    }
    if (!Array.isArray(paths) || paths.length > 100 || paths.some(p => typeof p !== 'string' || !path.isAbsolute(p) || !/\.(ost|pst)$/i.test(p) || p.startsWith('\\\\'))) {
      throw localError('INVALID_ARGUMENT', 'Configure at most 100 absolute local OST/PST paths (network shares are not supported)');
    }
    const stores = [];
    for (const filename of [...new Set(paths.map(p => path.resolve(p)))]) {
      try {
        const info = await stat(filename);
        if (!info.isFile()) continue;
        const key = createHash('sha256').update(filename.toLowerCase() + ':' + info.birthtimeMs).digest('hex');
        stores.push({ key, path: filename, name: path.basename(filename), size: info.size, modified: info.mtime.toISOString() });
      } catch (error) { if (error.code !== 'ENOENT') throw localError('OFFLINE_FILE_ACCESS', 'Cannot access a configured local store'); }
    }
    if (!stores.length) throw localError('OFFLINE_NO_STORES', 'No local OST/PST found. Configure local-stores.json or sync your mailbox using classic Outlook first.');
    return stores;
  }
  async query(store, sql, limit = 10000) { return this.invoke([store.path, '--sql', sql, '--limit', String(limit)]); }
  async folderRows(store) { return this.query(store, 'SELECT * FROM folders'); }
  folderInfo(row, rows, store) {
    return { id: encodeLocalId('folder', store.key, row.nid), name: row.name, path: row.path, item_count: row.item_count, unread_count: row.unread_count, child_count: rows.filter(r => r.parent_nid === row.nid).length };
  }
  async select(stores, value, kind) {
    if (kind === 'folder' && Object.hasOwn(aliases, value)) {
      const store = stores[0];
      const rows = await this.folderRows(store);
      const matches = rows.filter(r => !r.is_search_folder && aliases[value].includes(r.name?.toLowerCase()));
      const standard = matches.filter(r => /(?:^|\/)IPM_SUBTREE\/[^/]+$/.test(r.path));
      const candidates = standard.length ? standard : matches;
      if (candidates.length !== 1) throw localError('OFFLINE_FOLDER_AMBIGUOUS', 'Cannot uniquely resolve this folder alias. Use stores and folders to select an explicit local folder ID.');
      return { store, nid: candidates[0].nid, rows };
    }
    const id = decodeId(value);
    if (id?.v !== 2 || id.backend !== 'local' || id.kind !== kind || !Number.isSafeInteger(id.nid) || id.nid < 1) {
      throw localError('BACKEND_ID_MISMATCH', 'This ID is not a local-file ID. List stores/messages again with backend=local; COM and local IDs are different.');
    }
    const store = stores.find(s => s.key === id.store);
    if (!store) throw localError('OFFLINE_STORE_CHANGED', 'The store for this ID was removed or replaced. List stores/messages again.');
    return { store, nid: id.nid };
  }
  metadata(row, store) {
    return { id: encodeLocalId('message', store.key, row.nid), subject: row.subject ?? null, sender: sender(row.sender_name, row.sender_email), received_at: utc(row.delivered), unread: row.unread, size_bytes: row.size, attachment_count: row.has_attachments === false ? 0 : null, has_attachments: row.has_attachments ?? null, download_state: null };
  }
  async message(store, nid, maxChars = 1) {
    return this.invoke([store.path, '--message', String(nid), '--max-body-chars', String(maxChars)]);
  }
  attachmentMetadata(rows) {
    return rows.map((a, i) => ({ index: i + 1, filename: a.filename, size_bytes: a.data_len ?? a.declared_size, type: a.data_len === null ? null : 1, mime: a.mime }));
  }
  async perform(method, params) {
    const stores = await this.discover();
    // Validate all files touched during this operation after reading. This detects
    // common concurrent writes, not every possible torn-page race in this parser.
    const signatures = new Map();
    const track = async store => { if (signatures.has(store.path)) return; const s = await stat(store.path); signatures.set(store.path, { size: s.size, mtimeMs: s.mtimeMs }); };
    for (const store of stores) await track(store);
    let result;
    if (method === 'health') {
      const checked = [];
      for (const store of stores) {
        await track(store);
        try {
          const info = await this.invoke([store.path, '--info']);
          checked.push({ name: store.name, readable: true, format: info.kind, folders: info.folders, file_modified_at: store.modified });
        } catch (error) { checked.push({ name: store.name, readable: false, error: { code: error.code, message: error.message } }); }
      }
      if (!checked.some(s => s.readable)) throw localError('OFFLINE_NO_READABLE_STORES', `No local store can be read (${checked.map(s => s.error.code).join(', ')})`);
      result = { connected: true, offline: true, outlook_required: false, store_count: stores.length, readable_store_count: checked.filter(s => s.readable).length, stores: checked };
    } else if (method === 'stores') {
      const list = [];
      for (const store of stores) {
        await track(store);
        try {
          const rows = await this.folderRows(store);
          const rootRow = rows.find(r => r.parent_nid === null);
          if (!rootRow) throw localError('OFFLINE_READ_FAILED', 'No root folder in this file');
          list.push({ name: store.name, root_id: encodeLocalId('folder', store.key, rootRow.nid), is_data_file: true, exchange_store_type: null, readable: true, file_modified_at: store.modified });
        } catch (error) { list.push({ name: store.name, readable: false, root_id: null, error: { code: error.code, message: error.message } }); }
      }
      result = { stores: list };
    } else {
      const selected = await this.select(stores, params.id ?? params.folder, params.id ? 'message' : 'folder');
      const { store, nid } = selected;
      await track(store);
      if (method === 'folders') {
        const rows = selected.rows ?? await this.folderRows(store);
        const parent = rows.find(r => r.nid === nid);
        if (!parent) throw localError('INVALID_ID', 'Local folder no longer exists');
        const children = rows.filter(r => r.parent_nid === nid);
        const end = Math.min(children.length, params.offset + params.limit);
        result = { parent: this.folderInfo(parent, rows, store), folders: children.slice(params.offset, end).map(r => this.folderInfo(r, rows, store)), total: children.length, next_offset: end < children.length ? end : null };
      } else if (method === 'messages' || method === 'search') {
        const filters = [`folder_nid = ${nid}`, "(message_class = 'IPM.Note' OR message_class LIKE 'IPM.Note.%')"];
        if (params.unread !== undefined) filters.push(`unread = ${params.unread ? 'true' : 'false'}`);
        if (params.after) filters.push(`delivered >= TIMESTAMP '${new Date(params.after).toISOString().replace('T', ' ').replace('Z', '')}'`);
        if (params.before) filters.push(`delivered < TIMESTAMP '${new Date(params.before).toISOString().replace('T', ' ').replace('Z', '')}'`);
        const budget = Math.min(params.scan_limit, method === 'search' ? 50 : params.scan_limit);
        const rows = await this.query(store, `SELECT *, count(*) OVER() AS total FROM messages WHERE ${filters.join(' AND ')} ORDER BY delivered DESC NULLS LAST, nid DESC LIMIT ${budget + 1} OFFSET ${params.offset}`, budget + 1);
        const matches = [];
        let scanned = 0;
        let incompleteBodies = 0;
        for (const row of rows.slice(0, budget)) {
          scanned++;
          let message;
          let match = true;
          if (method === 'search') {
            const contains = value => String(value ?? '').toLowerCase().includes(params.query.toLowerCase());
            match = [row.subject, row.sender_name, row.sender_email].some(contains);
            if (!match && (!row.sender_name || !row.sender_email || params.search_body)) {
              message = await this.message(store, row.nid, params.search_body ? 2000000 : 1);
              match = [message.sender_name, message.sender_email].some(contains) || (params.search_body && contains(asText(message)));
              if (params.search_body && (message.body_truncated || message.body === null)) incompleteBodies++;
            }
          }
          if (match) {
            const item = this.metadata(row, store);
            if (message) { item.sender = sender(message.sender_name, message.sender_email); item.attachment_count = message.attachments.length; }
            matches.push(item);
            if (matches.length >= params.limit) break;
          }
        }
        const next = scanned < rows.length ? params.offset + scanned : null;
        result = { messages: matches, scanned, scan_limit_applied: budget, filtered_item_count: rows[0]?.total ?? null, next_offset: next, complete: next === null, scope: 'single_folder', pagination: 'live_offset_not_snapshot', body_search_incomplete: incompleteBodies > 0 };
      } else if (method === 'read') {
        if (params.body_offset + params.max_chars > 2000000) throw localError('OFFLINE_BODY_LIMIT', 'Local reader supports the first 2,000,000 body characters per message');
        // Convert a stable source prefix before pagination; independently converting
        // partial HTML at each requested offset would produce shifting text pages.
        const m = await this.message(store, nid, 4000000);
        const actualFormat = m.body_format === 'plain' ? 'text' : m.body_format;
        if (params.format === 'html' && m.body !== null && actualFormat !== 'html') throw localError('OFFLINE_BODY_FORMAT', `Local parser exposes ${actualFormat} for this message, not html. Use COM for other variants.`);
        const fullBody = params.format === 'text' ? asText(m) : m.body;
        const body = fullBody === null ? null : fullBody.slice(0, 2000000);
        const sourceTruncated = m.body_truncated || (fullBody?.length ?? 0) > 2000000;
        const start = body === null ? 0 : Math.min(params.body_offset, body.length);
        const end = body === null ? 0 : Math.min(start + params.max_chars, body.length);
        const next = body !== null && end < body.length ? end : null;
        result = { ...this.metadata(m, store), sender: sender(m.sender_name, m.sender_email), to: m.to, cc: m.cc, attachment_count: m.attachments.length, attachments: this.attachmentMetadata(m.attachments), body: { format: params.format, derived_from: actualFormat, text: body === null ? null : body.slice(start, end), available: body !== null, total_chars: sourceTruncated ? null : body?.length ?? null, offset: start, next_offset: next, truncated: next !== null || sourceTruncated, source_truncated: sourceTruncated } };
      } else {
        const attachments = await this.invoke([store.path, '--attachments', String(nid)]);
        if (method === 'attachments') result = { attachments: this.attachmentMetadata(attachments) };
        else if (method === 'read_attachment') {
          const a = attachments[params.index - 1];
          if (!a) throw localError('INVALID_ARGUMENT', 'Attachment index is out of range');
          if (a.data_len === null) throw localError('UNSUPPORTED_ATTACHMENT', 'Attachment has no locally readable byte stream');
          if (a.data_len > params.max_bytes) throw localError('ATTACHMENT_TOO_LARGE', 'Attachment exceeds max_bytes');
          const raw = await this.invoke([store.path, '--attachment', `${nid}:${a.nid}`, '--max-bytes', String(params.max_bytes)]);
          if (raw.truncated || raw.total_bytes > params.max_bytes) throw localError('ATTACHMENT_TOO_LARGE', 'Attachment exceeds max_bytes');
          const bytes = raw.base64 !== null ? Buffer.from(raw.base64, 'base64') : Buffer.from(raw.text ?? '', 'utf8');
          if (bytes.length !== raw.total_bytes) throw localError('OFFLINE_READ_FAILED', 'Incomplete attachment data');
          const offset = Math.min(params.byte_offset, bytes.length);
          const end = Math.min(offset + params.chunk_bytes, bytes.length);
          result = { filename: raw.filename, size_bytes: bytes.length, sha256: createHash('sha256').update(bytes).digest('hex'), encoding: 'base64', offset, chunk_size: end - offset, next_offset: end < bytes.length ? end : null, data: bytes.subarray(offset, end).toString('base64') };
        } else throw localError('INVALID_ARGUMENT', 'Unknown offline operation');
      }
      result.local_file_modified_at = store.modified;
    }
    for (const [filename, before] of signatures) {
      const after = await stat(filename);
      if (before.size !== after.size || before.mtimeMs !== after.mtimeMs) throw localError('OFFLINE_STORE_CHANGED', 'Local store changed while reading. Retry after Outlook finishes syncing or closes.');
    }
    return { ...result, backend: 'local', data_scope: 'cached_only', limitations: ['Only data already stored locally is available; file modification time is not last-sync time.', 'Parser compatibility is limited; encrypted stores and some formats are unsupported.', 'COM and local-file IDs are different.'] };
  }
}
