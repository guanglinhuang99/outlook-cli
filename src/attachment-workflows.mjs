import { spawn } from 'node:child_process';
import { createHash, randomUUID } from 'node:crypto';
import { link, open, realpath, stat, unlink } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { assertSafeMacPath } from './mac-paths.mjs';
import { call as serviceCall, readState as serviceState } from './runtime.mjs';
import { validateCall } from './tools.mjs';

const fileLimit = 10 * 1024 * 1024;
const chunkLimit = 1024 * 1024;
const batchLimit = 100;
const batchByteLimit = 100 * 1024 * 1024;
const textExtensions = new Set(['.txt', '.csv', '.tsv', '.md', '.log', '.json', '.xml', '.yaml', '.yml']);
const ooxmlHelper = fileURLToPath(new URL('../macos/attachment_text.py', import.meta.url));

function attachmentError(code, message) {
  return Object.assign(new Error(message), { code });
}

function backendOptions({ call = serviceCall, readState = serviceState, platform = process.platform,
  assertSafePath = assertSafeMacPath } = {}) {
  return { call, readState, platform, assertSafePath };
}

function sameState(first, second) {
  return first.port === second.port && first.token === second.token
    && first.pid === second.pid && first.startedAt === second.startedAt
    && JSON.stringify(first.protected_mail_roots) === JSON.stringify(second.protected_mail_roots);
}

async function pinnedBackend(options, output = false) {
  const { call, readState, platform } = options;
  const stateOptions = { requireMacProtectedRoots: output && platform === 'darwin' };
  const first = await readState(stateOptions);
  return {
    state: first,
    call: (method, params) => call(method, params, { state: first }),
    async check() {
      if (!sameState(first, await readState(stateOptions))) {
        throw attachmentError('SERVICE_RESTART_REQUIRED', 'Service state changed while reading attachments');
      }
    }
  };
}

const stateErrors = new Set(['SERVICE_RESTART_REQUIRED', 'SERVICE_UNAVAILABLE', 'OUTPUT_DIR_CHANGED']);

async function listedAttachments(pinned, params) {
  const listing = await pinned.call('attachments', validateCall('attachments', params));
  if (!Array.isArray(listing?.attachments)
      || listing.attachments.some(item => !Number.isInteger(item?.index) || item.index < 1 || item.index > 10000)
      || new Set(listing.attachments.map(item => item.index)).size !== listing.attachments.length) {
    throw attachmentError('ATTACHMENT_INTEGRITY_ERROR', 'Invalid attachment list');
  }
  return listing;
}

async function verifiedBytes(pinned, params, attachment, listing, maxBytes = fileLimit) {
  if (attachment.status && attachment.status !== 'cached_at_snapshot') {
    throw attachmentError('ATTACHMENT_NOT_CACHED', 'Attachment bytes are not cached');
  }
  if (Object.hasOwn(attachment, 'type') && attachment.type !== 1) {
    throw attachmentError('UNSUPPORTED_ATTACHMENT', 'Attachment does not expose a by-value file');
  }
  if (Number.isSafeInteger(attachment.size_bytes) && attachment.size_bytes > maxBytes) {
    throw attachmentError(maxBytes < fileLimit ? 'BATCH_BYTE_LIMIT' : 'ATTACHMENT_TOO_LARGE',
      'Attachment exceeds the remaining byte limit');
  }
  if (maxBytes < 1) throw attachmentError('BATCH_BYTE_LIMIT', 'Batch byte limit exhausted');
  const args = validateCall('read_attachment', {
    ...params, index: attachment.index, max_bytes: maxBytes, chunk_bytes: chunkLimit
  });
  let first;
  let offset = 0;
  const chunks = [];
  const hash = createHash('sha256');
  do {
    const part = await pinned.call('read_attachment', { ...args, byte_offset: offset });
    if (!part || part.encoding !== 'base64' || part.offset !== offset
        || !Number.isSafeInteger(part.size_bytes) || part.size_bytes < offset || part.size_bytes > maxBytes
        || typeof part.sha256 !== 'string' || !/^[a-f0-9]{64}$/.test(part.sha256)
        || typeof part.data !== 'string' || part.data.length > Math.ceil(chunkLimit / 3) * 4
        || (listing.backend && part.backend !== listing.backend)
        || (listing.snapshot && part.snapshot !== listing.snapshot)
        || (attachment.filename && part.filename !== attachment.filename)
        || (Number.isSafeInteger(attachment.size_bytes) && part.size_bytes !== attachment.size_bytes)
        || (first && (part.sha256 !== first.sha256 || part.size_bytes !== first.size_bytes
          || part.filename !== first.filename || part.backend !== first.backend || part.snapshot !== first.snapshot))) {
      throw attachmentError('ATTACHMENT_INTEGRITY_ERROR', 'Attachment identity or size changed while reading');
    }
    const bytes = Buffer.from(part.data, 'base64');
    const end = offset + bytes.length;
    const expectedNext = end === part.size_bytes ? null : end;
    if (bytes.toString('base64') !== part.data || bytes.length !== part.chunk_size
        || bytes.length > chunkLimit || end > part.size_bytes || part.next_offset !== expectedNext
        || (expectedNext !== null && end === offset)) {
      throw attachmentError('ATTACHMENT_INTEGRITY_ERROR', 'Invalid or non-advancing attachment chunk');
    }
    first ??= part;
    chunks.push(bytes);
    hash.update(bytes);
    offset = end;
  } while (offset < first.size_bytes);
  if (hash.digest('hex') !== first.sha256) {
    throw attachmentError('ATTACHMENT_INTEGRITY_ERROR', 'Attachment SHA-256 verification failed');
  }
  return { bytes: Buffer.concat(chunks, first.size_bytes), filename: first.filename,
    size_bytes: first.size_bytes, sha256: first.sha256 };
}

function truncateFilename(value, byteLimit) {
  let used = 0;
  let result = '';
  for (const character of value) {
    used += Buffer.byteLength(character);
    if (used > byteLimit) break;
    result += character;
  }
  return result;
}

function safeFilename(raw, index) {
  const leaf = typeof raw === 'string' ? raw.split(/[/\\]/).at(-1) : '';
  let safe = leaf.normalize('NFC')
    .replace(/[\x00-\x1f\x7f-\x9f<>:"/\\|?*\u202a-\u202e\u2066-\u2069]/g, '_')
    .trim().replace(/^[. ]+|[. ]+$/g, '');
  if (!safe) safe = 'attachment-' + index;
  if (/^(?:con|prn|aux|nul|conin\$|conout\$|clock\$|com[1-9¹²³]|lpt[1-9¹²³])(?:\.|$)/i.test(safe)) safe = '_' + safe;
  const extension = truncateFilename(path.extname(safe), 24);
  const stem = path.extname(safe) ? safe.slice(0, -path.extname(safe).length) : safe;
  return { stem: truncateFilename(stem, 170 - Buffer.byteLength(extension)), extension };
}

async function guardedOutput(pathname, pinned, options) {
  const resolved = await realpath(path.resolve(pathname));
  if (options.platform === 'darwin') {
    const checked = await options.assertSafePath(resolved, { protectedRoots: pinned.state.protected_mail_roots });
    if (checked !== resolved) throw attachmentError('OUTPUT_DIR_CHANGED', 'Output directory was redirected');
  }
  const info = await stat(resolved, { bigint: true });
  if (!info.isDirectory()) throw attachmentError('INVALID_ARGUMENT', 'Output must be an existing directory');
  return { path: resolved, dev: info.dev, ino: info.ino };
}

async function publishAttachment(directory, attachment, index, pinned, options) {
  const checkDirectory = async () => {
    await pinned.check();
    const current = await guardedOutput(directory.path, pinned, options);
    if (current.path !== directory.path || current.dev !== directory.dev || current.ino !== directory.ino) {
      throw attachmentError('OUTPUT_DIR_CHANGED', 'Output directory changed while reading attachments');
    }
  };
  await checkDirectory();
  const pending = path.join(directory.path, '.outlook-download-' + randomUUID());
  let ownPending = false;
  try {
    // The random private staging name is never used as a caller-visible output.
    // Link publishes a complete file and fails if the destination already exists.
    const handle = await open(pending, 'wx', 0o600);
    ownPending = true;
    try { await handle.writeFile(attachment.bytes); } finally { await handle.close(); }
    const { stem, extension } = safeFilename(attachment.filename, index);
    for (let suffix = 1; suffix <= 1000; suffix += 1) {
      await checkDirectory();
      const candidate = path.join(directory.path, stem + (suffix === 1 ? '' : '-' + suffix) + extension);
      try {
        await link(pending, candidate);
        return candidate;
      } catch (error) {
        if (error.code !== 'EEXIST') throw error;
      }
    }
    throw attachmentError('ATTACHMENT_NAME_CONFLICT', 'No available attachment filename');
  } finally {
    if (ownPending) await unlink(pending).catch(() => {});
  }
}

function attachmentFailure(attachment, error) {
  return { index: attachment.index, filename: attachment.filename ?? null,
    code: error.code ?? 'ATTACHMENT_READ_FAILED', message: String(error.message ?? error) };
}

export async function downloadAllAttachments({ id, out_dir, backend }, options) {
  if (typeof out_dir !== 'string' || !out_dir.trim() || out_dir.includes('\0')) {
    throw attachmentError('INVALID_ARGUMENT', 'An existing out_dir is required');
  }
  const args = validateCall('attachments', { id, backend });
  const dependencies = backendOptions(options);
  const pinned = await pinnedBackend(dependencies, true);
  const directory = await guardedOutput(out_dir, pinned, dependencies);
  const listing = await listedAttachments(pinned, args);
  const downloaded = [];
  const errors = [];
  let exportedBytes = 0;
  let stopped;
  for (const [position, item] of listing.attachments.entries()) {
    try {
      if (stopped) throw stopped;
      if (position >= batchLimit) throw attachmentError('BATCH_COUNT_LIMIT', 'Batch attachment limit reached');
      if (exportedBytes >= batchByteLimit) throw attachmentError('BATCH_BYTE_LIMIT', 'Batch byte limit exhausted');
      await pinned.check();
      const { bytes, ...details } = await verifiedBytes(pinned, args, item, listing,
        Math.min(fileLimit, batchByteLimit - exportedBytes));
      const output = await publishAttachment(directory, { bytes, ...details }, item.index, pinned, dependencies);
      exportedBytes += details.size_bytes;
      downloaded.push({ index: item.index, ...details, path: output });
    } catch (error) {
      errors.push(attachmentFailure(item, error));
      if (stateErrors.has(error.code)) stopped = error;
    }
  }
  return { id, backend: listing.backend, data_scope: listing.data_scope, directory: directory.path,
    attachments_complete: listing.attachments_complete ?? null, complete: errors.length === 0,
    size_bytes: exportedBytes, downloaded, errors };
}

function decodeText(bytes, allowPartial = false) {
  const encoding = bytes[0] === 0xff && bytes[1] === 0xfe ? 'utf-16le'
    : bytes[0] === 0xfe && bytes[1] === 0xff ? 'utf-16be' : 'utf-8';
  try {
    const value = new TextDecoder(encoding, { fatal: true }).decode(bytes, { stream: allowPartial });
    if (/[\x00-\x08\x0b\x0e-\x1f\x7f]/.test(value)) throw new Error('Unexpected binary controls');
    return value;
  } catch {
    throw attachmentError('ATTACHMENT_ENCODING_UNSUPPORTED', 'Attachment is not valid UTF-8 or BOM-marked UTF-16 text');
  }
}

async function capture(command, args, input, maxOutput, spawnProcess) {
  return new Promise((resolve, reject) => {
    const env = Object.fromEntries(Object.entries(process.env).filter(([key]) =>
      /^(path|pathext|systemroot|windir|lang|lc_all|tmpdir|temp|tmp)$/i.test(key)));
    const child = spawnProcess(command, args, { env, stdio: ['pipe', 'pipe', 'ignore'], windowsHide: true });
    const chunks = [];
    let used = 0;
    let exceeded = false;
    let timedOut = false;
    const timer = setTimeout(() => { timedOut = true; child.kill('SIGKILL'); }, 15000);
    child.stdin.on('error', () => {});
    child.on('error', error => {
      clearTimeout(timer);
      reject(attachmentError(error.code === 'ENOENT' ? 'ATTACHMENT_EXTRACTOR_UNAVAILABLE' : 'ATTACHMENT_TEXT_FAILED',
        error.code === 'ENOENT' ? 'Required local attachment extractor is unavailable' : 'Could not start local attachment extractor'));
    });
    child.stdout.on('data', data => {
      if (exceeded) return;
      const remaining = maxOutput - used;
      const part = data.subarray(0, remaining);
      used += part.length;
      if (part.length) chunks.push(part);
      if (part.length < data.length) { exceeded = true; child.kill('SIGKILL'); }
    });
    child.on('close', code => {
      clearTimeout(timer);
      if (timedOut) reject(attachmentError('ATTACHMENT_TEXT_TIMEOUT', 'Local attachment extraction timed out'));
      else if (code === 0 || exceeded) resolve({ bytes: Buffer.concat(chunks, used), exceeded });
      else reject(attachmentError('ATTACHMENT_TEXT_FAILED', 'Local extractor could not read this attachment'));
    });
    child.stdin.end(input);
  });
}

async function extractText({ filename, bytes }, maximum, { spawnProcess = spawn,
  python = process.env.OUTLOOK_MAC_PYTHON ?? (process.platform === 'win32' ? 'python' : 'python3'),
  pdfCommand = 'pdftotext' } = {}) {
  const ext = path.extname(filename ?? '').toLowerCase();
  if (textExtensions.has(ext)) return { text: decodeText(bytes), format: 'text', truncated: false };
  if (ext === '.pdf') {
    if (!bytes.subarray(0, 5).equals(Buffer.from('%PDF-'))) {
      throw attachmentError('ATTACHMENT_TEXT_FAILED', 'Attachment is not a PDF document');
    }
    const output = await capture(pdfCommand, ['-enc', 'UTF-8', '-', '-'], bytes, maximum * 4 + 4, spawnProcess);
    const text = decodeText(output.bytes, output.exceeded).trim();
    if (!text && !output.exceeded) {
      throw attachmentError('ATTACHMENT_NO_TEXT', 'PDF contains no extractable text');
    }
    return { text, format: 'pdf', truncated: output.exceeded };
  }
  if (ext === '.docx' || ext === '.xlsx') {
    const format = ext.slice(1);
    const result = await capture(python, ['-I', '-B', ooxmlHelper, format, String(maximum)],
      bytes, maximum * 12 + 32768, spawnProcess);
    let report;
    try {
      if (result.exceeded) throw new Error();
      report = JSON.parse(decodeText(result.bytes));
      if (typeof report.text !== 'string' || typeof report.truncated !== 'boolean') throw new Error();
    } catch {
      throw attachmentError('ATTACHMENT_TEXT_FAILED', 'Local OOXML extractor returned an invalid result');
    }
    return { ...report, format };
  }
  throw attachmentError('ATTACHMENT_FORMAT_UNSUPPORTED', 'Attachment format is not supported for local text extraction');
}

export async function readAttachmentText({ id, index, max_chars = 12000, backend }, options) {
  if (!Number.isInteger(max_chars) || max_chars < 1 || max_chars > 100000) {
    throw attachmentError('INVALID_ARGUMENT', 'max_chars must be 1..100000');
  }
  const args = validateCall('attachments', { id, backend });
  validateCall('read_attachment', { ...args, index });
  const pinned = await pinnedBackend(backendOptions(options));
  const listing = await listedAttachments(pinned, args);
  const attachment = listing.attachments.find(item => item.index === index);
  if (!attachment) throw attachmentError('INVALID_ARGUMENT', 'Attachment index is out of range');
  const { bytes, ...details } = await verifiedBytes(pinned, args, attachment, listing);
  await pinned.check();
  const report = await extractText({ bytes, ...details }, max_chars, options);
  let text = report.text.slice(0, max_chars);
  if (/[\ud800-\udbff]$/.test(text)) text = text.slice(0, -1);
  return { id, index, ...details, backend: listing.backend, data_scope: listing.data_scope,
    format: report.format, text, text_chars: text.length, truncated: report.truncated || text.length < report.text.length };
}
