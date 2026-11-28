import { createHash, randomBytes } from 'node:crypto';
import { constants } from 'node:fs';
import { lstat, open, realpath, rename, unlink } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { z } from 'zod';
import { prepareRuntime } from './runtime.mjs';

const graphOrigin = 'https://graph.microsoft.com';
const graphBase = '/v1.0';
const metadataFields = [
  'id', 'parentFolderId', 'internetMessageId', 'subject', 'from', 'sender',
  'toRecipients', 'ccRecipients', 'receivedDateTime', 'sentDateTime', 'isRead',
  'isDraft', 'flag', 'hasAttachments', 'bodyPreview'
].join(',');
const attachmentFields = 'id,name,contentType,size,isInline,lastModifiedDateTime';
const attachmentReadLimit = 10 * 1024 * 1024;
const draftAttachmentLimit = 3000000;
const hasControl = /[\u0000-\u001f\u007f]/;

function graphError(code, message, details = {}) {
  return Object.assign(new Error(message), { code, ...details });
}

function isLocalId(value) {
  if (!/^[A-Za-z0-9_-]+$/.test(value)) return false;
  try {
    const parsed = JSON.parse(Buffer.from(value, 'base64url'));
    return parsed?.backend === 'mac-cache' || parsed?.backend === 'local'
      || (typeof parsed?.v === 'number' && typeof parsed?.key === 'string')
      || (parsed?.v === 1 && typeof parsed?.entry === 'string' && typeof parsed?.store === 'string');
  } catch {
    return false;
  }
}

const graphId = z.string().min(1).max(4096).refine(value =>
  /^[A-Za-z0-9_+\/=-]+$/.test(value) && !isLocalId(value),
  'Expected a Graph ID. Resolve Mac cache mail by its exact internet_message_id.');
const limit = z.number().int().min(1).max(100).default(25);
const cursor = z.string().min(1).max(16000);
const filterDate = z.iso.datetime({ offset: true });
const page = { limit, cursor: cursor.optional() };
const message = { id: graphId };
const mutation = { request_id: z.string().min(1).max(128).regex(/^[A-Za-z0-9_.:-]+$/) };
const confirmed = { ...mutation, confirm: z.literal(true) };
const flags = z.enum(['notFlagged', 'flagged', 'complete']);
const recipients = z.array(z.email().max(254)).max(100);
const subject = z.string().max(998).refine(value => !hasControl.test(value), 'Subject cannot contain control characters');
const body = z.string().max(200000);
const draftFields = {
  subject: subject.optional(), body: body.optional(), to: recipients.optional(),
  cc: recipients.optional(), bcc: recipients.optional()
};

const actions = {
  status: z.strictObject({ verify: z.boolean().default(false) }),
  folders: z.strictObject({ parent: graphId.optional(), ...page }),
  list: z.strictObject({
    folder: graphId.optional(), after: filterDate.optional(), before: filterDate.optional(),
    unread: z.boolean().optional(), ...page
  }),
  drafts: z.strictObject(page),
  read: z.strictObject(message),
  mime: z.strictObject({
    ...message, max_bytes: z.number().int().min(1).max(attachmentReadLimit).default(attachmentReadLimit)
  }),
  attachments: z.strictObject({ ...message, ...page }),
  'attachment-read': z.strictObject({
    ...message, attachment_id: graphId,
    max_bytes: z.number().int().min(1).max(attachmentReadLimit).default(attachmentReadLimit)
  }),
  resolve: z.strictObject({
    internet_message_id: z.string().trim().min(1).max(998)
      .refine(value => !hasControl.test(value), 'Message ID cannot contain control characters'),
    folder: graphId.optional()
  }),
  'mark-read': z.strictObject({ ...message, ...mutation, is_read: z.boolean() }),
  flag: z.strictObject({ ...message, ...mutation, flag: flags }),
  move: z.strictObject({ ...message, ...confirmed, folder: graphId }),
  'batch-update': z.strictObject({
    ...mutation, ids: z.array(graphId).min(1).max(20), is_read: z.boolean().optional(),
    flag: flags.optional(), move_to: graphId.optional(), confirm: z.literal(true).optional()
  }).refine(params => params.is_read !== undefined || params.flag || params.move_to, 'Provide a read state, flag or destination')
    .refine(params => !params.move_to || params.confirm === true, 'Moving requires confirm=true')
    .refine(params => !params.move_to || !(params.is_read !== undefined || params.flag) || params.ids.length <= 10,
      'At most 10 IDs when updating and moving in the same batch'),
  compose: z.strictObject({ ...mutation, ...draftFields }),
  'draft-update': z.strictObject({ ...message, ...mutation, ...draftFields })
    .refine(params => Object.keys(draftFields).some(field => params[field] !== undefined), 'Provide at least one draft field'),
  'draft-attach': z.strictObject({
    ...message, ...confirmed,
    name: z.string().min(1).max(255).refine(value => !hasControl.test(value), 'Invalid attachment name'),
    content_type: z.string().max(255).regex(/^[A-Za-z0-9.+_-]+\/[A-Za-z0-9.+_-]+$/).optional(),
    content_base64: z.base64().max(draftAttachmentLimit * (4 / 3))
      .refine(value => Buffer.byteLength(value, 'base64') < draftAttachmentLimit, 'Graph draft file must be under 3 MB')
  }),
  'draft-discard': z.strictObject({ ...message, ...confirmed }),
  reply: z.strictObject({ ...message, ...mutation, comment: body.optional() }),
  'reply-all': z.strictObject({ ...message, ...mutation, comment: body.optional() }),
  forward: z.strictObject({ ...message, ...mutation, comment: body.optional(), to: recipients.min(1) }),
  'draft-send': z.strictObject({ ...message, ...confirmed })
};

export const GRAPH_ACTIONS = Object.freeze(Object.keys(actions));

function checkInput(action, input) {
  if (!Object.hasOwn(actions, action)) {
    throw graphError('INVALID_ARGUMENT', 'Unknown Graph action. Available: ' + GRAPH_ACTIONS.join(', '));
  }
  const result = actions[action].safeParse(input);
  if (!result.success) {
    throw graphError('INVALID_ARGUMENT', result.error.issues.map(issue => {
      const field = issue.path.join('.') || 'arguments';
      return field + ': ' + (issue.code === 'invalid_value' ? 'Unsupported value' : issue.message);
    }).join('; '));
  }
  const { after, before } = result.data;
  if (after && before && Date.parse(after) >= Date.parse(before)) {
    throw graphError('INVALID_ARGUMENT', 'after must be earlier than before');
  }
  return result.data;
}

function folderKey(value) {
  return encodeURIComponent(normalizedFolder(value));
}

function messagesResource(folder) {
  if (!folder) return '/me/messages';
  return '/me/mailFolders/' + folderKey(folder) + '/messages';
}

function messageResource(id, folder) {
  return messagesResource(folder) + '/' + encodeURIComponent(id);
}

function graphUrl(resource, query = {}) {
  const target = new URL(graphBase + resource, graphOrigin);
  for (const [key, value] of Object.entries(query)) {
    if (value !== undefined) target.searchParams.set(key, String(value));
  }
  return target;
}

function continuation(value, resource) {
  let target;
  try {
    target = new URL(value);
  } catch {
    throw graphError('GRAPH_INVALID_CURSOR', 'Graph pagination cursor is invalid');
  }
  if (target.origin !== graphOrigin || target.username || target.password || target.hash
      || target.pathname !== graphBase + resource) {
    throw graphError('GRAPH_INVALID_CURSOR', 'Graph pagination cursor does not match this mailbox resource');
  }
  return target;
}

function httpGraphError(status) {
  if (status === 401) return graphError('GRAPH_AUTH_REQUIRED', 'Graph rejected the delegated access token', { graphStatus: status });
  if (status === 403) return graphError('GRAPH_FORBIDDEN', 'Graph denied the operation. Check delegated permissions and mailbox policy', { graphStatus: status });
  if (status === 404) return graphError('GRAPH_NOT_FOUND', 'Graph could not find the mailbox item', { graphStatus: status });
  if (status === 429) return graphError('GRAPH_RATE_LIMITED', 'Graph rate-limited the operation', { graphStatus: status });
  return graphError('GRAPH_REQUEST_FAILED', 'Graph returned HTTP ' + status, { graphStatus: status });
}

function validToken(token) {
  return typeof token === 'string' && token.length > 0 && token.length <= 65536 && !/\s/.test(token);
}

async function readBinary(result, binaryLimit) {
  const length = result.headers.get('content-length');
  if (length !== null && Number(length) > binaryLimit) {
    await result.body?.cancel().catch(() => {});
    throw graphError('ATTACHMENT_TOO_LARGE', 'Graph file exceeds the configured byte limit');
  }
  const chunks = [];
  let read = 0;
  if (result.body) {
    for await (const chunk of result.body) {
      read += chunk.length;
      if (read > binaryLimit) throw graphError('ATTACHMENT_TOO_LARGE', 'Graph file exceeds the configured byte limit');
      chunks.push(chunk);
    }
  }
  return Buffer.concat(chunks, read);
}

function graphClient(token, fetcher) {
  if (!validToken(token)) {
    throw graphError('GRAPH_AUTH_REQUIRED', 'Set OUTLOOK_GRAPH_ACCESS_TOKEN to an approved Microsoft Graph delegated access token');
  }
  return async (method, target, data, binaryLimit) => {
    const headers = {
      Authorization: 'Bearer ' + token,
      Accept: binaryLimit === undefined ? 'application/json' : '*/*',
      Prefer: 'IdType="ImmutableId", outlook.body-content-type="text"'
    };
    if (data !== undefined) headers['Content-Type'] = 'application/json';
    let result;
    try {
      result = await fetcher(target, {
        method, headers, redirect: 'error', signal: AbortSignal.timeout(30000),
        ...(data === undefined ? {} : { body: JSON.stringify(data) })
      });
    } catch {
      throw graphError('GRAPH_TRANSPORT_ERROR', 'Graph connection failed or timed out');
    }
    if (!result.ok) throw httpGraphError(result.status);
    if (binaryLimit !== undefined) {
      try {
        return { status: result.status, bytes: await readBinary(result, binaryLimit) };
      } catch (error) {
        if (error.code === 'ATTACHMENT_TOO_LARGE') throw error;
        throw graphError('GRAPH_TRANSPORT_ERROR', 'Graph could not finish streaming the response');
      }
    }
    if (result.status === 202 || result.status === 204) return { status: result.status };
    try {
      const response = await result.json();
      if (!response || typeof response !== 'object' || Array.isArray(response)) throw new Error();
      return { status: result.status, response };
    } catch {
      throw graphError('GRAPH_BAD_RESPONSE', 'Graph returned an invalid JSON response');
    }
  };
}

async function getPage(request, resource, cursorValue, query) {
  const { response } = await request('GET', cursorValue ? continuation(cursorValue, resource) : graphUrl(resource, query));
  if (!Array.isArray(response?.value)) throw graphError('GRAPH_BAD_RESPONSE', 'Graph response has no item list');
  const next = response['@odata.nextLink'];
  return { backend: 'graph', items: response.value,
    next_cursor: next === undefined ? null : continuation(next, resource).href };
}

async function resolveMessage(request, params) {
  const resource = messagesResource(params.folder);
  const escaped = params.internet_message_id.replaceAll("'", "''");
  const query = { $filter: "internetMessageId eq '" + escaped + "'", $top: 2, $select: metadataFields };
  const matches = new Map();
  const visited = new Set();
  let next;
  do {
    const pageResult = await getPage(request, resource, next, query);
    for (const item of pageResult.items) {
      if (item?.internetMessageId !== params.internet_message_id) continue;
      if (!graphId.safeParse(item.id).success) throw graphError('GRAPH_BAD_RESPONSE', 'Graph message has no valid ID');
      matches.set(item.id, item);
      if (matches.size > 1) {
        throw graphError('GRAPH_ID_AMBIGUOUS', 'More than one Graph message has this exact Internet Message-ID. Supply a folder or choose an explicit Graph ID.');
      }
    }
    next = pageResult.next_cursor;
    if (next && (visited.has(next) || visited.size >= 99)) {
      throw graphError('GRAPH_RESOLVE_INCOMPLETE', 'Could not establish a unique Graph ID within the pagination limit');
    }
    if (next) visited.add(next);
  } while (next);
  if (matches.size === 0) throw graphError('GRAPH_NOT_FOUND', 'No Graph message has this exact Internet Message-ID');
  const [result] = matches.values();
  return { backend: 'graph', id: result.id, message: result };
}

function digest(data) {
  return createHash('sha256').update(data).digest('hex');
}

function graphStateDirectory() {
  if (process.env.OUTLOOK_GRAPH_STATE_DIR !== undefined) return process.env.OUTLOOK_GRAPH_STATE_DIR;
  if (process.platform === 'darwin') return path.join(os.homedir(), 'Library', 'Application Support', 'outlook-cli', 'graph');
  if (process.platform === 'win32') return path.join(process.env.LOCALAPPDATA ?? path.join(os.homedir(), 'AppData', 'Local'), 'outlook-cli', 'graph');
  return path.join(os.homedir(), '.local', 'state', 'outlook-cli', 'graph');
}

function journalError(code, requestId) {
  const message = code === 'GRAPH_REQUEST_CONFLICT'
    ? 'request_id has already been used for different arguments'
    : 'A previous Graph mutation may have been accepted. Verify the mailbox before attempting a new action';
  return graphError(code, message + ' (' + requestId + ')');
}

async function readJournal(file) {
  if (process.platform === 'win32' && (await lstat(file)).isSymbolicLink()) throw new Error();
  const handle = await open(file, constants.O_RDONLY | (constants.O_NOFOLLOW ?? 0) | (constants.O_NONBLOCK ?? 0));
  try {
    const info = await handle.stat();
    if (!info.isFile() || info.nlink !== 1 || info.size > 65536
        || (process.platform !== 'win32' && (info.uid !== process.getuid() || (info.mode & 0o077) !== 0))) {
      throw new Error();
    }
    const bytes = Buffer.alloc(info.size);
    const { bytesRead } = await handle.read(bytes, 0, bytes.length, 0);
    if (bytesRead !== info.size) throw new Error();
    return JSON.parse(bytes);
  } finally {
    await handle.close();
  }
}

async function writeJournalFile(file, data) {
  const handle = await open(file, 'wx', 0o600);
  try {
    await handle.writeFile(JSON.stringify(data));
    await handle.sync();
  } finally {
    await handle.close();
  }
}

async function syncDirectory(directory) {
  if (process.platform === 'win32') return;
  const handle = await open(directory, 'r');
  try {
    await handle.sync();
  } finally {
    await handle.close();
  }
}

async function persistJournal(file, record) {
  const temporary = file + '.' + randomBytes(12).toString('hex');
  try {
    await writeJournalFile(temporary, record);
    await rename(temporary, file);
    await syncDirectory(path.dirname(file));
  } finally {
    await unlink(temporary).catch(() => {});
  }
}

function definitelyRejected(error) {
  return (error.graphStatus >= 400 && error.graphStatus < 500
      && ![408, 409, 423, 429].includes(error.graphStatus))
    || error.code === 'GRAPH_NOT_DRAFT';
}

function fileJournal(stateDirectory) {
  return {
    async run(key, fingerprint, requestId, operation, { releaseOnRejection = false } = {}) {
      await prepareRuntime(stateDirectory);
      const directory = await realpath(stateDirectory);
      if (directory !== path.resolve(stateDirectory)) await prepareRuntime(directory);
      const file = path.join(directory, 'request-' + digest(key) + '.json');
      try {
        await writeJournalFile(file, { fingerprint, state: 'pending' });
        await syncDirectory(directory);
      } catch (error) {
        if (error.code !== 'EEXIST') {
          throw graphError('GRAPH_JOURNAL_FAILED', 'Could not reserve the Graph mutation request in the private journal');
        }
        let previous;
        try {
          previous = await readJournal(file);
        } catch {
          throw journalError('GRAPH_REQUEST_IN_DOUBT', requestId);
        }
        if (previous?.fingerprint !== fingerprint) throw journalError('GRAPH_REQUEST_CONFLICT', requestId);
        if (previous.state === 'complete') return { ...previous.result, replayed: true };
        if (previous.state === 'failed') throw graphError(previous.code, previous.message);
        throw journalError('GRAPH_REQUEST_IN_DOUBT', requestId);
      }
      let result;
      try {
        result = await operation();
      } catch (error) {
        if (!definitelyRejected(error)) throw journalError('GRAPH_REQUEST_IN_DOUBT', requestId);
        try {
          if (releaseOnRejection) {
            await unlink(file);
            await syncDirectory(directory);
          } else {
            await persistJournal(file, { fingerprint, state: 'failed', code: error.code, message: error.message });
          }
        } catch {
          throw journalError('GRAPH_REQUEST_IN_DOUBT', requestId);
        }
        throw error;
      }
      try {
        await persistJournal(file, { fingerprint, state: 'complete', result });
      } catch {
        throw journalError('GRAPH_REQUEST_IN_DOUBT', requestId);
      }
      return result;
    }
  };
}

function graphRecipients(addresses) {
  return addresses.map(address => ({ emailAddress: { address } }));
}

function draftBody(params) {
  const result = {};
  if (params.subject !== undefined) result.subject = params.subject;
  if (params.body !== undefined) result.body = { contentType: 'Text', content: params.body };
  for (const [key, property] of [['to', 'toRecipients'], ['cc', 'ccRecipients'], ['bcc', 'bccRecipients']]) {
    if (params[key] !== undefined) result[property] = graphRecipients(params[key]);
  }
  return result;
}

function updateBody(params) {
  const update = {};
  if (params.is_read !== undefined) update.isRead = params.is_read;
  if (params.flag !== undefined) update.flag = { flagStatus: params.flag };
  return update;
}

function receipt(result, originalId) {
  const id = result.response?.id ?? originalId;
  if (!graphId.safeParse(id).success) throw graphError('GRAPH_BAD_RESPONSE', 'Graph mutation returned no valid item ID');
  return { status: result.status, id };
}

async function requireDraft(request, id) {
  const { response } = await request('GET', graphUrl(messageResource(id, 'drafts'), { $select: 'id,isDraft' }));
  if (response?.isDraft !== true) throw graphError('GRAPH_NOT_DRAFT', 'Graph item is not a draft');
}

async function batchUpdate(request, params) {
  const requests = [];
  const updates = updateBody(params);
  const hasUpdate = Object.keys(updates).length > 0;
  const expected = new Map();
  for (const id of new Set(params.ids)) {
    const resource = messageResource(id);
    let dependency;
    if (hasUpdate) {
      dependency = String(requests.length + 1);
      expected.set(dependency, { id, action: 'update' });
      requests.push({ id: dependency, method: 'PATCH', url: resource, body: updates,
        headers: { 'Content-Type': 'application/json', Prefer: 'IdType="ImmutableId"' } });
    }
    if (params.move_to) {
      const step = String(requests.length + 1);
      expected.set(step, { id, action: 'move' });
      requests.push({ id: step, method: 'POST', url: resource + '/move',
        headers: { 'Content-Type': 'application/json', Prefer: 'IdType="ImmutableId"' },
        body: { destinationId: normalizedFolder(params.move_to) },
        ...(dependency ? { dependsOn: [dependency] } : {}) });
    }
  }
  const { response } = await request('POST', graphUrl('/$batch'), { requests });
  if (!Array.isArray(response?.responses)) throw graphError('GRAPH_BAD_RESPONSE', 'Graph returned no batch results');
  const responses = new Map();
  for (const result of response.responses) {
    if (!result || !expected.has(result.id) || responses.has(result.id) || !Number.isInteger(result.status)) {
      throw graphError('GRAPH_BAD_RESPONSE', 'Graph returned inconsistent batch results');
    }
    responses.set(result.id, result);
  }
  const items = [];
  for (const [step, item] of expected) {
    const result = responses.get(step);
    const status = result?.status ?? null;
    const ok = status >= 200 && status < 300;
    items.push({ ...item, ok, status,
      ...(ok && typeof result.body?.id === 'string' ? { result_id: result.body.id } : {}),
      ...(ok ? {} : { error: status === null ? 'GRAPH_OUTCOME_UNKNOWN' : httpGraphError(status).code }) });
  }
  return { items, complete: items.every(item => item.ok) };
}

function normalizedFolder(value) {
  const aliases = { sent: 'sentitems', deleted: 'deleteditems', junk: 'junkemail' };
  const alias = value.toLowerCase();
  return Object.hasOwn(aliases, alias) ? aliases[alias] : value;
}

async function mutate(request, action, params) {
  const resource = params.id && messageResource(params.id);
  const draft = params.id && messageResource(params.id, 'drafts');
  switch (action) {
    case 'mark-read':
    case 'flag':
      return receipt(await request('PATCH', graphUrl(resource), updateBody(params)), params.id);
    case 'move':
      return receipt(await request('POST', graphUrl(resource + '/move'), {
        destinationId: normalizedFolder(params.folder)
      }), params.id);
    case 'batch-update':
      return batchUpdate(request, params);
    case 'compose':
      return receipt(await request('POST', graphUrl('/me/messages'), draftBody(params)));
    case 'draft-update':
      await requireDraft(request, params.id);
      return receipt(await request('PATCH', graphUrl(draft), draftBody(params)), params.id);
    case 'draft-attach':
      await requireDraft(request, params.id);
      return receipt(await request('POST', graphUrl(draft + '/attachments'), {
        '@odata.type': '#microsoft.graph.fileAttachment',
        name: params.name,
        contentBytes: params.content_base64,
        ...(params.content_type ? { contentType: params.content_type } : {})
      }));
    case 'draft-discard':
      await requireDraft(request, params.id);
      return receipt(await request('DELETE', graphUrl(draft)), params.id);
    case 'reply':
    case 'reply-all':
    case 'forward': {
      const operation = { reply: 'createReply', 'reply-all': 'createReplyAll', forward: 'createForward' }[action];
      const input = {};
      if (params.comment !== undefined) input.comment = params.comment;
      if (action === 'forward') input.toRecipients = graphRecipients(params.to);
      return receipt(await request('POST', graphUrl(resource + '/' + operation), input));
    }
    case 'draft-send': {
      await requireDraft(request, params.id);
      const result = await request('POST', graphUrl(resource + '/send'));
      if (result.status !== 202) throw graphError('GRAPH_BAD_RESPONSE', 'Graph did not acknowledge draft submission');
      return receipt(result, params.id);
    }
    default:
      throw graphError('INVALID_ARGUMENT', 'Unknown Graph mutation');
  }
}

async function readWorkflow(request, action, params) {
  switch (action) {
    case 'status': {
      await getPage(request, '/me/mailFolders', undefined, { $top: 1, $select: 'id' });
      return { backend: 'graph', configured: true, verified: true };
    }
    case 'folders': {
      const resource = params.parent ? '/me/mailFolders/' + folderKey(params.parent) + '/childFolders' : '/me/mailFolders';
      return getPage(request, resource, params.cursor, {
        $top: params.limit, $select: 'id,displayName,parentFolderId,childFolderCount,unreadItemCount,totalItemCount'
      });
    }
    case 'list':
    case 'drafts': {
      const resource = messagesResource(action === 'drafts' ? 'drafts' : params.folder);
      const filters = [];
      if (params.after) filters.push('receivedDateTime ge ' + new Date(params.after).toISOString());
      if (params.before) filters.push('receivedDateTime lt ' + new Date(params.before).toISOString());
      if (params.unread !== undefined) filters.push('isRead eq ' + !params.unread);
      return getPage(request, resource, params.cursor, {
        $top: params.limit, $select: metadataFields, $filter: filters.length ? filters.join(' and ') : undefined,
        $orderby: params.unread === undefined || params.after || params.before ? 'receivedDateTime desc' : undefined
      });
    }
    case 'read': {
      const { response } = await request('GET', graphUrl(messageResource(params.id), { $select: metadataFields + ',body' }));
      return { backend: 'graph', message: response };
    }
    case 'mime': {
      let bytes;
      try {
        ({ bytes } = await request('GET', graphUrl(messageResource(params.id) + '/$value'), undefined, params.max_bytes));
      } catch (error) {
        if (error.code === 'ATTACHMENT_TOO_LARGE') {
          throw graphError('MIME_TOO_LARGE', 'Graph MIME exceeds the configured byte limit');
        }
        throw error;
      }
      if (!bytes.length) throw graphError('GRAPH_BAD_RESPONSE', 'Graph returned empty MIME');
      return { backend: 'graph', id: params.id, format: 'rfc822',
        source: 'graph_mime_representation', size_bytes: bytes.length,
        data: bytes.toString('base64'), sha256: digest(bytes) };
    }
    case 'attachments':
      return getPage(request, messageResource(params.id) + '/attachments', params.cursor, {
        $top: params.limit, $select: attachmentFields
      });
    case 'attachment-read': {
      const resource = messageResource(params.id) + '/attachments/' + encodeURIComponent(params.attachment_id);
      const { response: attachment } = await request('GET', graphUrl(resource, { $select: attachmentFields }));
      if (String(attachment?.['@odata.type']).replace(/^#/, '') !== 'microsoft.graph.fileAttachment') {
        throw graphError('ATTACHMENT_UNSUPPORTED', 'Graph attachment is not a by-value file');
      }
      if (!Number.isInteger(attachment.size) || attachment.size < 0) {
        throw graphError('GRAPH_BAD_RESPONSE', 'Graph attachment size is unavailable');
      }
      if (attachment.size > params.max_bytes) {
        throw graphError('ATTACHMENT_TOO_LARGE', 'Graph file exceeds the configured byte limit');
      }
      const { bytes } = await request('GET', graphUrl(resource + '/$value'), undefined, params.max_bytes);
      if (bytes.length !== attachment.size) {
        throw graphError('GRAPH_ATTACHMENT_CHANGED', 'Graph attachment size changed during retrieval');
      }
      return { backend: 'graph', id: params.id, attachment_id: params.attachment_id,
        filename: attachment.name, content_type: attachment.contentType, size_bytes: bytes.length,
        data: bytes.toString('base64'), sha256: digest(bytes) };
    }
    case 'resolve':
      return resolveMessage(request, params);
    default:
      throw graphError('INVALID_ARGUMENT', 'Unknown Graph operation');
  }
}

export async function runGraphWorkflow(action, input = {}, options = {}) {
  const params = checkInput(action, input);
  const token = options.token ?? process.env.OUTLOOK_GRAPH_ACCESS_TOKEN;
  if (action === 'status' && !params.verify) {
    return { backend: 'graph', configured: validToken(token), verified: false };
  }
  const request = graphClient(token, options.fetch ?? globalThis.fetch);
  const { request_id: requestId, ...argumentsWithoutId } = params;
  if (requestId === undefined) return readWorkflow(request, action, params);
  const journal = options.journal ?? fileJournal(options.stateDirectory ?? graphStateDirectory());
  const { response: mailbox } = await request('GET', graphUrl('/me/mailFolders/msgfolderroot', { $select: 'id' }));
  if (!graphId.safeParse(mailbox?.id).success) throw graphError('GRAPH_BAD_RESPONSE', 'Could not identify the signed-in Graph mailbox');
  const mailboxKey = digest(mailbox.id) + ':';
  const fingerprint = digest(JSON.stringify([action, argumentsWithoutId]));
  const operation = () => mutate(request, action, params);
  const result = await journal.run(mailboxKey + 'request:' + requestId, fingerprint, requestId, action === 'draft-send'
    ? () => journal.run(mailboxKey + 'send:' + params.id, fingerprint, requestId, operation, { releaseOnRejection: true })
    : operation);
  return { backend: 'graph', action, request_id: requestId, ...result };
}
