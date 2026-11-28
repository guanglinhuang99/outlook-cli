// Default agent-facing view of mailbox results. Agents pay for every returned character,
// so list, read and bundle results keep only what is needed to choose and read mail:
// diagnostic evidence, duplicated identifiers, constant descriptions, false flags and zero
// counts are left out. Any abnormal state stays visible. detail=full returns everything.
export const compactMethods = new Set([
  'messages', 'today', 'search', 'read', 'context_bundle', 'today_bundle', 'messages_bundle'
]);

const normalStatus = {
  received_time_status: ['verified'],
  folder_status: ['verified_latest_observation'],
  subject_status: ['paired', 'verified'],
  name_status: ['verified']
};
const keptFlags = ['body_search_incomplete', 'person_search_incomplete', 'unread_filter_incomplete',
  'date_filter_incomplete', 'retry_required', 'folders_truncated', 'attachments_truncated'];
const keptCounts = ['received_time_unknown', 'unread_unknown_count'];
const passthrough = ['backend', 'data_scope', 'fallback_reason', 'fallback_message', 'outlook_connection',
  'next_offset', 'next_cursor', 'complete', 'mode', 'folder', 'date', 'errors', 'warnings', 'body_complete'];

function present(value) {
  return value !== undefined && value !== null && value !== '' && !(Array.isArray(value) && value.length === 0);
}

function compactBody(body) {
  if (!body || typeof body !== 'object') return body;
  const result = { text: body.text, total_chars: body.total_chars };
  if (body.offset) result.offset = body.offset;
  if (body.next_offset !== null && body.next_offset !== undefined) result.next_offset = body.next_offset;
  if (body.truncated) result.truncated = true;
  if (body.available === false || body.text === null || body.text === undefined) {
    for (const key of ['available', 'status', 'diagnosis']) if (present(body[key])) result[key] = body[key];
  }
  if (body.format && body.format !== 'text') result.format = body.format;
  return result;
}

export function compactMessage(message) {
  if (!message || typeof message !== 'object') return message;
  const result = {};
  const take = (key, value = message[key]) => { if (present(value)) result[key] = value; };
  take('id');
  take('subject');
  if (message.sender) {
    const sender = {};
    if (present(message.sender.name)) sender.name = message.sender.name;
    const address = message.sender.smtp_address ?? message.sender.address;
    if (present(address)) sender.address = address;
    if (Object.keys(sender).length) result.sender = sender;
  }
  take('received_at');
  if (message.unread !== undefined) result.unread = message.unread;
  take('folder', message.folder?.path ?? message.folder_path ?? message.folder?.name);
  if (message.attachment_count) result.attachment_count = message.attachment_count;
  for (const key of ['preview', 'body_snippet', 'matched_person_roles', 'matched_fields', 'to', 'cc']) take(key);
  if (Array.isArray(message.attachments) && message.attachments.length) {
    result.attachments = message.attachments.map(item => {
      const entry = { index: item.index, filename: item.filename, size_bytes: item.size_bytes };
      if (item.status && item.status !== 'cached_at_snapshot') entry.status = item.status;
      if (item.type !== undefined && item.type !== null && item.type !== 1) entry.type = item.type;
      return entry;
    });
  }
  if (message.body !== undefined) result.body = compactBody(message.body);
  for (const [key, normal] of Object.entries(normalStatus)) {
    if (present(message[key]) && !normal.includes(message[key])) result[key] = message[key];
  }
  if (message.download_state === 0) result.header_only = true;
  return result;
}

export function compactResult(method, data) {
  if (!data || typeof data !== 'object' || Array.isArray(data)) return data;
  const result = method === 'read' ? compactMessage(data) : {};
  for (const key of passthrough) {
    if (key in data && (present(data[key]) || key === 'next_offset' || key === 'next_cursor' || key === 'complete')) {
      if ((key === 'errors' || key === 'warnings') && !present(data[key])) continue;
      if (key === 'body_complete' && data[key] !== false) continue;
      if (key === 'folder' && method === 'read') continue;
      result[key] = data[key];
    }
  }
  for (const key of keptFlags) if (data[key] === true) result[key] = true;
  for (const key of keptCounts) if (Number.isFinite(data[key]) && data[key] > 0) result[key] = data[key];
  const fulltext = data.fulltext;
  if (fulltext && (fulltext.complete === false || fulltext.pending > 0 || data.retry_required || data.body_search_incomplete)) {
    result.fulltext = fulltext;
  }
  for (const key of ['messages', 'items']) {
    if (Array.isArray(data[key])) result[key] = data[key].map(compactMessage);
  }
  return result;
}
