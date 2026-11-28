import { spawn } from 'node:child_process';
import { createHash } from 'node:crypto';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { runtimeDir } from './runtime.mjs';
import { buildSwiftHelper } from './swift-helper.mjs';
const source = fileURLToPath(new URL('../macos/calendar.swift', import.meta.url));
const bundleIdentifier = 'io.github.guanglinhuang99.outlook-cli.calendar';
const infoPlist = `<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0"><dict>
<key>CFBundleIdentifier</key><string>${bundleIdentifier}</string>
<key>CFBundleName</key><string>Outlook CLI Calendar</string>
<key>NSCalendarsFullAccessUsageDescription</key><string>Outlook CLI reads calendars and agendas you can access on this Mac.</string>
</dict></plist>`;
const authorizations = new Set(['not_determined', 'denied', 'restricted', 'full_access', 'write_only', 'unknown']);
const defaults = { source: 'eventkit', data_scope: 'local_calendar_store' };
const dayMs = 86400000;

function calendarError(code, message) {
  return Object.assign(new Error(message), { code });
}

function invalid(message) {
  return calendarError('INVALID_ARGUMENT', message);
}

function digest(value) {
  return createHash('sha256').update(JSON.stringify(value)).digest('hex');
}

function buildHelper(stateDir) {
  return buildSwiftHelper({
    stateDir, source, name: 'calendar', files: { 'Info.plist': infoPlist }, sign: bundleIdentifier,
    flags: build => ['-parse-as-library', '-Xlinker', '-sectcreate', '-Xlinker', '__TEXT', '-Xlinker', '__info_plist',
      '-Xlinker', path.join(build, 'Info.plist')],
    fail: reason => calendarError('CALENDAR_UNAVAILABLE', `EventKit ${reason}`)
  });
}

async function readEventKit(request, stateDir) {
  if (process.platform !== 'darwin') {
    throw calendarError('CALENDAR_UNAVAILABLE', 'Calendar commands require macOS EventKit');
  }
  const helper = await buildHelper(stateDir);
  return new Promise((resolve, reject) => {
    const child = spawn(helper, [], { stdio: ['pipe', 'pipe', 'ignore'] });
    const chunks = [];
    let size = 0;
    let failure;
    const timeout = setTimeout(() => {
      failure ??= calendarError('CALENDAR_TIMEOUT', 'EventKit did not return before the timeout');
      child.kill('SIGKILL');
    }, request.action === 'authorize' ? 90000 : 25000);
    child.once('error', () => {
      failure ??= calendarError('CALENDAR_UNAVAILABLE', 'Could not start the local EventKit helper');
    });
    child.stdin.on('error', () => {});
    child.stdout.on('data', chunk => {
      size += chunk.length;
      if (size > 16 * 1024 * 1024) {
        failure ??= calendarError('CALENDAR_QUERY_TOO_LARGE', 'EventKit result exceeds 16 MiB. Narrow the date or calendar scope');
        child.kill('SIGKILL');
      } else {
        chunks.push(chunk);
      }
    });
    child.once('close', code => {
      clearTimeout(timeout);
      if (failure) { reject(failure); return; }
      if (code !== 0) {
        reject(calendarError('CALENDAR_UNAVAILABLE', 'EventKit did not complete in the current macOS host session'));
        return;
      }
      try { resolve(JSON.parse(Buffer.concat(chunks).toString('utf8'))); }
      catch { reject(calendarError('CALENDAR_UNAVAILABLE', 'Invalid response from the local EventKit helper')); }
    });
    child.stdin.end(JSON.stringify(request));
  });
}

function dayValue(value, field) {
  if (typeof value === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(value)) {
    const ms = Date.parse(`${value}T00:00:00.000Z`);
    if (Number.isFinite(ms) && new Date(ms).toISOString().slice(0, 10) === value
        && value >= '1970-01-01' && value <= '9999-12-31') return ms;
  }
  throw invalid(`${field} must be a valid YYYY-MM-DD date between 1970-01-01 and 9999-12-31`);
}

function isoDay(ms) {
  return new Date(ms).toISOString().slice(0, 10);
}

function agendaScope(params, savedCursor, now) {
  const timeZone = params.time_zone === undefined ? 'Asia/Singapore' : params.time_zone;
  if (typeof timeZone !== 'string' || timeZone.length > 128 || !/^[A-Za-z][A-Za-z0-9_+/-]*$/.test(timeZone)) {
    throw invalid('time_zone must be a named IANA time zone');
  }
  let clock;
  try {
    clock = new Intl.DateTimeFormat('en-CA', { year: 'numeric', month: '2-digit', day: '2-digit', timeZone });
  } catch {
    throw invalid('time_zone must be a named IANA time zone');
  }
  const usingRange = params.from_date !== undefined || params.to_date !== undefined;
  const dateChoices = Number(params.date !== undefined) + Number(usingRange) + Number(params.days !== undefined);
  if (dateChoices > 1) throw invalid('Choose date, from_date with to_date, or days');
  let start;
  let end;
  if (params.date !== undefined) {
    start = end = dayValue(params.date, 'date');
  } else if (usingRange) {
    start = dayValue(params.from_date, 'from_date');
    end = dayValue(params.to_date, 'to_date');
  } else {
    const days = params.days === undefined ? 1 : boundedInt(params.days, 'days', 366);
    const parts = Object.fromEntries(clock.formatToParts(now).map(({ type, value }) => [type, value]));
    const anchor = savedCursor?.anchor ?? `${parts.year}-${parts.month}-${parts.day}`;
    start = dayValue(anchor, 'date');
    end = start + (days - 1) * dayMs;
  }
  if (end < start || (end - start) / dayMs >= 366 || end >= Date.UTC(9999, 11, 31)) {
    throw invalid('Agenda requires a range of 1 to 366 days, ending by 9999-12-30');
  }
  const calendarIDs = params.calendar_ids === undefined ? [] : params.calendar_ids;
  if (!Array.isArray(calendarIDs) || calendarIDs.length > 50
      || calendarIDs.some(id => typeof id !== 'string' || id.trim() !== id || !id || id.length > 512)
      || new Set(calendarIDs).size !== calendarIDs.length) {
    throw invalid('calendar_ids must contain up to 50 different exact calendar IDs');
  }
  const query = params.query === undefined ? null : params.query;
  if (params.query !== undefined && (typeof query !== 'string' || !query.trim() || query.length > 256)) {
    throw invalid('query must be nonempty text up to 256 characters');
  }
  for (const key of ['include_details', 'include_canceled']) {
    if (params[key] !== undefined && typeof params[key] !== 'boolean') throw invalid(`${key} must be a boolean`);
  }
  return {
    from_date: isoDay(start), end_date_exclusive: isoDay(end + dayMs), time_zone: timeZone,
    calendar_ids: [...calendarIDs].sort(), query: query?.trim() ?? null,
    include_details: params.include_details === true, include_canceled: params.include_canceled === true,
  };
}

function boundedInt(value, name, maximum) {
  if (!Number.isSafeInteger(value) || value < 1 || value > maximum) {
    throw invalid(`${name} must be an integer from 1 to ${maximum}`);
  }
  return value;
}

function parseCursor(value) {
  if (value === undefined) return null;
  if (typeof value === 'string' && value.length <= 512 && /^[A-Za-z0-9_-]+$/.test(value)) {
    try {
      const cursor = JSON.parse(Buffer.from(value, 'base64url').toString('utf8'));
      if (cursor.v === 1 && /^[a-f0-9]{64}$/.test(cursor.scope) && /^[a-f0-9]{64}$/.test(cursor.rows)
          && Number.isSafeInteger(cursor.offset) && cursor.offset > 0 && cursor.offset < 4096
          && (cursor.anchor === null || (typeof cursor.anchor === 'string' && dayValue(cursor.anchor, 'cursor date')))) {
        return cursor;
      }
    } catch {}
  }
  throw invalid('Invalid calendar cursor');
}

function sortItems(action, items) {
  const keys = action === 'calendars' ? ['source_title', 'title', 'id'] : [
    'start_at', 'end_at', 'calendar_id', 'id', 'occurrence_at', 'title',
  ];
  return items.sort((left, right) => {
    for (const key of keys) {
      const a = left[key] ?? '';
      const b = right[key] ?? '';
      if (a < b) return -1;
      if (a > b) return 1;
    }
    return 0;
  });
}

export async function runCalendarWorkflow(action, params = {}, {
  stateDir = `${runtimeDir}-calendar`, now = new Date(), queryNative,
} = {}) {
  const paginationKeys = ['limit', 'cursor'];
  const allowed = {
    status: [],
    authorize: [],
    calendars: paginationKeys,
    agenda: [...paginationKeys, 'date', 'from_date', 'to_date', 'days', 'time_zone', 'calendar_ids',
      'query', 'include_details', 'include_canceled'],
  };
  if (!Object.hasOwn(allowed, action)) throw invalid('Unknown calendar action');
  if (!params || typeof params !== 'object' || Array.isArray(params)
      || Object.keys(params).some(key => !allowed[action].includes(key))) throw invalid('Unexpected calendar arguments');
  const invoke = queryNative ?? (request => readEventKit(request, stateDir));
  async function query(request) {
    const result = await invoke(request);
    if (result?.error) throw calendarError(result.error.code, result.error.message);
    return result;
  }
  if (action === 'status' || action === 'authorize') {
    const result = await query({ action });
    if (!authorizations.has(result?.authorization)) {
      throw calendarError('CALENDAR_UNAVAILABLE', 'EventKit authorization status is unavailable');
    }
    const authorization = result.authorization;
    if (action === 'authorize' && authorization !== 'full_access') {
      throw calendarError('CALENDAR_PERMISSION_DENIED', 'Full Calendar access was not granted');
    }
    return { ...defaults, authorization, can_read: authorization === 'full_access' };
  }
  const limit = params.limit === undefined ? 50 : boundedInt(params.limit, 'limit', 100);
  const cursor = parseCursor(params.cursor);
  const scope = action === 'agenda' ? agendaScope(params, cursor, now) : {};
  const scopeID = digest([action, scope]);
  if (cursor && cursor.scope !== scopeID) throw invalid('Calendar cursor does not match these filters');
  const result = await query({ action, ...scope });
  if (!Array.isArray(result?.items) || result.items.length > 4096
      || result.items.some(item => !item || typeof item !== 'object' || typeof item.id !== 'string')) {
    throw calendarError('CALENDAR_UNAVAILABLE', 'Invalid items from the local EventKit helper');
  }
  const items = sortItems(action, [...result.items]);
  const rows = digest(items);
  if (cursor && (cursor.rows !== rows || cursor.offset >= items.length)) {
    throw calendarError('CALENDAR_CURSOR_STALE', 'The local calendar changed. Start again from the first page');
  }
  const offset = cursor?.offset ?? 0;
  const next = offset + limit;
  const complete = next >= items.length;
  const nextCursor = complete ? null : Buffer.from(JSON.stringify({
    v: 1, scope: scopeID, rows, offset: next, anchor: scope.from_date ?? null,
  })).toString('base64url');
  return {
    ...defaults, ...scope, authorization: 'full_access', items: items.slice(offset, next),
    total_count: items.length, complete, next_cursor: nextCursor,
  };
}
