import { OutlookBridge } from './bridge.mjs';
import { MacCacheBridge } from './mac.mjs';
import { LocalFileBridge, decodeId, localError } from './offline.mjs';

const modes = ['auto', 'com', 'local', 'mac-cache'];
const singaporeTimezone = 'Asia/Singapore';
const minimumToday = '1970-01-01';
const maximumToday = '9999-12-30';
const singaporeOffsetFormatter = new Intl.DateTimeFormat('en', {
  timeZone: singaporeTimezone,
  timeZoneName: 'longOffset'
});

function invalidDate(message) {
  throw localError('INVALID_ARGUMENT', message);
}

function parseDay(value, allowNextDay = false) {
  if (typeof value !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(value)) invalidDate('date must be YYYY-MM-DD');
  const maximum = allowNextDay ? '9999-12-31' : maximumToday;
  if (value < minimumToday || value > maximum) {
    invalidDate(`date must be between ${minimumToday} and ${maximumToday}`);
  }
  const [year, month, day] = value.split('-').map(Number);
  const date = new Date(0);
  date.setUTCFullYear(year, month - 1, day);
  date.setUTCHours(0, 0, 0, 0);
  if (date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) invalidDate('date must be a valid calendar date');
  return date;
}

function formatDay(date) {
  return `${date.getUTCFullYear().toString().padStart(4, '0')}-${String(date.getUTCMonth() + 1).padStart(2, '0')}-${String(date.getUTCDate()).padStart(2, '0')}`;
}

function singaporeOffset(value) {
  const zone = singaporeOffsetFormatter.formatToParts(new Date(value)).find(part => part.type === 'timeZoneName')?.value;
  if (zone === 'GMT') return 0;
  const match = /^GMT([+-])(\d{1,2})(?::(\d{2}))?$/.exec(zone ?? '');
  if (!match) invalidDate('Cannot determine Singapore offset');
  const minutes = Number(match[2]) * 60 + Number(match[3] ?? 0);
  return (match[1] === '-' ? -1 : 1) * minutes;
}

function formatOffset(minutes) {
  const sign = minutes < 0 ? '-' : '+';
  const absolute = Math.abs(minutes);
  return `${sign}${String(Math.floor(absolute / 60)).padStart(2, '0')}:${String(absolute % 60).padStart(2, '0')}`;
}

function singaporeMidnight(value, allowNextDay = false) {
  const start = parseDay(value, allowNextDay);
  const utcMidnight = start.getTime();
  let offset = singaporeOffset(utcMidnight);
  for (let attempt = 0; attempt < 3; attempt += 1) {
    const instant = utcMidnight - offset * 60000;
    const corrected = singaporeOffset(instant);
    if (corrected === offset) return `${value}T00:00:00${formatOffset(offset)}`;
    offset = corrected;
  }
  invalidDate('Cannot determine Singapore local midnight');
}

function singaporeDay(value) {
  const parts = Object.fromEntries(new Intl.DateTimeFormat('en', {
    timeZone: singaporeTimezone,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit'
  }).formatToParts(new Date(value)).filter(part => part.type !== 'literal').map(part => [part.type, part.value]));
  if (!parts.year || !parts.month || !parts.day) invalidDate('Cannot determine Singapore date');
  return `${parts.year}-${parts.month}-${parts.day}`;
}

function withWindowsDateBasis(result) {
  return { ...result, date_basis: 'received_at' };
}

export class BackendRouter {
  constructor({ com, local, mac, mode = process.env.OUTLOOK_LOCAL_BACKEND ?? 'auto', platform = process.platform, now = () => Date.now() } = {}) {
    if (!modes.includes(mode)) throw new Error('OUTLOOK_LOCAL_BACKEND must be auto, com, local or mac-cache');
    this.mode = mode;
    this.platform = platform;
    this.now = now;
    if (platform === 'darwin') {
      this.mac = mac ?? new MacCacheBridge();
      this.com = com;
      this.local = local;
    } else if (platform === 'win32') {
      this.com = com ?? new OutlookBridge();
      this.local = local ?? new LocalFileBridge();
      this.mac = mac;
    } else {
      this.com = com;
      this.local = local;
      this.mac = mac;
    }
  }
  get bridges() { return [...new Set([this.com, this.local, this.mac].filter(Boolean))]; }
  get queue() { return this.bridges.flatMap(bridge => bridge.queue ?? []); }
  get active() { return this.bridges.find(bridge => bridge.active) ?? false; }
  async call(method, params = {}) {
    if (method === 'today') return this.today(params);
    const { backend, ...args } = params;
    const mode = backend ?? this.mode;
    const id = decodeId(args.id ?? args.folder);
    if (!modes.includes(mode)) throw localError('INVALID_ARGUMENT', 'backend must be auto, com, local or mac-cache');
    if (this.platform !== 'darwin' && this.platform !== 'win32') throw localError('UNSUPPORTED_PLATFORM', 'This platform is not supported');
    if (this.platform === 'darwin') {
      if (mode === 'com') throw localError('UNSUPPORTED_BACKEND', 'The COM backend is only available on Windows');
      if (!['auto', 'local', 'mac-cache'].includes(mode)) throw localError('UNSUPPORTED_BACKEND', 'The requested backend is not available on macOS');
      if (id && id.backend !== 'mac-cache') throw localError('BACKEND_ID_MISMATCH', 'Windows IDs cannot be passed to the macOS cache backend. List messages again with backend=mac-cache.');
      if (!this.mac) throw localError('UNSUPPORTED_BACKEND', 'The macOS cache backend is unavailable');
      return this.mac.call(method, args);
    }
    if (mode === 'mac-cache') throw localError('UNSUPPORTED_BACKEND', 'The macOS cache backend is only available on macOS');
    if (id?.backend === 'mac-cache') throw localError('BACKEND_ID_MISMATCH', 'macOS cache IDs cannot be passed to Windows. List messages again with backend=com or local.');
    if (mode === 'local') {
      if (!this.local) throw localError('UNSUPPORTED_BACKEND', 'The local-file backend is unavailable');
      return withWindowsDateBasis(await this.local.call(method, args));
    }
    if (id?.backend === 'local') {
      if (mode === 'com') throw localError('BACKEND_ID_MISMATCH', 'Local IDs cannot be passed to COM. List messages with backend=com again.');
      if (!this.local) throw localError('UNSUPPORTED_BACKEND', 'The local-file backend is unavailable');
      return withWindowsDateBasis(await this.local.call(method, args));
    }
    if (!this.com) throw localError('UNSUPPORTED_BACKEND', 'The COM backend is unavailable');
    try {
      return withWindowsDateBasis({ ...(await this.com.call(method, args)), backend: 'com', data_scope: 'outlook_session' });
    }
    catch (error) {
      // A busy Outlook or an unreadable item must not silently change backends.
      if (mode !== 'auto' || error.code !== 'OUTLOOK_UNAVAILABLE') throw error;
      if (id) throw localError('BACKEND_ID_MISMATCH', 'Outlook is closed and this is a COM ID. Relist messages/folders to obtain local-file IDs.');
      if (!this.local) throw localError('UNSUPPORTED_BACKEND', 'The local-file backend is unavailable');
      return withWindowsDateBasis({ ...(await this.local.call(method, args)), fallback_reason: 'OUTLOOK_UNAVAILABLE' });
    }
  }
  async today(params = {}) {
    const date = params.date ?? singaporeDay(this.now());
    const start = parseDay(date);
    const next = formatDay(new Date(start.getTime() + 86400000));
    const args = { ...params, after: singaporeMidnight(date), before: singaporeMidnight(next, true) };
    delete args.date;
    const result = await this.call('messages', args);
    return { ...result, date, timezone: singaporeTimezone, date_basis: 'received_at' };
  }
  close() { for (const bridge of this.bridges) bridge.close(); }
}
