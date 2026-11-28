import { localError } from './offline.mjs';

export const singaporeTimezone = 'Asia/Singapore';
const minimumToday = '1970-01-01';
const maximumToday = '9999-12-30';
const singaporeOffsetFormatter = new Intl.DateTimeFormat('en', {
  timeZone: singaporeTimezone,
  timeZoneName: 'longOffset'
});

function invalidDate(message) {
  throw localError('INVALID_ARGUMENT', message);
}

export function parseDay(value, allowNextDay = false) {
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

export function formatDay(date) {
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

export function singaporeMidnight(value, allowNextDay = false) {
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

export function singaporeDay(value) {
  const parts = Object.fromEntries(new Intl.DateTimeFormat('en', {
    timeZone: singaporeTimezone,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit'
  }).formatToParts(new Date(value)).filter(part => part.type !== 'literal').map(part => [part.type, part.value]));
  if (!parts.year || !parts.month || !parts.day) invalidDate('Cannot determine Singapore date');
  return `${parts.year}-${parts.month}-${parts.day}`;
}
