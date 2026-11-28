import assert from 'node:assert/strict';
import test from 'node:test';
import { runCalendarWorkflow } from '../src/calendar-workflow.mjs';

function nativeFixture(resolve) {
  const requests = [];
  return {
    requests,
    queryNative(request) {
      requests.push(request);
      return resolve(request);
    },
  };
}

test('only a separate explicit authorization action can request full access', async () => {
  const native = nativeFixture(request => request.action === 'status' ? { authorization: 'not_determined' }
    : request.action === 'authorize' ? { authorization: 'full_access' }
      : { error: { code: 'CALENDAR_AUTH_REQUIRED', message: 'Full access is required' } });
  assert.deepEqual(await runCalendarWorkflow('status', {}, native), {
    source: 'eventkit', data_scope: 'local_calendar_store', authorization: 'not_determined', can_read: false,
  });
  await assert.rejects(runCalendarWorkflow('calendars', {}, native), { code: 'CALENDAR_AUTH_REQUIRED' });
  assert.deepEqual(native.requests, [{ action: 'status' }, { action: 'calendars' }]);
  assert.equal((await runCalendarWorkflow('authorize', {}, native)).can_read, true);
  assert.deepEqual(native.requests.at(-1), { action: 'authorize' });

  const writeOnly = nativeFixture(() => ({ authorization: 'write_only' }));
  assert.equal((await runCalendarWorkflow('status', {}, writeOnly)).can_read, false);
  await assert.rejects(runCalendarWorkflow('authorize', {}, writeOnly), { code: 'CALENDAR_PERMISSION_DENIED' });
});

test('calendar pages sort by source and title and reject a stale snapshot', async () => {
  const calendar = (id, title, source_title) => ({ id, title, source_id: source_title, source_title });
  const work = calendar('b', 'Office', 'Exchange');
  const personal = calendar('a', 'Home', 'iCloud');
  const rows = [personal, work];
  const native = nativeFixture(() => ({ items: [...rows] }));
  const first = await runCalendarWorkflow('calendars', { limit: 1 }, native);
  assert.deepEqual(first.items, [work]);
  assert.equal(first.complete, false);
  assert.equal(first.total_count, 2);
  const second = await runCalendarWorkflow('calendars', { cursor: first.next_cursor }, native);
  assert.deepEqual(second.items, [personal]);
  assert.equal(second.complete, true);
  assert.equal(second.next_cursor, null);

  rows.push(calendar('c', 'Other', 'Exchange'));
  await assert.rejects(runCalendarWorkflow('calendars', { cursor: first.next_cursor }, native), {
    code: 'CALENDAR_CURSOR_STALE',
  });
});

test('agenda uses a bounded zoned date range and preserves distinct recurring occurrences', async () => {
  const event = {
    id: 'recurrence-id', calendar_id: 'Exact-ID', title: 'Meeting', start_at: '2026-03-08T15:00:00.000Z',
    end_at: '2026-03-08T16:00:00.000Z', occurrence_at: '2026-03-08T15:00:00.000Z',
    recurring: true, canceled: false, all_day: false, time_zone: 'America/Los_Angeles',
  };
  const next = {
    ...event, start_at: '2026-03-09T15:00:00.000Z', end_at: '2026-03-09T16:00:00.000Z',
    occurrence_at: '2026-03-09T15:00:00.000Z',
  };
  const allDay = {
    id: 'all-day', calendar_id: 'Another-ID', title: 'Travel', start_at: '2026-03-08T08:00:00.000Z',
    end_at: '2026-03-09T07:00:00.000Z', recurring: false, canceled: false, all_day: true,
    time_zone: 'America/Los_Angeles', start_date: '2026-03-08', end_date_exclusive: '2026-03-09',
  };
  const native = nativeFixture(request => ({
    items: [next, event, allDay].map(item => request.include_details ? { ...item, notes: 'private' } : item),
  }));
  const args = {
    from_date: '2026-03-08', to_date: '2026-03-09', time_zone: 'America/Los_Angeles',
    calendar_ids: ['Exact-ID', 'Another-ID'], limit: 2,
  };
  const first = await runCalendarWorkflow('agenda', args, native);
  assert.deepEqual(native.requests[0], {
    action: 'agenda', from_date: '2026-03-08', end_date_exclusive: '2026-03-10',
    time_zone: 'America/Los_Angeles', calendar_ids: ['Another-ID', 'Exact-ID'], query: null,
    include_details: false, include_canceled: false,
  });
  assert.deepEqual(first.items, [allDay, event]);
  const second = await runCalendarWorkflow('agenda', { ...args, cursor: first.next_cursor }, native);
  assert.deepEqual(second.items, [next]);
  await assert.rejects(runCalendarWorkflow('agenda', {
    ...args, cursor: first.next_cursor, calendar_ids: ['exact-id'],
  }, native), { code: 'INVALID_ARGUMENT' });
  const detailed = await runCalendarWorkflow('agenda', {
    ...args, include_details: true, include_canceled: true, query: ' Meeting ',
  }, native);
  assert.equal(detailed.items[0].notes, 'private');
  assert.equal(native.requests.at(-1).include_canceled, true);
  assert.equal(native.requests.at(-1).query, 'Meeting');
});

test('relative agenda starts in the selected timezone and a cursor keeps its original date', async () => {
  const native = nativeFixture(() => ({ items: [
    { id: 'a', start_at: '2026-11-25T16:00:00Z' },
    { id: 'b', start_at: '2026-11-25T17:00:00Z' },
  ] }));
  const beforeMidnight = new Date('2026-11-24T16:01:00Z');
  const first = await runCalendarWorkflow('agenda', { limit: 1 }, { ...native, now: beforeMidnight });
  assert.equal(first.from_date, '2026-11-25');
  assert.equal(first.end_date_exclusive, '2026-11-26');
  const forLA = await runCalendarWorkflow('agenda', { time_zone: 'America/Los_Angeles', days: 2 },
    { ...native, now: beforeMidnight });
  assert.equal(forLA.from_date, '2026-11-24');
  assert.equal(forLA.end_date_exclusive, '2026-11-26');
  const nextPage = await runCalendarWorkflow('agenda', { cursor: first.next_cursor },
    { ...native, now: new Date('2026-11-25T16:05:00Z') });
  assert.equal(nextPage.from_date, first.from_date);
  assert.equal(nextPage.items[0].id, 'b');
});

test('invalid scope is rejected before native access', async () => {
  const native = nativeFixture(() => assert.fail('invalid arguments must never reach EventKit'));
  for (const args of [
    { from_date: '2026-02-28' },
    { date: '2026-02-30' },
    { date: '2026-02-28', days: 3 },
    { from_date: '2026-05-01', to_date: '2026-02-28' },
    { from_date: '2026-01-01', to_date: '2027-01-03' },
    { time_zone: null },
    { time_zone: '+08:00' },
    { time_zone: 'Unknown/Zone' },
    { calendar_ids: ['Duplicate', 'Duplicate'] },
    { include_details: 'yes' },
    { query: ' ' },
    { cursor: 'too-short' },
    { limit: 101 },
  ]) {
    await assert.rejects(runCalendarWorkflow('agenda', args, native), { code: 'INVALID_ARGUMENT' });
  }
  await assert.rejects(runCalendarWorkflow('authorize', { auto: true }, native), { code: 'INVALID_ARGUMENT' });
  assert.deepEqual(native.requests, []);
});
