import test from 'node:test';
import assert from 'node:assert/strict';
import { validateCall } from '../src/tools.mjs';
import { BackendRouter } from '../src/backend.mjs';

test('COM body mode is forwarded only to COM, leaving other backends unchanged', async () => {
  const seen = [];
  const bridge = name => ({ call: async (method, params) => { seen.push({ name, params }); return {}; } });
  const win = new BackendRouter({ platform: 'win32', com: bridge('com'), local: bridge('local') });
  const mac = new BackendRouter({ platform: 'darwin', mac: bridge('mac') });
  const params = { query: 'x', search_body: true, com_body_mode: 'indexed' };
  await win.call('search', { ...params, backend: 'com' });
  await win.call('search', { ...params, backend: 'local' });
  await mac.call('search', params);
  assert.equal(seen[0].params.com_body_mode, 'indexed');
  assert.equal(seen[1].params.com_body_mode, undefined);
  assert.equal(seen[2].params.com_body_mode, undefined);
});

test('COM body mode is optional and accepts auto, direct or indexed modes', () => {
  assert.equal(validateCall('search', { query: 'x', search_body: true }).com_body_mode, undefined);
  for (const mode of ['auto', 'direct', 'indexed']) assert.equal(validateCall('search', { query: 'x', com_body_mode: mode }).com_body_mode, mode);
  assert.throws(() => validateCall('search', { query: 'x', com_body_mode: 'invalid' }), { code: 'INVALID_ARGUMENT' });
});

test('search accepts keyword, person, or their intersection with an explicit role', () => {
  assert.equal(validateCall('search', { query: '预算' }).person_role, 'any');
  const person = validateCall('search', { person: ' 张三 ', person_role: 'to' });
  assert.equal(person.person, '张三');
  assert.equal(person.query, '');
  const combined = validateCall('search', { query: '预算', person: 'a@example.test', person_role: 'cc' });
  assert.equal(combined.query, '预算');
  assert.equal(combined.person_role, 'cc');
});

test('search rejects empty conditions, unsupported roles and role without person', () => {
  for (const params of [{}, { query: '  ' }, { person: ' ' },
    { query: 'x', person_role: 'to' }, { person: 'x', person_role: 'bcc' },
    { person: 'x'.repeat(501) }]) {
    assert.throws(() => validateCall('search', params), { code: 'INVALID_ARGUMENT' });
  }
});
