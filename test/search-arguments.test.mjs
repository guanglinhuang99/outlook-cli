import test from 'node:test';
import assert from 'node:assert/strict';
import { validateCall } from '../src/tools.mjs';

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
