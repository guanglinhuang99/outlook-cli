import assert from 'node:assert/strict';
import { mkdtemp, readdir, rm, stat, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { graphAccessToken, graphLogin, graphLogout } from '../src/graph-auth.mjs';

const client = '00001111-aaaa-2222-bbbb-3333cccc4444';
const env = { OUTLOOK_GRAPH_CLIENT_ID: client, OUTLOOK_GRAPH_TENANT_ID: 'organizations' };
const reply = (data, status = 200) => Response.json(data, { status });

async function withState(t) {
  const root = await mkdtemp(path.join(os.tmpdir(), 'outlook-graph-auth-test-'));
  t.after(() => rm(root, { recursive: true, force: true }));
  return path.join(root, 'auth');
}

test('explicit device login caches an access grant privately and refreshes it without printing tokens', async t => {
  const directory = await withState(t);
  let time = 100000;
  const requests = [];
  const announced = [];
  const replies = [
    reply({ device_code: 'device-fixture', user_code: 'ABC-1234', expires_in: 300, interval: 5 }),
    reply({ error: 'authorization_pending' }, 400),
    reply({ access_token: 'access-one', refresh_token: 'refresh-one', expires_in: 1800 }),
    reply({ access_token: 'access-two', refresh_token: 'refresh-two', expires_in: 1800 })
  ];
  const options = { env, directory, now: () => time, sleep: async ms => { time += ms; },
    notify: value => announced.push(value), fetchImpl: async (url, opts) => {
      assert.equal(new URL(url).origin, 'https://login.microsoftonline.com');
      assert.equal(opts.redirect, 'error');
      requests.push({ route: new URL(url).pathname, body: Object.fromEntries(opts.body) });
      return replies.shift();
    } };
  assert.equal(await graphAccessToken(options), undefined);
  const status = await graphLogin({ send: true }, options);
  assert.equal(status.authorized, true);
  assert.equal(status.can_refresh, true);
  assert.ok(!JSON.stringify(status).includes('access-one'));
  assert.deepEqual(announced, [{ url: 'https://microsoft.com/devicelogin', code: 'ABC-1234' }]);
  assert.match(requests[0].body.scope, /Mail\.ReadWrite/);
  assert.match(requests[0].body.scope, /Mail\.Send/);
  assert.equal(await graphAccessToken(options), 'access-one');
  for (const name of await readdir(directory)) {
    if (process.platform !== 'win32') assert.equal((await stat(path.join(directory, name))).mode & 0o077, 0);
  }
  time += 1790000;
  assert.equal(await graphAccessToken({ ...options, refresh: false }), undefined);
  assert.equal(requests.length, 3);
  assert.equal(await graphAccessToken(options), 'access-two');
  assert.equal(requests[3].body.refresh_token, 'refresh-one');
  assert.equal(requests[3].body.grant_type, 'refresh_token');
  await assert.rejects(graphAccessToken({ ...options, env: { ...env, OUTLOOK_GRAPH_TENANT_ID: 'consumers' } }),
    { code: 'GRAPH_AUTH_APP_CHANGED' });
  assert.deepEqual(await graphLogout(options), { logged_out: true, removed_local_grant: true });
  assert.equal(await graphAccessToken(options), undefined);
  assert.equal((await readdir(directory)).filter(name => /^grant-/.test(name)).length, 0);
});

test('missing app permission never starts sign-in; external tokens remain opt-in', async t => {
  const directory = await withState(t);
  const options = { directory, env: {}, fetchImpl: async () => assert.fail('Unexpected authorization request') };
  await assert.rejects(graphLogin({}, options), { code: 'GRAPH_APP_REQUIRED' });
  await assert.rejects(graphLogin({ send: true, read_only: true }, options), { code: 'INVALID_ARGUMENT' });
  assert.equal(await graphAccessToken({ ...options, env: { OUTLOOK_GRAPH_ACCESS_TOKEN: 'delegated-fixture' } }),
    'delegated-fixture');
  await assert.rejects(stat(directory), { code: 'ENOENT' });
});

test('credentials are not read from an unmanaged or invalid pointer', async t => {
  const directory = await withState(t);
  const args = { directory, env, now: () => 0, sleep: async () => {}, notify: () => {},
    fetchImpl: async url => url.endsWith('/devicecode')
      ? reply({ device_code: 'fixture', user_code: 'ABC-1234', expires_in: 300, interval: 5 })
      : reply({ error: 'authorization_declined' }, 400) };
  await assert.rejects(graphLogin({ read_only: true }, args), { code: 'GRAPH_AUTH_DENIED' });
  await writeFile(path.join(directory, 'current.json'), '{"id":"../../unexpected"}', { mode: 0o600 });
  await assert.rejects(graphAccessToken(args), { code: 'GRAPH_AUTH_STATE_INVALID' });
});
