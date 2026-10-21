import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { createServer } from 'node:http';
import { execFileSync, spawn, spawnSync } from 'node:child_process';
import { chmod, link, lstat, mkdir, mkdtemp, readFile, readdir, rm, stat, symlink, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';
import { openRuntimeLog, prepareRuntime, publishState, readState, runtimeMarkerName } from '../src/runtime.mjs';

async function fixture() {
  const base = await mkdtemp(path.join(os.tmpdir(), 'outlook-runtime-'));
  const runtime = path.join(base, 'runtime');
  await mkdir(runtime);
  return { base, runtime };
}

test('prepareRuntime creates a marker and repairs an owned runtime directory', async t => {
  const item = await fixture();
  t.after(() => rm(item.base, { recursive: true, force: true }));
  if (process.platform === 'win32') return t.skip('POSIX directory mode is not the Windows ACL');
  const managed = path.join(item.base, 'managed');
  await prepareRuntime(managed);
  assert.equal((await stat(managed)).mode & 0o777, 0o700);
  assert.equal((await stat(path.join(managed, runtimeMarkerName))).mode & 0o777, 0o600);
  await chmod(managed, 0o755);
  await prepareRuntime(managed);
  assert.equal((await stat(managed)).mode & 0o777, 0o700);
});

test('prepareRuntime refuses an unmanaged existing directory without changing it', async t => {
  const item = await fixture();
  t.after(() => rm(item.base, { recursive: true, force: true }));
  if (process.platform === 'win32') return t.skip('POSIX directory mode is not the Windows ACL');
  await chmod(item.runtime, 0o755);
  const before = await readdir(item.runtime);
  await assert.rejects(prepareRuntime(item.runtime), /not managed/);
  assert.deepEqual(await readdir(item.runtime), before);
  assert.equal((await stat(item.runtime)).mode & 0o777, 0o755);
});

test('prepareRuntime refuses unrelated files without adding a marker', async t => {
  const item = await fixture();
  t.after(() => rm(item.base, { recursive: true, force: true }));
  if (process.platform === 'win32') return t.skip('POSIX directory mode is not the Windows ACL');
  await chmod(item.runtime, 0o755);
  await writeFile(path.join(item.runtime, 'unrelated.txt'), 'keep');
  await assert.rejects(prepareRuntime(item.runtime), /not managed/);
  assert.deepEqual(await readdir(item.runtime), ['unrelated.txt']);
  assert.equal((await stat(item.runtime)).mode & 0o777, 0o755);
});

test('prepareRuntime refuses a shared parent without changing the managed leaf', async t => {
  const item = await fixture();
  t.after(() => rm(item.base, { recursive: true, force: true }));
  if (process.platform === 'win32') return t.skip('POSIX parent ownership and mode check');
  const managed = path.join(item.base, 'managed');
  await prepareRuntime(managed);
  const marker = path.join(managed, runtimeMarkerName);
  const markerBefore = await readFile(marker, 'utf8');
  await chmod(item.base, 0o777);
  await assert.rejects(prepareRuntime(managed), /untrusted parent/);
  assert.equal((await stat(item.base)).mode & 0o777, 0o777);
  assert.equal((await stat(managed)).mode & 0o777, 0o700);
  assert.deepEqual(await readdir(managed), [runtimeMarkerName]);
  assert.equal(await readFile(marker, 'utf8'), markerBefore);
});

test('prepareRuntime migrates a valid legacy state directory', async t => {
  const item = await fixture();
  t.after(() => rm(item.base, { recursive: true, force: true }));
  if (process.platform === 'win32') return t.skip('POSIX directory mode is not the Windows ACL');
  await chmod(item.runtime, 0o755);
  await writeFile(path.join(item.runtime, 'service.json'), JSON.stringify({
    port: 17379,
    token: 'a'.repeat(64),
    pid: 123,
    startedAt: '2026-10-20T00:00:00.000Z'
  }));
  await writeFile(path.join(item.runtime, 'service.log'), 'legacy log');
  await prepareRuntime(item.runtime);
  assert.equal((await stat(item.runtime)).mode & 0o777, 0o700);
  assert.equal((await stat(path.join(item.runtime, runtimeMarkerName))).mode & 0o777, 0o600);
});

test('prepareRuntime migrates a stopped legacy directory with only the ready log', async t => {
  const item = await fixture();
  t.after(() => rm(item.base, { recursive: true, force: true }));
  if (process.platform === 'win32') return t.skip('POSIX directory mode is not the Windows ACL');
  await chmod(item.runtime, 0o755);
  await writeFile(path.join(item.runtime, 'service.log'), 'Outlook local service ready; pid=12345\nservice stopped\n');
  await prepareRuntime(item.runtime);
  assert.equal((await stat(item.runtime)).mode & 0o777, 0o700);
  assert.equal((await stat(path.join(item.runtime, runtimeMarkerName))).mode & 0o777, 0o600);
});

test('prepareRuntime refuses empty or unrelated legacy logs', async t => {
  const item = await fixture();
  t.after(() => rm(item.base, { recursive: true, force: true }));
  if (process.platform === 'win32') return t.skip('POSIX directory mode is not the Windows ACL');
  const contents = ['', 'service ready; pid=123\n', 'Outlook local service ready; pid=0\n', 'Outlook local service ready; pid=123 extra\n'];
  for (const [index, content] of contents.entries()) {
    const directory = path.join(item.base, `legacy-${index}`);
    await mkdir(directory);
    await chmod(directory, 0o755);
    await writeFile(path.join(directory, 'service.log'), content);
    await assert.rejects(prepareRuntime(directory), /not managed/);
    assert.equal((await stat(directory)).mode & 0o777, 0o755);
    assert.deepEqual(await readdir(directory), ['service.log']);
  }
});

test('prepareRuntime rejects marker symlinks and hardlinks', async t => {
  const item = await fixture();
  t.after(() => rm(item.base, { recursive: true, force: true }));
  if (process.platform === 'win32') return t.skip('Windows has no portable marker-link test');
  for (const kind of ['symlink', 'hardlink']) {
    const managed = path.join(item.base, kind);
    await prepareRuntime(managed);
    const marker = path.join(managed, runtimeMarkerName);
    const protectedTarget = path.join(item.base, `${kind}-target`);
    await writeFile(protectedTarget, 'protected');
    await rm(marker);
    if (kind === 'symlink') await symlink(protectedTarget, marker);
    else await link(protectedTarget, marker);
    await chmod(managed, 0o755);
    await assert.rejects(prepareRuntime(managed));
    assert.equal((await stat(managed)).mode & 0o777, 0o755);
    assert.equal(await readFile(protectedTarget, 'utf8'), 'protected');
  }
});

test('prepareRuntime rejects a marker FIFO without blocking startup', async t => {
  const item = await fixture();
  t.after(() => rm(item.base, { recursive: true, force: true }));
  if (process.platform === 'win32') return t.skip('Windows has no portable FIFO test');
  const managed = path.join(item.base, 'marker-fifo');
  await prepareRuntime(managed);
  const marker = path.join(managed, runtimeMarkerName);
  await rm(marker);
  execFileSync('mkfifo', [marker]);
  await chmod(managed, 0o755);
  await assert.rejects(prepareRuntime(managed));
  assert.equal((await stat(managed)).mode & 0o777, 0o755);
});

test('runtime environment rejects empty, relative, root, home and cwd ancestors', async t => {
  const values = [
    '',
    'relative-runtime',
    path.parse(process.cwd()).root,
    os.homedir(),
    process.cwd(),
    path.dirname(process.cwd())
  ];
  const moduleUrl = new URL('../src/runtime.mjs', import.meta.url).href;
  for (const value of values) {
    const env = { ...process.env, OUTLOOK_LOCAL_STATE_DIR: value };
    const result = spawnSync(process.execPath, ['--input-type=module', '-e', `import(${JSON.stringify(moduleUrl)})`], {
      env,
      encoding: 'utf8'
    });
    assert.notEqual(result.status, 0, `expected rejection for ${JSON.stringify(value)}`);
  }
});

test('openRuntimeLog applies private permissions to an existing log', async t => {
  const item = await fixture();
  t.after(() => rm(item.base, { recursive: true, force: true }));
  const logPath = path.join(item.runtime, 'service.log');
  await writeFile(logPath, 'old log', { mode: 0o644 });
  const handle = await openRuntimeLog(logPath);
  await handle.close();
  if (process.platform !== 'win32') assert.equal((await stat(logPath)).mode & 0o777, 0o600);
});

test('openRuntimeLog rejects symlinks without touching the protected target', async t => {
  const item = await fixture();
  t.after(() => rm(item.base, { recursive: true, force: true }));
  if (process.platform === 'win32') return t.skip('Windows has no portable symlink test permissions');
  const protectedTarget = path.join(item.base, 'HxStore.hxd');
  const logPath = path.join(item.runtime, 'service.log');
  await writeFile(protectedTarget, 'original cache');
  await symlink(protectedTarget, logPath);
  await assert.rejects(openRuntimeLog(logPath));
  assert.equal(await readFile(protectedTarget, 'utf8'), 'original cache');
  assert.equal((await lstat(logPath)).isSymbolicLink(), true);
});

test('openRuntimeLog rejects hardlinks without touching the protected target', async t => {
  const item = await fixture();
  t.after(() => rm(item.base, { recursive: true, force: true }));
  if (process.platform === 'win32') return t.skip('Windows hardlink metadata is not portable here');
  const protectedTarget = path.join(item.base, 'HxStore.hxd');
  const logPath = path.join(item.runtime, 'service.log');
  await writeFile(protectedTarget, 'original cache');
  await link(protectedTarget, logPath);
  await assert.rejects(openRuntimeLog(logPath), /single-link/);
  assert.equal(await readFile(protectedTarget, 'utf8'), 'original cache');
});

test('openRuntimeLog rejects a FIFO without blocking startup', async t => {
  const item = await fixture();
  t.after(() => rm(item.base, { recursive: true, force: true }));
  if (process.platform === 'win32') return t.skip('Windows has no portable FIFO test');
  const logPath = path.join(item.runtime, 'service.log');
  execFileSync('mkfifo', [logPath]);
  await assert.rejects(openRuntimeLog(logPath));
});

test('readState rejects a FIFO without blocking startup', async t => {
  const item = await fixture();
  t.after(() => rm(item.base, { recursive: true, force: true }));
  if (process.platform === 'win32') return t.skip('Windows has no portable FIFO test');
  const statePath = path.join(item.runtime, 'service.json');
  execFileSync('mkfifo', [statePath]);
  await assert.rejects(readState({ target: statePath }), error => error.code === 'SERVICE_UNAVAILABLE');
});

test('readState rejects forged service state under a shared parent', async t => {
  const item = await fixture();
  t.after(() => rm(item.base, { recursive: true, force: true }));
  if (process.platform === 'win32') return t.skip('POSIX parent ownership and mode check');
  const runtime = path.join(item.base, 'managed');
  await prepareRuntime(runtime);
  const statePath = path.join(runtime, 'service.json');
  const state = { port: 17379, token: 'a'.repeat(64) };
  await publishState(state, statePath);
  await chmod(item.base, 0o777);
  const forged = { ...state, token: 'b'.repeat(64) };
  await writeFile(statePath, JSON.stringify(forged));
  await assert.rejects(readState({ target: statePath }), error => error.code === 'SERVICE_UNAVAILABLE');
  assert.deepEqual(JSON.parse(await readFile(statePath, 'utf8')), forged);
  assert.equal((await stat(item.base)).mode & 0o777, 0o777);
  assert.equal((await stat(runtime)).mode & 0o777, 0o700);
});

test('publishState atomically replaces a symlink without touching its target', async t => {
  const item = await fixture();
  t.after(() => rm(item.base, { recursive: true, force: true }));
  if (process.platform === 'win32') return t.skip('Windows has no portable symlink test permissions');
  const protectedTarget = path.join(item.base, 'HxStore.hxd');
  const statePath = path.join(item.runtime, 'service.json');
  const state = { port: 17379, token: 'a'.repeat(64), pid: 123, startedAt: 'now' };
  await writeFile(protectedTarget, 'original cache');
  await symlink(protectedTarget, statePath);
  await publishState(state, statePath);
  assert.equal(await readFile(protectedTarget, 'utf8'), 'original cache');
  assert.deepEqual(JSON.parse(await readFile(statePath, 'utf8')), state);
  assert.equal((await stat(statePath)).mode & 0o777, 0o600);
  assert.equal((await lstat(statePath)).isSymbolicLink(), false);
});

test('readState validates protected roots and requires them for Mac downloads', async t => {
  const item = await fixture();
  t.after(() => rm(item.base, { recursive: true, force: true }));
  const statePath = path.join(item.runtime, 'service.json');
  const state = { port: 17379, token: 'a'.repeat(64), pid: 123, startedAt: 'now', protected_mail_roots: [path.join(item.base, 'profile')] };
  await publishState(state, statePath);
  assert.deepEqual(await readState({ target: statePath, requireMacProtectedRoots: true }), state);
  await publishState({ ...state, protected_mail_roots: 'not-an-array' }, statePath);
  await assert.rejects(readState({ target: statePath }), /Service state unavailable/);
  await publishState({ port: state.port, token: state.token, pid: state.pid, startedAt: state.startedAt }, statePath);
  if (process.platform === 'darwin') {
    await assert.rejects(readState({ target: statePath, requireMacProtectedRoots: true }), error => (
      error.code === 'SERVICE_RESTART_REQUIRED' && /restart service/.test(error.message)
    ));
  }
});

test('CLI download keeps a custom profile protected after its environment is absent', async t => {
  const item = await fixture();
  t.after(() => rm(item.base, { recursive: true, force: true }));
  if (process.platform !== 'darwin') return t.skip('macOS download path guard only');
  const profile = path.join(item.base, 'custom-profile');
  await mkdir(profile);
  const state = { port: 17379, token: 'a'.repeat(64), pid: 123, startedAt: 'now', protected_mail_roots: [profile] };
  await publishState(state, path.join(item.runtime, 'service.json'));
  const output = path.join(profile, 'download.bin');
  const env = { ...process.env, OUTLOOK_LOCAL_STATE_DIR: item.runtime };
  delete env.OUTLOOK_MAC_PROFILE;
  const result = spawnSync(process.execPath, [
    fileURLToPath(new URL('../bin/outlook.mjs', import.meta.url)),
    'download', '--id', 'abcdefgh', '--index', '1', '--out', output
  ], { env, encoding: 'utf8' });
  assert.equal(result.status, 1);
  assert.match(result.stderr, /Path is inside a protected macOS mail cache directory/);
  await assert.rejects(readFile(output));
});

test('CLI download pins the service state across chunks and refuses a restart before writing', async t => {
  const item = await fixture();
  const stateTarget = path.join(item.runtime, 'service.json');
  const output = path.join(item.base, 'download.bin');
  const profileA = path.join(item.base, 'profile-a');
  const profileB = path.join(item.base, 'profile-b');
  await mkdir(profileA);
  await mkdir(profileB);
  const bytes = Buffer.from('ABCD');
  const sha256 = createHash('sha256').update(bytes).digest('hex');
  const tokenA = 'a'.repeat(64);
  const tokenB = 'b'.repeat(64);
  let originalRequests = 0;
  let replacementRequests = 0;
  let replacementPort;

  const send = (res, data, status = 200) => {
    res.writeHead(status, { 'Content-Type': 'application/json', Connection: 'close' });
    res.end(JSON.stringify(data));
  };
  const originalServer = createServer((req, res) => {
    void (async () => {
      for await (const _chunk of req) { /* drain request */ }
      if (req.method !== 'POST' || req.url !== '/rpc') return send(res, { error: { code: 'NOT_FOUND', message: 'not found' } }, 404);
      originalRequests += 1;
      if (originalRequests === 1) {
        await publishState({
          port: replacementPort,
          token: tokenB,
          pid: 123,
          startedAt: '2026-10-20T00:00:00.000Z',
          protected_mail_roots: [profileB]
        }, stateTarget);
        return send(res, { result: {
          data: bytes.subarray(0, 2).toString('base64'),
          size_bytes: bytes.length,
          sha256,
          offset: 0,
          next_offset: 2
        } });
      }
      return send(res, { result: {
        data: bytes.subarray(2).toString('base64'),
        size_bytes: bytes.length,
        sha256,
        offset: 2,
        next_offset: null
      } });
    })().catch(error => send(res, { error: { code: 'TEST_ERROR', message: error.message } }, 500));
  });
  const replacementServer = createServer((req, res) => {
    void (async () => {
      for await (const _chunk of req) { /* drain request */ }
      replacementRequests += 1;
      return send(res, { result: {
        data: bytes.subarray(2).toString('base64'),
        size_bytes: bytes.length,
        sha256,
        offset: 2,
        next_offset: null
      } });
    })().catch(error => send(res, { error: { code: 'TEST_ERROR', message: error.message } }, 500));
  });
  const listen = server => new Promise((resolve, reject) => {
    server.once('error', reject);
    server.listen(0, '127.0.0.1', resolve);
  });
  const close = server => server.listening ? new Promise(resolve => server.close(resolve)) : Promise.resolve();
  await Promise.all([listen(originalServer), listen(replacementServer)]);
  replacementPort = replacementServer.address().port;
  const originalPort = originalServer.address().port;
  await publishState({
    port: originalPort,
    token: tokenA,
    pid: 123,
    startedAt: '2026-10-20T00:00:00.000Z',
    protected_mail_roots: [profileA]
  }, stateTarget);
  t.after(async () => {
    await close(originalServer);
    await close(replacementServer);
    await rm(item.base, { recursive: true, force: true });
  });

  const env = { ...process.env, OUTLOOK_LOCAL_STATE_DIR: item.runtime };
  delete env.OUTLOOK_MAC_PROFILE;
  const child = spawn(process.execPath, [
    fileURLToPath(new URL('../bin/outlook.mjs', import.meta.url)),
    'download', '--id', 'abcdefgh', '--index', '1', '--out', output
  ], { env, stdio: ['ignore', 'pipe', 'pipe'] });
  let stderr = '';
  child.stderr.setEncoding('utf8');
  child.stderr.on('data', chunk => { stderr += chunk; });
  const exitCode = await new Promise((resolve, reject) => {
    child.once('error', reject);
    child.once('close', resolve);
  });
  assert.equal(exitCode, 1);
  assert.match(stderr, /Service state changed during download/);
  assert.equal(originalRequests, 2);
  assert.equal(replacementRequests, 0);
  await assert.rejects(readFile(output));
});
