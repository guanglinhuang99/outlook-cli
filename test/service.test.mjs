import test from 'node:test';
import assert from 'node:assert/strict';
import net from 'node:net';
import { execFile } from 'node:child_process';
import { mkdtemp, writeFile, rm } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { promisify } from 'node:util';
import { Client, InMemoryTransport, StreamableHTTPClientTransport } from '@modelcontextprotocol/client';
import { startService, authenticate } from '../src/service.mjs';
import { validateCall } from '../src/tools.mjs';
import { createMcp } from '../src/mcp.mjs';

test('installed MCP checker accepts the real platform-aware annotations', async t => {
  const directory = await mkdtemp(path.join(os.tmpdir(), 'outlook-mcp-check-'));
  t.after(() => rm(directory, { recursive: true, force: true }));
  const config = path.join(directory, 'mcp.json');
  await writeFile(config, JSON.stringify({ mcpServers: { 'outlook-local': {
    command: process.execPath, args: [fileURLToPath(new URL('../bin/outlook.mjs', import.meta.url)), 'mcp']
  } } }));
  const result = await promisify(execFile)(process.execPath, [
    fileURLToPath(new URL('../scripts/check-mcp.mjs', import.meta.url)), config
  ], { timeout: 20000 });
  const report = JSON.parse(result.stdout);
  assert.equal(report.mcp, 'ready');
  assert.equal(report.tools.length, 9);
  assert.ok(report.tools.includes('outlook_today'));
});

test('shared MCP annotations reflect the backend platform', async t => {
  async function listAnnotations(platform) {
    const server = createMcp(async () => ({}), { platform });
    const client = new Client({ name: `annotation-test-${platform}`, version: '1.0.0' });
    const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
    await server.connect(serverTransport);
    await client.connect(clientTransport);
    t.after(async () => {
      await client.close();
      await server.close();
    });
    return (await client.listTools()).tools;
  }

  const windows = await listAnnotations('win32');
  assert.equal(windows.length, 9);
  assert.ok(windows.every(tool => tool.annotations.readOnlyHint === true && tool.annotations.destructiveHint === false));
  const mac = await listAnnotations('darwin');
  assert.equal(mac.length, 9);
  assert.ok(mac.every(tool => tool.annotations.readOnlyHint === false && tool.annotations.destructiveHint === true));
});

test('input contract rejects writes, unknown fields, bad bounds, malformed dates', () => {
  for (const name of ['send', 'delete', 'move', '__proto__', 'constructor']) assert.throws(() => validateCall(name, {}));
  assert.throws(() => validateCall('read', { id: 'whatever', max_chars: 200001 }));
  assert.throws(() => validateCall('messages', { unread: 'false' }));
  assert.throws(() => validateCall('messages', { after: '2026-10-20' }));
  assert.throws(() => validateCall('messages', { after: '2026-10-21T00:00:00Z', before: '2026-10-20T00:00:00Z' }));
  assert.throws(() => validateCall('health', { command: 'anything' }));
  assert.throws(() => validateCall('today', { date: '2026-02-29' }));
  assert.throws(() => validateCall('today', { date: '1969-12-31' }));
  assert.throws(() => validateCall('today', { after: '2026-10-20T00:00:00Z' }));
  assert.throws(() => validateCall('today', { limit: true }));
  assert.equal(validateCall('today', { date: '2026-10-20', folder: 'sent' }).date, '2026-10-20');
  assert.equal(validateCall('messages', {}).limit, 25);
  assert.equal(validateCall('search', { query: "中文 'quoted'" }).query, "中文 'quoted'");
});

test('HTTP rejects missing token, browser origins and rebinding hosts', () => {
  const token = 'a'.repeat(64);
  const valid = { host: '127.0.0.1:17379', authorization: `Bearer ${token}` };
  assert.equal(authenticate({ headers: valid }, token, 17379), true);
  for (const changes of [{ authorization: '' }, { authorization: 'Bearer bad' }, { origin: 'https://example.com' }, { host: 'attacker.test:17379' }]) {
    assert.equal(authenticate({ headers: { ...valid, ...changes } }, token, 17379), false);
  }
});

test('RPC and official MCP HTTP client share validated read-only backend', async t => {
  const reservation = net.createServer();
  await new Promise(resolve => reservation.listen(0, '127.0.0.1', resolve));
  const port = reservation.address().port;
  await new Promise(resolve => reservation.close(resolve));
  const calls = [];
  const bridge = { queue: [], call: async (method, params) => { calls.push({ method, params }); return { connected: true, sample: '中文' }; }, close() {} };
  const service = await startService({ bridge, persist: false, listenPort: port });
  t.after(() => service.close());
  const base = `http://127.0.0.1:${port}`;
  const headers = { Authorization: `Bearer ${service.token}`, 'Content-Type': 'application/json' };
  assert.equal((await fetch(base + '/status')).status, 403);
  let response = await fetch(base + '/rpc', { method: 'POST', headers, body: JSON.stringify({ method: 'send', params: {} }) });
  assert.equal(response.status, 400);
  assert.equal(calls.length, 0);
  response = await fetch(base + '/rpc', { method: 'POST', headers, body: '{broken' });
  assert.equal(response.status, 400);
  response = await fetch(base + '/rpc', { method: 'POST', headers, body: JSON.stringify({ method: 'messages', params: { limit: 2 } }) });
  assert.equal(response.status, 200);
  assert.equal((await response.json()).result.sample, '中文');
  assert.equal(calls.at(-1).params.limit, 2);
  const client = new Client({ name: 'integration-test', version: '1.0.0' });
  t.after(() => client.close());
  await client.connect(new StreamableHTTPClientTransport(new URL(base + '/mcp'), { requestInit: { headers } }));
  const list = await client.listTools();
  assert.equal(list.tools.length, 9);
  const cacheWrites = process.platform !== 'win32';
  assert.ok(list.tools.every(tool => tool.annotations.readOnlyHint === !cacheWrites && tool.annotations.destructiveHint === cacheWrites));
  const result = await client.callTool({ name: 'outlook_health', arguments: {} });
  assert.equal(result.isError, undefined);
  assert.equal(result.structuredContent.connected, true);
});
