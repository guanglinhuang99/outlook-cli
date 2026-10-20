import http from 'node:http';
import { randomBytes, timingSafeEqual } from 'node:crypto';
import { writeFile, readFile, unlink } from 'node:fs/promises';
import { NodeStreamableHTTPServerTransport } from '@modelcontextprotocol/node';
import { BackendRouter } from './backend.mjs';
import { validateCall } from './tools.mjs';
import { createMcp } from './mcp.mjs';
import { prepareRuntime, statePath, port } from './runtime.mjs';

export function authenticate(req, token, listenPort) {
  if (![ `127.0.0.1:${listenPort}`, `localhost:${listenPort}` ].includes(req.headers.host)) return false;
  // Native local clients do not send Origin. Never allow browser-origin calls.
  if (req.headers.origin !== undefined) return false;
  const actual = Buffer.from(req.headers.authorization ?? '');
  const expected = Buffer.from(`Bearer ${token}`);
  return actual.length === expected.length && timingSafeEqual(actual, expected);
}
async function parseBody(req) {
  if (!(req.headers['content-type'] ?? '').toLowerCase().startsWith('application/json')) {
    throw Object.assign(new Error('Content-Type must be application/json'), { code: 'INVALID_ARGUMENT', status: 415 });
  }
  let size = 0;
  const parts = [];
  for await (const chunk of req) {
    size += chunk.length;
    if (size > 65536) throw Object.assign(new Error('Request body exceeds 64 KiB'), { code: 'INVALID_ARGUMENT', status: 413 });
    parts.push(chunk);
  }
  try { return JSON.parse(Buffer.concat(parts).toString('utf8')); }
  catch { throw Object.assign(new Error('Invalid JSON'), { code: 'INVALID_ARGUMENT', status: 400 }); }
}
function json(res, status, data) {
  if (res.writableEnded || res.destroyed) return;
  res.writeHead(status, { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store', 'X-Content-Type-Options': 'nosniff' });
  res.end(JSON.stringify(data));
}
export async function startService({ bridge = new BackendRouter(), listenPort = port, token = randomBytes(32).toString('hex'), persist = true } = {}) {
  const startedAt = new Date().toISOString();
  let closing = false;
  const sockets = new Set();
  const invoke = (method, params) => bridge.call(method, validateCall(method, params));
  const server = http.createServer(async (req, res) => {
    try {
      if (!authenticate(req, token, listenPort)) return json(res, 403, { error: { code: 'FORBIDDEN', message: 'Local bearer authentication required' } });
      const url = req.url;
      if (url === '/status' && req.method === 'GET') return json(res, 200, { service: 'outlook-local', version: '0.2.1', pid: process.pid, port: listenPort, started_at: startedAt, queue_length: bridge.queue.length, busy: Boolean(bridge.active), read_only: true });
      if (closing) return json(res, 503, { error: { code: 'STOPPING', message: 'Service is stopping' } });
      if (url === '/shutdown' && req.method === 'POST') {
        await parseBody(req);
        json(res, 200, { stopped: true });
        setTimeout(() => void close(), 50);
        return;
      }
      if (url === '/rpc' && req.method === 'POST') {
        const body = await parseBody(req);
        if (!body || Array.isArray(body) || typeof body.method !== 'string' || Object.keys(body).some(k => !['method', 'params'].includes(k))) {
          throw Object.assign(new Error('Expected {method, params}'), { code: 'INVALID_ARGUMENT' });
        }
        const result = await invoke(body.method, body.params);
        return json(res, 200, { result });
      }
      if (url === '/mcp' && req.method === 'POST') {
        const body = await parseBody(req);
        const mcp = createMcp(invoke);
        const transport = new NodeStreamableHTTPServerTransport({ sessionIdGenerator: undefined, enableJsonResponse: true });
        res.once('close', () => { void mcp.close(); });
        await mcp.connect(transport);
        await transport.handleRequest(req, res, body);
        return;
      }
      if (url === '/mcp') {
        res.setHeader('Allow', 'POST');
        return json(res, 405, { error: { code: 'METHOD_NOT_ALLOWED', message: 'Stateless MCP endpoint accepts POST' } });
      }
      json(res, 404, { error: { code: 'NOT_FOUND', message: 'Unknown route or unsupported HTTP method' } });
    } catch (error) {
      json(res, error.status ?? (error.code === 'INVALID_ARGUMENT' ? 400 : error.code === 'BUSY' ? 429 : 503), { error: { code: error.code ?? 'SERVICE_ERROR', message: error.message } });
    }
  });
  server.requestTimeout = 15000;
  server.headersTimeout = 10000;
  server.on('connection', socket => { sockets.add(socket); socket.once('close', () => sockets.delete(socket)); });
  async function close() {
    if (closing) return;
    closing = true;
    bridge.close();
    server.close();
    for (const socket of sockets) socket.destroy();
    if (persist) {
      try {
        const state = JSON.parse(await readFile(statePath, 'utf8'));
        if (state.token === token) await unlink(statePath);
      } catch { }
    }
  }
  if (persist) await prepareRuntime();
  await new Promise((resolve, reject) => { server.once('error', reject); server.listen(listenPort, '127.0.0.1', resolve); });
  if (persist) {
    try { await writeFile(statePath, JSON.stringify({ port: listenPort, token, pid: process.pid, startedAt }), { mode: 0o600 }); }
    catch (error) { await close(); throw error; }
  }
  return { server, token, close };
}
