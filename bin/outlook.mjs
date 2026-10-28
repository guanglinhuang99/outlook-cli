#!/usr/bin/env node
import { spawn } from 'node:child_process';
import { writeFile } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { parseArgs } from 'node:util';
import { setTimeout as delay } from 'node:timers/promises';
import { openRuntimeLog, prepareRuntime, runtimeDir, root, statePath, readState, request, call } from '../src/runtime.mjs';
import { assertSafeMacPath } from '../src/mac-paths.mjs';
import { validateCall, definitions } from '../src/tools.mjs';

const print = value => process.stdout.write(JSON.stringify(value, null, 2) + '\n');
function sameDownloadState(left, right) {
  const leftRoots = left.protected_mail_roots;
  const rightRoots = right.protected_mail_roots;
  return left.port === right.port
    && left.token === right.token
    && ((leftRoots === undefined && rightRoots === undefined)
      || (Array.isArray(leftRoots) && Array.isArray(rightRoots)
        && leftRoots.length === rightRoots.length
        && leftRoots.every((root, index) => root === rightRoots[index])));
}

function downloadStateChanged() {
  return Object.assign(new Error('Service state changed during download; refusing to write attachment'), {
    code: 'SERVICE_RESTART_REQUIRED'
  });
}

const help = `Outlook Local Service — read-only, no AI

  outlook-local start | stop | restart | status | health
  outlook-local stores
  outlook-local folders [--folder inbox|<folder-id>] [--limit 25] [--offset 0]
  outlook-local messages [--folder inbox] [--limit 25] [--unread true]
  outlook-local today [--folder inbox|sent] [--limit 25] [--date YYYY-MM-DD]
  outlook-local search [--query text] [--person name-or-email] [--person-role any|sender|to|cc]
                       [--search-body true] [--com-body-mode auto|direct|indexed] [--scan-limit 500]
  outlook-local read --id <message-id> [--max-chars 20000] [--body-offset 0]
  outlook-local attachments --id <message-id>
  outlook-local download --id <message-id> --index 1 --out <file>
  outlook-local mcp              MCP stdio adapter; requires running service
  outlook-local mcp-config       Print MCP stdio configuration (no secrets)
  outlook-local rpc <operation> --json '<JSON object>'
  outlook-local serve            Run service in foreground

Messages/search: --after/--before ISO time with timezone, --offset,
--limit, --scan-limit. Search covers one folder per call; follow next_offset.
All mailbox commands accept --backend auto|com|local|mac-cache (default auto).
auto uses Windows COM/local OST/PST or the macOS HxStore cache by platform.
Both platforms filter by received_at; macOS uses verified LastDeliveryTime.
On macOS, Python 3.11+ is required; unread/body-search/HTML are unsupported.
Local reads contain cached data only; COM and local IDs are different.
Use stores -> root_id -> folders to discover other folders.
All output is JSON, except help. Errors go to stderr with nonzero exit status.
`;

async function start() {
  await prepareRuntime();
  try { return { ...(await request('/status', undefined, { timeoutMs: 1500 })), already_running: true }; } catch { }
  const log = await openRuntimeLog(path.join(runtimeDir, 'service.log'));
  const child = spawn(process.execPath, [fileURLToPath(import.meta.url), 'serve'], {
    cwd: root, detached: true, windowsHide: true, stdio: ['ignore', log.fd, log.fd], env: process.env
  });
  let spawnError;
  child.on('error', error => { spawnError = error; });
  child.unref();
  await log.close();
  for (let i = 0; i < 60; i++) {
    if (spawnError) throw spawnError;
    await delay(250);
    try { return await request('/status', undefined, { timeoutMs: 800 }); } catch { }
    if (child.exitCode !== null) break;
  }
  throw new Error(`Service did not start. Check ${path.join(runtimeDir, 'service.log')}`);
}

async function stop() {
  const result = await request('/shutdown', {});
  for (let i = 0; i < 40; i++) {
    await delay(100);
    try { await request('/status', undefined, { timeoutMs: 500 }); }
    catch { return result; }
  }
  throw new Error('Service did not stop within the deadline');
}

async function main() {
  const [command = 'help', ...args] = process.argv.slice(2);
  if (['help', '--help', '-h'].includes(command)) { process.stdout.write(help); return; }
  if (command === 'start') return print(await start());
  if (command === 'stop') return print(await stop());
  if (command === 'restart') { await stop(); return print(await start()); }
  if (command === 'status') return print(await request('/status'));
  if (command === 'mcp-config') return print({ mcpServers: { 'outlook-local': { command: process.execPath, args: [fileURLToPath(import.meta.url), 'mcp'], env: { OUTLOOK_LOCAL_STATE_DIR: runtimeDir } } } });
  if (command === 'serve') {
    const { startService } = await import('../src/service.mjs');
    const service = await startService();
    const close = () => { void service.close(); };
    process.once('SIGINT', close);
    process.once('SIGTERM', close);
    process.stderr.write(`Outlook local service ready; pid=${process.pid}\n`);
    return;
  }
  if (command === 'mcp') {
    const { createMcp } = await import('../src/mcp.mjs');
    const { StdioServerTransport } = await import('@modelcontextprotocol/server/stdio');
    const server = createMcp(call);
    await server.connect(new StdioServerTransport());
    process.stdin.once('end', () => { void server.close(); });
    return;
  }
  const operation = command === 'download' ? 'read_attachment' : command === 'rpc' ? args.shift() : command;
  if (!Object.hasOwn(definitions, operation)) throw new Error(`Unknown command: ${command}`);
  const parsed = parseArgs({ args, options: Object.fromEntries([
    'backend', 'folder', 'limit', 'offset', 'unread', 'after', 'before', 'date', 'scan-limit', 'query', 'person', 'person-role', 'search-body', 'com-body-mode', 'id', 'format', 'body-offset', 'max-chars', 'index', 'max-bytes', 'byte-offset', 'chunk-bytes', 'out', 'json'
  ].map(key => [key, { type: 'string' }])), strict: true, allowPositionals: false });
  const values = { ...parsed.values };
  const outputPath = values.out;
  delete values.out;
  let params;
  if (values.json !== undefined) {
    if (Object.keys(values).length !== 1) throw new Error('--json cannot be combined with operation flags');
    params = JSON.parse(values.json);
  } else {
    params = {};
    for (const [key, value] of Object.entries(values)) {
      const name = key.replaceAll('-', '_');
      if (['unread', 'search_body'].includes(name)) {
        if (!['true', 'false'].includes(value)) throw new Error(`--${key} requires true or false`);
        params[name] = value === 'true';
      } else if (['limit', 'offset', 'scan_limit', 'body_offset', 'max_chars', 'index', 'max_bytes', 'byte_offset', 'chunk_bytes'].includes(name)) params[name] = Number(value);
      else params[name] = value;
    }
  }
  validateCall(operation, params);
  if (command === 'download' && !outputPath) throw new Error('download requires --out');
  if (command !== 'download' && outputPath) throw new Error('--out is only supported for download');
  if (command === 'download') {
    if (params.byte_offset) throw new Error('download saves a complete attachment; byte-offset must be zero');
    const destination = path.resolve(outputPath);
    const initialState = await readState({ requireMacProtectedRoots: process.platform === 'darwin' });
    let writeDestination = destination;
    if (process.platform === 'darwin') {
      writeDestination = await assertSafeMacPath(destination, { protectedRoots: initialState.protected_mail_roots });
    }
    const requestOptions = { state: initialState };
    const result = await call(operation, { ...params, chunk_bytes: 1048576 }, requestOptions);
    const chunks = [Buffer.from(result.data, 'base64')];
    let next = result.next_offset;
    while (next !== null) {
      const part = await call(operation, { ...params, chunk_bytes: 1048576, byte_offset: next }, requestOptions);
      if (part.sha256 !== result.sha256 || part.size_bytes !== result.size_bytes || part.offset !== next || (part.next_offset !== null && part.next_offset <= next)) throw new Error('Attachment changed during download');
      chunks.push(Buffer.from(part.data, 'base64'));
      next = part.next_offset;
    }
    const bytes = Buffer.concat(chunks);
    const sha256 = createHash('sha256').update(bytes).digest('hex');
    if (sha256 !== result.sha256 || bytes.length !== result.size_bytes) throw new Error('Attachment integrity verification failed');
    const currentState = await readState({ requireMacProtectedRoots: process.platform === 'darwin' });
    if (!sameDownloadState(initialState, currentState)) throw downloadStateChanged();
    if (process.platform === 'darwin') {
      const initialDestination = await assertSafeMacPath(destination, { protectedRoots: initialState.protected_mail_roots });
      if (initialDestination !== writeDestination) throw downloadStateChanged();
      const currentDestination = await assertSafeMacPath(destination, { protectedRoots: currentState.protected_mail_roots });
      if (currentDestination !== writeDestination) throw downloadStateChanged();
    }
    await writeFile(writeDestination, bytes, { flag: 'wx', mode: 0o600 });
    return print({ path: writeDestination, size_bytes: bytes.length, sha256 });
  }
  print(await call(operation, params));
}

main().catch(error => { process.stderr.write(JSON.stringify({ error: { code: error.code ?? 'CLI_ERROR', message: error.message } }) + '\n'); process.exitCode = 1; });
