#!/usr/bin/env node
import { spawn } from 'node:child_process';
import { constants } from 'node:fs';
import { open, writeFile } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { parseArgs } from 'node:util';
import { setTimeout as delay } from 'node:timers/promises';
import { openRuntimeLog, prepareRuntime, runtimeDir, root, statePath, readState, request, call } from '../src/runtime.mjs';
import { assertSafeMacPath } from '../src/mac-paths.mjs';
import { validateCall, definitions, isMacWorkflow } from '../src/tools.mjs';
import { nodeVersionError } from '../src/node-version.mjs';

const print = value => process.stdout.write(JSON.stringify(value, null, 2) + '\n');
const inputLimit = 4 * 1024 * 1024 + 32768;
const oversizedInput = () => new Error(`Workflow JSON exceeds ${inputLimit} bytes`);

async function readWorkflowInput(file) {
  const chunks = [];
  let size = 0;
  const handle = file === '-' ? null : await open(file, constants.O_RDONLY | (constants.O_NONBLOCK ?? 0));
  try {
    if (handle) {
      const info = await handle.stat();
      if (!info.isFile()) throw new Error('--json-file requires a regular file or -');
      if (info.size > inputLimit) throw oversizedInput();
    }
    const input = handle ? handle.createReadStream({ autoClose: false }) : process.stdin;
    for await (const chunk of input) {
      size += chunk.length;
      if (size > inputLimit) throw oversizedInput();
      chunks.push(chunk);
    }
    return Buffer.concat(chunks).toString('utf8');
  } finally {
    if (handle) await handle.close();
  }
}

async function workflowArguments(args) {
  const { values } = parseArgs({ args, options: { json: { type: 'string' }, 'json-file': { type: 'string' } },
    strict: true, allowPositionals: false });
  if (values.json !== undefined && values['json-file'] !== undefined) {
    throw new Error('--json and --json-file are mutually exclusive');
  }
  const source = values['json-file'] !== undefined ? await readWorkflowInput(values['json-file']) : (values.json ?? '{}');
  if (Buffer.byteLength(source) > inputLimit) throw oversizedInput();
  let params;
  try { params = JSON.parse(source); }
  catch { throw Object.assign(new Error('--json/--json-file must be a valid JSON object'), { code: 'INVALID_ARGUMENT' }); }
  if (!params || typeof params !== 'object' || Array.isArray(params)) {
    throw Object.assign(new Error('--json/--json-file must be a valid JSON object'), { code: 'INVALID_ARGUMENT' });
  }
  return params;
}

async function attachmentCommand(command, args) {
  const options = { id: { type: 'string' }, backend: { type: 'string' } };
  if (command === 'download-all') options['out-dir'] = { type: 'string' };
  else {
    options.index = { type: 'string' };
    options['max-chars'] = { type: 'string' };
  }
  const { values } = parseArgs({ args, options, strict: true, allowPositionals: false });
  const { id, backend } = validateCall('attachments', { id: values.id, backend: values.backend });
  const workflows = await import('../src/attachment-workflows.mjs');
  const service = { call: callStarting, readState: readStateStarting };
  if (command === 'download-all') {
    if (!values['out-dir']) throw new Error('download-all requires --out-dir');
    return workflows.downloadAllAttachments({ id, backend, out_dir: values['out-dir'] }, service);
  }
  const { index } = validateCall('read_attachment', { id, backend, index: Number(values.index) });
  const maximum = values['max-chars'] === undefined ? {} : { max_chars: Number(values['max-chars']) };
  return workflows.readAttachmentText({ id, backend, index, ...maximum }, service);
}

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

const help = `Outlook CLI — read-only local Outlook mail for agents

  outlook-local start | stop | restart | status | health
  outlook-local clipboard start|stop|status  Windows Outlook copy-to-MSG bridge
  outlook-local stores
  outlook-local folders [--folder inbox|<folder-id>] [--limit 25] [--offset 0]
  outlook-local messages [--folder inbox] [--limit 25] [--unread true]
  outlook-local today [--folder inbox|sent] [--limit 25] [--date YYYY-MM-DD]
  outlook-local search [--query text] [--person name-or-email] [--person-role any|sender|to|cc]
                       [--search-body true] [--com-body-mode auto|direct|indexed] [--scan-limit 500]
  outlook-local read --id <message-id> [--max-chars 20000] [--body-offset 0]
  outlook-local context-bundle [--folder inbox|sent|all|<id>] [--from-date YYYY-MM-DD] [--to-date YYYY-MM-DD]
                               [--query text] [--unread true] [--mode metadata|context] [--cursor <cursor>]
  outlook-local today-bundle [--folder inbox|sent|all|<id>] [--date YYYY-MM-DD] [--mode context|metadata]
  outlook-local messages-bundle --id <message-id> [--id <message-id>] [--max-chars 12000]
  outlook-local attachments --id <message-id>
  outlook-local download --id <message-id> --index 1 --out <file>
  outlook-local download-all --id <message-id> --out-dir <directory>
  outlook-local attachment-text --id <message-id> --index 1 [--max-chars N]
  outlook-local conversation | delta-bundle | task-bundle [--json '<JSON object>' | --json-file <file|->]
  outlook-local knowledge status|sync|search|read|source|export|bundle [--json '<JSON object>' | --json-file <file|->]
  outlook-local calendar status|calendars|agenda|authorize [--json '<JSON object>' | --json-file <file|->]
  outlook-local semantic sync|search|status [--json '<JSON object>' | --json-file <file|->]
  outlook-local mcp              MCP stdio adapter; starts the service on demand
  outlook-local mcp-config       Print MCP stdio configuration (no secrets)
  outlook-local rpc <operation> --json '<JSON object>'
  outlook-local serve            Run service in foreground

Messages/search: --after/--before ISO time with timezone, --offset,
--limit, --scan-limit. Search covers one folder per call; follow next_offset.
All mailbox commands accept --backend auto|com|local|mac-cache (default auto).
auto uses Windows COM/local OST/PST or the macOS HxStore cache by platform.
Windows auto/com may start classic Outlook on demand (OUTLOOK_AUTO_START=0 disables it).
local never starts Outlook. auto reports explicit warnings when falling back to cached mail.
Startup timeouts do not auto-fallback; check Outlook prompts or explicitly select local.
Both platforms filter by received_at; macOS uses verified LastDeliveryTime.
On macOS, Python 3.11+ is required. Unread filters use verified local state.
Set OUTLOOK_MAC_FOLDER_SCOPE=account and separate service/cache directories
to opt in to verified account and archive folder trees. HTML is unsupported.
Bundle commands are macOS-only. By default, folder queries include descendants;
use --include-subfolders false to query the selected roots or folder only.
Metadata mode allows 100 results per page. Context mode reads cached bodies
with --max-chars and --text-budget-chars and allows 25 results per page.
Local reads contain cached data only; COM and local IDs are different.
Semantic search uses a local embedding provider and the macOS cache workflows.
Use --json-file - to pass workflow JSON through stdin.
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

// After a reboot or logout the service is gone; mail commands and the MCP adapter start
// it on demand. One start runs at a time, and pinned-state calls never restart silently.
let starting;
function ensureStarted() {
  starting ??= start().finally(() => { starting = undefined; });
  return starting;
}

async function callStarting(method, params, options) {
  try { return await call(method, params, options); }
  catch (error) {
    if (error.code !== 'SERVICE_NOT_RUNNING' || options?.state) throw error;
    await ensureStarted();
    return call(method, params, options);
  }
}

async function readStateStarting(options) {
  try { return await readState(options); }
  catch (error) {
    if (error.code !== 'SERVICE_NOT_RUNNING') throw error;
    await ensureStarted();
    return readState(options);
  }
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
  const unsupportedNode = nodeVersionError();
  if (unsupportedNode) throw unsupportedNode;
  if (command === 'clipboard') {
    if (args.length !== 1) throw new Error('Usage: outlook-local clipboard start|stop|status');
    const { clipboardCommand } = await import('../src/clipboard.mjs');
    return print(await clipboardCommand(args[0]));
  }
  if (command === 'start') return print(await start());
  if (command === 'stop') return print(await stop());
  if (command === 'restart') {
    // A service that is not running (or left a stale state file) only needs a start.
    try { await stop(); } catch (error) { if (!['SERVICE_NOT_RUNNING', 'SERVICE_UNAVAILABLE'].includes(error.code)) throw error; }
    return print(await start());
  }
  if (command === 'status') return print(await request('/status'));
  if (command === 'download-all' || command === 'attachment-text') return print(await attachmentCommand(command, args));
  if (['conversation', 'delta-bundle', 'task-bundle'].includes(command)) {
    const { runReadParity } = await import('../src/read-parity.mjs');
    return print(await runReadParity(command, await workflowArguments(args), { call: callStarting, stateDir: path.join(runtimeDir, 'read-parity') }));
  }
  if (command === 'knowledge' || command === 'calendar') {
    const [action, ...options] = args;
    const actions = command === 'knowledge'
      ? ['status', 'sync', 'search', 'read', 'source', 'export', 'bundle']
      : ['status', 'calendars', 'agenda', 'authorize'];
    if (!action || action === 'help' || action === '--help') return print({ actions });
    const params = await workflowArguments(options);
    const stateDir = path.join(runtimeDir, command);
    if (command === 'calendar') {
      const { runCalendarWorkflow } = await import('../src/calendar-workflow.mjs');
      return print(await runCalendarWorkflow(action, params));
    }
    const { readAttachmentText } = await import('../src/attachment-workflows.mjs');
    const { runKnowledgeWorkflow } = await import('../src/knowledge-workflow.mjs');
    return print(await runKnowledgeWorkflow(action, params, { call: callStarting, stateDir, readAttachmentText }));
  }
  if (command === 'semantic') {
    const [action, ...options] = args;
    const { runSemanticWorkflow } = await import('../src/semantic-workflow.mjs');
    if (!action || action === 'help' || action === '--help') return print({ actions: ['status', 'sync', 'search'] });
    return print(await runSemanticWorkflow(action, await workflowArguments(options), {
      call: callStarting, stateDir: path.join(runtimeDir, 'semantic')
    }));
  }
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
    const server = createMcp(callStarting, { workflows: {
      call: callStarting, readState: readStateStarting, stateDir: path.join(runtimeDir, 'read-parity')
    } });
    await server.connect(new StdioServerTransport());
    process.stdin.once('end', () => { void server.close(); });
    return;
  }
  const alias = command.replaceAll('-', '_');
  const commandOperation = isMacWorkflow(alias) ? alias : command;
  const operation = command === 'download' ? 'read_attachment' : command === 'rpc' ? args.shift() : commandOperation;
  if (!Object.hasOwn(definitions, operation)) throw new Error(`Unknown command: ${command}`);
  if (isMacWorkflow(operation) && process.platform !== 'darwin') {
    throw Object.assign(new Error('This workflow is only supported by the macOS cache backend'), { code: 'UNSUPPORTED_PLATFORM' });
  }
  const parsed = parseArgs({ args, options: Object.fromEntries([
    'backend', 'folder', 'limit', 'offset', 'unread', 'after', 'before', 'date', 'scan-limit', 'query', 'person', 'person-role', 'search-body', 'com-body-mode', 'id', 'format', 'body-offset', 'max-chars', 'index', 'max-bytes', 'byte-offset', 'chunk-bytes', 'out', 'json',
    'mode', 'from-date', 'to-date', 'cursor', 'field', 'include-subfolders', 'has-cached-attachments', 'text-budget-chars'
  ].map(key => [key, { type: 'string', multiple: key === 'id' && operation === 'messages_bundle' }])), strict: true, allowPositionals: false });
  const values = { ...parsed.values };
  const outputPath = values.out;
  delete values.out;
  let params;
  if (values.json !== undefined) {
    if (Object.keys(values).length !== 1) throw new Error('--json cannot be combined with operation flags');
    try { params = JSON.parse(values.json); }
    catch { throw Object.assign(new Error('--json must be a valid JSON object'), { code: 'INVALID_ARGUMENT' }); }
    if (!params || typeof params !== 'object' || Array.isArray(params)) {
      throw Object.assign(new Error('--json must be a valid JSON object'), { code: 'INVALID_ARGUMENT' });
    }
  } else {
    params = {};
    for (const [key, value] of Object.entries(values)) {
      const name = key.replaceAll('-', '_');
      if (name === 'id' && operation === 'messages_bundle') params.ids = value;
      else if (['unread', 'search_body', 'include_subfolders', 'has_cached_attachments'].includes(name)) {
        if (!['true', 'false'].includes(value)) throw new Error(`--${key} requires true or false`);
        params[name] = value === 'true';
      } else if (['limit', 'offset', 'scan_limit', 'body_offset', 'max_chars', 'index', 'max_bytes', 'byte_offset', 'chunk_bytes', 'text_budget_chars'].includes(name)) params[name] = Number(value);
      else params[name] = value;
    }
  }
  validateCall(operation, params);
  if (command === 'download' && !outputPath) throw new Error('download requires --out');
  if (command !== 'download' && outputPath) throw new Error('--out is only supported for download');
  if (command === 'download') {
    if (params.byte_offset) throw new Error('download saves a complete attachment; byte-offset must be zero');
    const destination = path.resolve(outputPath);
    const initialState = await readStateStarting({ requireMacProtectedRoots: process.platform === 'darwin' });
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
    return print({ path: writeDestination, size_bytes: bytes.length, sha256,
      backend: result.backend, data_scope: result.data_scope,
      outlook_connection: result.outlook_connection, outlook_auto_start: result.outlook_auto_start,
      fallback_reason: result.fallback_reason, fallback_message: result.fallback_message,
      warnings: result.warnings });
  }
  print(await callStarting(operation, params));
}

main().catch(error => { process.stderr.write(JSON.stringify({ error: { code: error.code ?? 'CLI_ERROR', message: error.message } }) + '\n'); process.exitCode = 1; });
