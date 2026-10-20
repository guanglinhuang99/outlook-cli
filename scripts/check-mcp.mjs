// Check the installed stdio configuration without a separate MCP client dependency.
import { readFile } from 'node:fs/promises';
import { spawn } from 'node:child_process';
import { createInterface } from 'node:readline';
import { definitions } from '../src/tools.mjs';

const configPath = process.argv[2];
let child;
let lines;
let deadline;
try {
  if (!configPath) throw new Error('Usage: node check-mcp.mjs <mcp-config.json> [--health]');
  const config = JSON.parse(await readFile(configPath, 'utf8'));
  const server = config.mcpServers?.['outlook-local'];
  if (!server?.command || !Array.isArray(server.args)) throw new Error('Missing outlook-local stdio configuration.');
  child = spawn(server.command, server.args, {
    env: { ...process.env, ...server.env }, windowsHide: true,
    stdio: ['pipe', 'pipe', 'pipe']
  });
  // Never echo a child log that could contain mailbox data or credentials.
  child.stderr.resume();
  lines = createInterface({ input: child.stdout });
  const pending = new Map();
  let sequence = 0;
  let failure;
  const fail = error => {
    failure = error;
    for (const { reject } of pending.values()) reject(error);
    pending.clear();
  };
  child.once('error', fail);
  child.stdin.on('error', fail);
  child.once('exit', code => fail(new Error(`MCP process exited unexpectedly (${code}).`)));
  lines.on('line', line => {
    try {
      const response = JSON.parse(line);
      const item = pending.get(response.id);
      if (!item) return;
      pending.delete(response.id);
      if (response.error) item.reject(new Error(`MCP JSON-RPC error ${response.error.code}.`));
      else item.resolve(response.result);
    } catch { fail(new Error('MCP stdout contains invalid JSON.')); }
  });
  const timeout = process.platform === 'darwin' ? 190000 : 60000;
  deadline = setTimeout(() => fail(new Error('MCP check timed out.')), timeout);
  function send(message) { child.stdin.write(JSON.stringify({ jsonrpc: '2.0', ...message }) + '\n'); }
  function request(method, params = {}) {
    if (failure) return Promise.reject(failure);
    return new Promise((resolve, reject) => {
      const id = ++sequence;
      pending.set(id, { resolve, reject });
      send({ id, method, params });
    });
  }
  const initialized = await request('initialize', {
    protocolVersion: '2025-03-26', capabilities: {},
    clientInfo: { name: 'outlook-local-install-check', version: '1.0.0' }
  });
  if (!initialized?.capabilities?.tools || !initialized.protocolVersion) throw new Error('MCP initialization did not advertise tools.');
  send({ method: 'notifications/initialized' });
  const list = await request('tools/list');
  const expected = Object.keys(definitions).map(name => `outlook_${name}`);
  const cacheWrites = process.platform !== 'win32';
  if (list.tools?.length !== expected.length || !expected.every(name => list.tools.some(tool => (
    tool.name === name
      && tool.annotations?.readOnlyHint === !cacheWrites
      && tool.annotations?.destructiveHint === cacheWrites
  )))) {
    throw new Error('MCP tools/list does not match the expected platform-aware Outlook annotations.');
  }
  const report = { mcp: 'ready', transport: 'stdio', protocol: initialized.protocolVersion, tools: list.tools.map(tool => tool.name) };
  if (process.argv.includes('--health')) {
    const health = await request('tools/call', { name: 'outlook_health', arguments: {} });
    if (health.isError) throw new Error('MCP protocol works, but outlook_health failed. Start the service and run its health command for details.');
    if (!health.structuredContent?.connected) throw new Error('MCP health returned no connected mailbox backend.');
    report.backend = health.structuredContent.backend;
  }
  console.log(JSON.stringify(report, null, 2));
} catch (error) {
  console.error(`MCP check failed: ${error.message}`);
  process.exitCode = 1;
} finally {
  clearTimeout(deadline);
  lines?.close();
  child?.stdin.destroy();
  child?.kill();
}
