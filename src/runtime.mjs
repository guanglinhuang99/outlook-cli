import { mkdir, readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { execFileSync } from 'node:child_process';

export const root = fileURLToPath(new URL('..', import.meta.url));
export const runtimeDir = path.resolve(process.env.OUTLOOK_LOCAL_STATE_DIR ?? path.join(root, '.runtime/service'));
export const statePath = path.join(runtimeDir, 'service.json');
export const port = Number(process.env.OUTLOOK_LOCAL_PORT ?? 17379);
if (!Number.isInteger(port) || port < 1024 || port > 65535) throw new Error('OUTLOOK_LOCAL_PORT must be 1024..65535');

export async function prepareRuntime() {
  await mkdir(runtimeDir, { recursive: true, mode: 0o700 });
  if (process.platform === 'win32') {
    // Protect the bearer token and logs from other local Windows accounts.
    const sid = execFileSync('whoami.exe', ['/user', '/fo', 'csv', '/nh'], { encoding: 'utf8', windowsHide: true }).match(/S-1-[0-9-]+/)?.[0];
    if (!sid) throw new Error('Cannot resolve current Windows SID');
    execFileSync('icacls.exe', [runtimeDir, '/inheritance:r', '/grant:r', `*${sid}:(OI)(CI)F`, '*S-1-5-18:(OI)(CI)F'], { stdio: 'pipe', windowsHide: true });
  }
}
export async function readState() {
  try {
    const state = JSON.parse(await readFile(statePath, 'utf8'));
    if (!Number.isInteger(state.port) || state.port < 1024 || state.port > 65535 || !/^[a-f0-9]{64}$/.test(state.token)) throw new Error('Invalid service state');
    return state;
  } catch (cause) {
    throw Object.assign(new Error('Service state unavailable. Run: outlook-local start', { cause }), { code: 'SERVICE_UNAVAILABLE' });
  }
}
export async function request(endpoint, body, { timeoutMs = 55000 } = {}) {
  const state = await readState();
  let response;
  try {
    response = await fetch(`http://127.0.0.1:${state.port}${endpoint}`, {
      method: body === undefined ? 'GET' : 'POST',
      headers: { Authorization: `Bearer ${state.token}`, 'Content-Type': 'application/json' },
      ...(body === undefined ? {} : { body: JSON.stringify(body) }),
      signal: AbortSignal.timeout(timeoutMs)
    });
  } catch (cause) {
    throw Object.assign(new Error('Local service is unavailable or timed out. Run: outlook-local start', { cause }), { code: 'SERVICE_UNAVAILABLE' });
  }
  const data = await response.json();
  if (!response.ok || data.error) throw Object.assign(new Error(data.error?.message ?? `HTTP ${response.status}`), { code: data.error?.code ?? 'HTTP_ERROR' });
  return data;
}
export async function call(method, params = {}) { return (await request('/rpc', { method, params })).result; }
