import { randomUUID } from 'node:crypto';
import { constants } from 'node:fs';
import { lstat, open, rename, unlink, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { setTimeout as delay } from 'node:timers/promises';
import { prepareRuntime, runtimeDir } from './runtime.mjs';

const authOrigin = 'https://login.microsoftonline.com';
const uuid = /^[0-9a-f]{8}-(?:[0-9a-f]{4}-){3}[0-9a-f]{12}$/i;
const tenantName = /^(?:common|organizations|consumers|[a-z0-9]+(?:[a-z0-9-]*[a-z0-9])?(?:\.[a-z0-9]+(?:[a-z0-9-]*[a-z0-9])?)+)$/i;
const tokenPattern = /^[\x21-\x7e]{1,65536}$/;
const defaultDirectory = path.join(runtimeDir, 'graph-auth');
const fileLimit = 192 * 1024;
const guardedRead = constants.O_RDONLY | (constants.O_NOFOLLOW ?? 0) | (constants.O_NONBLOCK ?? 0);

const validToken = value => typeof value === 'string' && tokenPattern.test(value);

function authError(code, message) {
  return Object.assign(new Error(message), { code });
}

function validateApp(client, tenant) {
  if (typeof client !== 'string' || !uuid.test(client)
      || typeof tenant !== 'string' || tenant.length > 253 || !(uuid.test(tenant) || tenantName.test(tenant))) {
    throw authError('GRAPH_APP_REQUIRED', 'Set OUTLOOK_GRAPH_CLIENT_ID and an approved OUTLOOK_GRAPH_TENANT_ID');
  }
}

async function optionalJson(directory, name) {
  const file = path.join(directory, name);
  let handle;
  try {
    const stat = await lstat(file);
    if (!stat.isFile() || stat.nlink !== 1 || stat.size > fileLimit || (process.getuid && (
      stat.uid !== process.getuid() || (stat.mode & 0o077) !== 0
    ))) throw authError('GRAPH_AUTH_STATE_INVALID', 'Graph authorization state must be a private regular file');
    handle = await open(file, guardedRead);
    const opened = await handle.stat();
    if (!opened.isFile() || opened.ino !== stat.ino || opened.dev !== stat.dev || opened.size > fileLimit) {
      throw authError('GRAPH_AUTH_STATE_INVALID', 'Graph authorization state changed while reading');
    }
    try { return JSON.parse(await handle.readFile('utf8')); }
    catch { throw authError('GRAPH_AUTH_STATE_INVALID', 'Graph authorization state is invalid'); }
  } catch (error) {
    if (error?.code === 'ENOENT') return undefined;
    throw error;
  } finally {
    if (handle) await handle.close();
  }
}

async function atomicJson(directory, name, value) {
  const temp = path.join(directory, `.oauth-${randomUUID()}.tmp`);
  try {
    await writeFile(temp, JSON.stringify(value) + '\n', { flag: 'wx', mode: 0o600 });
    await rename(temp, path.join(directory, name));
  } finally {
    await unlink(temp).catch(error => { if (error?.code !== 'ENOENT') throw error; });
  }
}

async function activeGrant(directory) {
  const directoryInfo = await lstat(directory).catch(error => {
    if (error?.code === 'ENOENT') return undefined;
    throw error;
  });
  if (!directoryInfo) return undefined;
  if (!directoryInfo.isDirectory() || (process.getuid && (
    directoryInfo.uid !== process.getuid() || (directoryInfo.mode & 0o077) !== 0
  ))) throw authError('GRAPH_AUTH_STATE_INVALID', 'Graph authorization directory is unsafe');
  const current = await optionalJson(directory, 'current.json');
  if (!current) return undefined;
  if (typeof current.id !== 'string' || !uuid.test(current.id)) {
    throw authError('GRAPH_AUTH_STATE_INVALID', 'Graph authorization state is invalid');
  }
  const name = `grant-${current.id}.json`;
  const grant = await optionalJson(directory, name);
  if (!grant || !validToken(grant.accessToken) || !Number.isSafeInteger(grant.expiresAt)
      || (grant.refreshToken !== null && !validToken(grant.refreshToken))) {
    throw authError('GRAPH_AUTH_STATE_INVALID', 'Graph authorization grant is unavailable or invalid');
  }
  validateApp(grant.client, grant.tenant);
  return { id: current.id, name, grant };
}

async function identityRequest(tenant, endpoint, params, fetchImpl) {
  try {
    const response = await fetchImpl(`${authOrigin}/${tenant}/oauth2/v2.0/${endpoint}`, {
      method: 'POST', body: new URLSearchParams(params), redirect: 'error', signal: AbortSignal.timeout(20000)
    });
    return { ok: response.ok, data: await response.json() };
  } catch {
    throw authError('GRAPH_AUTH_NETWORK', 'Could not contact Microsoft Entra authorization');
  }
}

function grantFrom(data, client, tenant, previousRefresh, now) {
  if (!data || !validToken(data.access_token) || !Number.isSafeInteger(data.expires_in)
      || data.expires_in < 1 || data.expires_in > 86400
      || (data.refresh_token !== undefined && !validToken(data.refresh_token))) {
    throw authError('GRAPH_AUTH_INVALID_RESPONSE', 'Microsoft Entra returned an invalid authorization response');
  }
  return { client, tenant, accessToken: data.access_token, refreshToken: data.refresh_token ?? previousRefresh,
    expiresAt: now() + data.expires_in * 1000 };
}

export async function graphLogin(params = {}, {
  env = process.env, directory = defaultDirectory, fetchImpl = globalThis.fetch,
  now = Date.now, sleep = delay, notify = ({ url, code }) => process.stderr.write(`Sign in at ${url} with code ${code}\n`)
} = {}) {
  if (Object.keys(params).some(key => !['read_only', 'send'].includes(key))
      || (params.read_only !== undefined && typeof params.read_only !== 'boolean')
      || (params.send !== undefined && typeof params.send !== 'boolean')
      || (params.read_only && params.send)) throw authError('INVALID_ARGUMENT', 'Invalid Graph login options');
  const client = env.OUTLOOK_GRAPH_CLIENT_ID;
  const tenant = env.OUTLOOK_GRAPH_TENANT_ID ?? 'organizations';
  validateApp(client, tenant);
  await prepareRuntime(directory);
  const scope = ['offline_access', 'https://graph.microsoft.com/User.Read',
    `https://graph.microsoft.com/${params.read_only ? 'Mail.Read' : 'Mail.ReadWrite'}`];
  if (params.send) scope.push('https://graph.microsoft.com/Mail.Send');
  const device = await identityRequest(tenant, 'devicecode', { client_id: client, scope: scope.join(' ') }, fetchImpl);
  if (!device.ok) throw authError('GRAPH_AUTH_DENIED', 'Microsoft Entra declined to initiate device sign-in');
  const { device_code: deviceCode, user_code: code, expires_in: expiresIn, interval } = device.data ?? {};
  if (!validToken(deviceCode) || typeof code !== 'string' || !/^[A-Z0-9-]{4,24}$/i.test(code)
      || !Number.isSafeInteger(expiresIn) || expiresIn < 1 || expiresIn > 1800
      || !Number.isSafeInteger(interval) || interval < 1 || interval > 60) {
    throw authError('GRAPH_AUTH_INVALID_RESPONSE', 'Microsoft Entra returned an invalid device sign-in response');
  }
  notify({ url: 'https://microsoft.com/devicelogin', code });
  const deadline = now() + expiresIn * 1000;
  let pollingSeconds = interval;
  while (now() + pollingSeconds * 1000 < deadline) {
    await sleep(pollingSeconds * 1000);
    const result = await identityRequest(tenant, 'token', {
      grant_type: 'urn:ietf:params:oauth:grant-type:device_code', client_id: client, device_code: deviceCode
    }, fetchImpl);
    if (result.ok) {
      const previous = await activeGrant(directory);
      const grant = grantFrom(result.data, client, tenant, null, now);
      const id = randomUUID();
      await atomicJson(directory, `grant-${id}.json`, grant);
      await atomicJson(directory, 'current.json', { id });
      if (previous) await unlink(path.join(directory, previous.name)).catch(error => { if (error?.code !== 'ENOENT') throw error; });
      return { authorized: true, expires_at: new Date(grant.expiresAt).toISOString(), can_refresh: Boolean(grant.refreshToken) };
    }
    if (result.data?.error === 'slow_down') pollingSeconds += 5;
    else if (result.data?.error !== 'authorization_pending') {
      throw authError('GRAPH_AUTH_DENIED', 'Microsoft Entra device sign-in was declined or expired');
    }
  }
  throw authError('GRAPH_AUTH_EXPIRED', 'Microsoft Entra device sign-in expired');
}

export async function graphAccessToken({ env = process.env, directory = defaultDirectory,
  fetchImpl = globalThis.fetch, now = Date.now, refresh = true } = {}) {
  if (env.OUTLOOK_GRAPH_ACCESS_TOKEN) return env.OUTLOOK_GRAPH_ACCESS_TOKEN;
  const current = await activeGrant(directory);
  if (!current) return undefined;
  const grant = current.grant;
  if ((env.OUTLOOK_GRAPH_CLIENT_ID && env.OUTLOOK_GRAPH_CLIENT_ID.toLowerCase() !== grant.client.toLowerCase())
      || (env.OUTLOOK_GRAPH_TENANT_ID && env.OUTLOOK_GRAPH_TENANT_ID.toLowerCase() !== grant.tenant.toLowerCase())) {
    throw authError('GRAPH_AUTH_APP_CHANGED', 'The cached Graph authorization belongs to another app or tenant');
  }
  if (grant.expiresAt > now() + 60000) return grant.accessToken;
  if (!refresh) return undefined;
  if (!grant.refreshToken) throw authError('GRAPH_AUTH_EXPIRED', 'Run graph login again');
  const result = await identityRequest(grant.tenant, 'token', {
    grant_type: 'refresh_token', client_id: grant.client, refresh_token: grant.refreshToken
  }, fetchImpl);
  if (!result.ok) throw authError('GRAPH_AUTH_EXPIRED', 'Microsoft Entra declined to refresh; run graph login again');
  const updated = grantFrom(result.data, grant.client, grant.tenant, grant.refreshToken, now);
  const latest = await activeGrant(directory);
  if (!latest || latest.id !== current.id) throw authError('GRAPH_AUTH_CHANGED', 'Graph authorization changed during refresh');
  await atomicJson(directory, current.name, updated);
  return updated.accessToken;
}

export async function graphLogout({ directory = defaultDirectory } = {}) {
  const current = await activeGrant(directory);
  if (!current) return { logged_out: true, removed_local_grant: false };
  await unlink(path.join(directory, 'current.json'));
  await unlink(path.join(directory, current.name)).catch(error => { if (error?.code !== 'ENOENT') throw error; });
  return { logged_out: true, removed_local_grant: true };
}
