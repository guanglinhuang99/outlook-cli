import { constants, realpathSync } from 'node:fs';
import { chmod, lstat, mkdir, open, readdir, rename, unlink } from 'node:fs/promises';
import { randomBytes } from 'node:crypto';
import path from 'node:path';
import os from 'node:os';
import { fileURLToPath } from 'node:url';
import { execFileSync } from 'node:child_process';
import { assertSafeMacPath } from './mac-paths.mjs';

export const root = fileURLToPath(new URL('..', import.meta.url));
export const runtimeMarkerName = '.outlook-local-runtime';
const runtimeMarkerContent = 'outlook-local-runtime-v1\n';

function runtimePathError(message) {
  return Object.assign(new Error(message), { code: 'INVALID_ARGUMENT' });
}

function isWithin(target, parent) {
  const relative = path.relative(parent, target);
  return relative === '' || (relative !== '..' && !relative.startsWith(`..${path.sep}`) && !path.isAbsolute(relative));
}

function canonicalPath(value) {
  let current = path.resolve(value);
  const tail = [];
  while (true) {
    try {
      return path.join(realpathSync(current), ...tail);
    } catch (error) {
      if (error?.code !== 'ENOENT' && error?.code !== 'ENOTDIR') throw error;
      const parent = path.dirname(current);
      if (parent === current) return path.join(current, ...tail);
      tail.unshift(path.basename(current));
      current = parent;
    }
  }
}

function resolveRuntimeDirectory(value) {
  if (typeof value !== 'string' || value.length === 0 || value.includes('\0')) {
    throw runtimePathError('OUTLOOK_LOCAL_STATE_DIR must be a non-empty absolute path');
  }
  if (!path.isAbsolute(value)) {
    throw runtimePathError('OUTLOOK_LOCAL_STATE_DIR must be an absolute path');
  }
  let resolved;
  try {
    resolved = canonicalPath(value);
  } catch (error) {
    throw runtimePathError(`OUTLOOK_LOCAL_STATE_DIR is not a safe path: ${error.message}`);
  }
  const filesystemRoot = path.parse(resolved).root;
  const home = canonicalPath(os.homedir());
  const cwd = canonicalPath(process.cwd());
  if (resolved === filesystemRoot || isWithin(home, resolved) || isWithin(cwd, resolved)) {
    throw runtimePathError('OUTLOOK_LOCAL_STATE_DIR must not be the filesystem root, home, current directory, or an ancestor of either');
  }
  return resolved;
}

function trustedPosixDirectory(stats) {
  if (!stats.isDirectory() || stats.isSymbolicLink() || !Number.isInteger(stats.uid)) return false;
  const uid = process.getuid();
  if (stats.uid !== 0 && stats.uid !== uid) return false;
  return (stats.mode & 0o022) === 0 || (stats.mode & 0o1000) !== 0;
}

async function verifyPosixParentChain(directory) {
  if (process.platform === 'win32') return;
  let current = path.dirname(directory);
  while (true) {
    let stats;
    try {
      stats = await lstat(current);
    } catch (error) {
      if (error?.code === 'ENOENT') {
        const parent = path.dirname(current);
        if (parent === current) return;
        current = parent;
        continue;
      }
      throw error;
    }
    if (!trustedPosixDirectory(stats)) throw runtimePathError('runtime state directory has an untrusted parent directory');
    const parent = path.dirname(current);
    if (parent === current) return;
    current = parent;
  }
}

function verifyPosixLeafOwner(stats) {
  if (process.platform !== 'win32' && stats.uid !== process.getuid()) {
    throw runtimePathError('runtime state directory must be owned by the current user');
  }
}

const defaultRuntimeDir = process.platform === 'darwin'
  ? path.join(os.homedir(), 'Library/Caches/outlook-cli/service') : path.join(root, '.runtime/service');
export const runtimeDir = resolveRuntimeDirectory(process.env.OUTLOOK_LOCAL_STATE_DIR ?? defaultRuntimeDir);
export const statePath = path.join(runtimeDir, 'service.json');
export const runtimeMarkerPath = path.join(runtimeDir, runtimeMarkerName);
export const port = Number(process.env.OUTLOOK_LOCAL_PORT ?? 17379);
if (!Number.isInteger(port) || port < 1024 || port > 65535) throw new Error('OUTLOOK_LOCAL_PORT must be 1024..65535');

const noFollow = process.platform === 'win32' ? 0 : (constants.O_NOFOLLOW ?? 0);
const guardedReadFlags = constants.O_RDONLY
  | noFollow
  | (process.platform === 'win32' ? 0 : (constants.O_NONBLOCK ?? 0));
const legacyStateMaxBytes = 64 * 1024;

async function markerState(directory) {
  const marker = path.join(directory, runtimeMarkerName);
  if (process.platform === 'win32') {
    try {
      if ((await lstat(marker)).isSymbolicLink()) throw runtimePathError('runtime marker must not be a symbolic link');
    } catch (error) {
      if (error?.code === 'ENOENT') return false;
      throw error;
    }
  }
  let handle;
  try {
    handle = await open(marker, guardedReadFlags);
    const stats = await handle.stat();
    if (!stats.isFile() || stats.nlink !== 1 || stats.size > 256) throw runtimePathError('runtime marker must be a single-link regular file');
    const content = await handle.readFile('utf8');
    if (content !== runtimeMarkerContent) throw runtimePathError('runtime marker is not recognized');
    if (process.platform !== 'win32') await handle.chmod(0o600);
    return true;
  } catch (error) {
    if (error?.code === 'ENOENT') return false;
    if (error?.code === 'ELOOP') throw runtimePathError('runtime marker must not be a symbolic link');
    throw error;
  } finally {
    if (handle) await handle.close().catch(() => {});
  }
}

async function createMarker(directory) {
  const marker = path.join(directory, runtimeMarkerName);
  const flags = constants.O_WRONLY | constants.O_CREAT | constants.O_EXCL | noFollow;
  let handle;
  let created = false;
  let published = false;
  try {
    handle = await open(marker, flags, 0o600);
    created = true;
    const stats = await handle.stat();
    if (!stats.isFile() || stats.nlink !== 1) throw runtimePathError('runtime marker must be a single-link regular file');
    await handle.writeFile(runtimeMarkerContent, 'utf8');
    if (process.platform !== 'win32') await handle.chmod(0o600);
    await handle.close();
    handle = undefined;
    published = true;
  } catch (error) {
    if (error?.code === 'EEXIST') {
      if (await markerState(directory)) return;
      throw runtimePathError('runtime marker already exists but is not recognized');
    }
    throw error;
  } finally {
    if (handle) await handle.close().catch(() => {});
    if (created && !published) await unlink(marker).catch(() => {});
  }
}

async function singleRegularFile(file, label) {
  let stats;
  try {
    stats = await lstat(file);
  } catch {
    return false;
  }
  if (!stats.isFile() || stats.nlink !== 1) throw runtimePathError(`${label} must be a single-link regular file`);
  return true;
}

async function validLegacyReadyLog(file) {
  if (process.platform === 'win32') {
    try {
      if ((await lstat(file)).isSymbolicLink()) return false;
    } catch (error) {
      if (error?.code === 'ENOENT') return false;
      throw error;
    }
  }
  let handle;
  try {
    handle = await open(file, guardedReadFlags);
    const stats = await handle.stat();
    if (!stats.isFile() || stats.nlink !== 1) return false;
    const buffer = Buffer.alloc(256);
    const { bytesRead } = await handle.read(buffer, 0, buffer.length, 0);
    const firstBytes = buffer.subarray(0, bytesRead).toString('utf8');
    const firstLine = firstBytes.split(/\r?\n/, 1)[0];
    const hasBoundary = firstBytes.includes('\n') || (stats.size <= buffer.length && bytesRead === stats.size);
    return hasBoundary && /^Outlook local service ready; pid=[1-9]\d*$/.test(firstLine);
  } catch (error) {
    if (error?.code === 'ENOENT' || error?.code === 'ELOOP') return false;
    throw error;
  } finally {
    if (handle) await handle.close().catch(() => {});
  }
}

function validLegacyState(state) {
  if (!state || typeof state !== 'object' || Array.isArray(state)) return false;
  const allowed = new Set(['port', 'token', 'pid', 'startedAt', 'protected_mail_roots']);
  const required = ['port', 'token', 'pid', 'startedAt'];
  if (Object.keys(state).some(key => !allowed.has(key)) || required.some(key => !Object.hasOwn(state, key))) return false;
  if (!Number.isInteger(state.pid) || state.pid <= 0 || typeof state.startedAt !== 'string') return false;
  const parsed = Date.parse(state.startedAt);
  if (!Number.isFinite(parsed) || new Date(parsed).toISOString() !== state.startedAt) return false;
  try {
    validateState(state, false);
    return true;
  } catch {
    return false;
  }
}

async function readStateFile(file) {
  if (process.platform === 'win32') {
    try {
      if ((await lstat(file)).isSymbolicLink()) throw runtimePathError('service state must not be a symbolic link');
    } catch (error) {
      if (error?.code === 'ENOENT') throw error;
      throw error;
    }
  }
  let handle;
  try {
    handle = await open(file, guardedReadFlags);
    const stats = await handle.stat();
    if (!stats.isFile() || stats.nlink !== 1 || !Number.isSafeInteger(stats.size) || stats.size > legacyStateMaxBytes) {
      throw runtimePathError('service state must be a single-link regular file within the size limit');
    }
    const buffer = Buffer.alloc(stats.size);
    let bytesRead = 0;
    while (bytesRead < buffer.length) {
      const result = await handle.read(buffer, bytesRead, buffer.length - bytesRead, bytesRead);
      if (!result.bytesRead) break;
      bytesRead += result.bytesRead;
    }
    const after = await handle.stat();
    if (after.nlink !== 1 || after.size !== stats.size || bytesRead !== stats.size) {
      throw runtimePathError('service state changed while reading');
    }
    return JSON.parse(buffer.toString('utf8'));
  } finally {
    if (handle) await handle.close().catch(() => {});
  }
}

async function validLegacyStateFile(file) {
  try {
    return validLegacyState(await readStateFile(file));
  } catch {
    return false;
  }
}

async function validLegacyRuntime(directory) {
  let names;
  try {
    names = await readdir(directory);
  } catch {
    return false;
  }
  const allowed = new Set(['service.json', 'service.log']);
  if (!names.length || names.some(name => !allowed.has(name))) return false;
  if (!names.includes('service.json')) {
    return names.length === 1 && await validLegacyReadyLog(path.join(directory, 'service.log'));
  }
  if (!await validLegacyStateFile(path.join(directory, 'service.json'))) return false;
  if (names.includes('service.log') && !await singleRegularFile(path.join(directory, 'service.log'), 'runtime log')) return false;
  return true;
}

async function applyWindowsAcl(directory) {
  // Protect the bearer token and logs from other local Windows accounts.
  const sid = execFileSync('whoami.exe', ['/user', '/fo', 'csv', '/nh'], { encoding: 'utf8', windowsHide: true }).match(/S-1-[0-9-]+/)?.[0];
  if (!sid) throw new Error('Cannot resolve current Windows SID');
  execFileSync('icacls.exe', [directory, '/inheritance:r', '/grant:r', `*${sid}:(OI)(CI)F`, '*S-1-5-18:(OI)(CI)F'], { stdio: 'pipe', windowsHide: true });
}

export async function prepareRuntime(directory = runtimeDir) {
  const safeDirectory = resolveRuntimeDirectory(directory);
  if (process.platform === 'darwin') await assertSafeMacPath(safeDirectory);
  await verifyPosixParentChain(safeDirectory);
  let created = false;
  try {
    await mkdir(path.dirname(safeDirectory), { recursive: true, mode: 0o700 });
    await mkdir(safeDirectory, { mode: 0o700 });
    created = true;
  } catch (error) {
    if (error?.code !== 'EEXIST') throw error;
  }
  const directoryStats = await lstat(safeDirectory);
  if (!directoryStats.isDirectory() || directoryStats.isSymbolicLink()) throw runtimePathError('runtime state directory must be a real directory');
  verifyPosixLeafOwner(directoryStats);
  await verifyPosixParentChain(safeDirectory);
  if (created) {
    await createMarker(safeDirectory);
  } else if (!await markerState(safeDirectory)) {
    if (!await validLegacyRuntime(safeDirectory)) throw runtimePathError('runtime state directory is not managed by outlook-local service');
    await createMarker(safeDirectory);
  }
  if (process.platform !== 'win32') await chmod(safeDirectory, 0o700);
  if (process.platform === 'win32') await applyWindowsAcl(safeDirectory);
}

const logFlags = constants.O_WRONLY | constants.O_APPEND | constants.O_CREAT
  | (process.platform === 'win32' ? 0 : (constants.O_NONBLOCK ?? 0))
  | (process.platform === 'win32' ? 0 : (constants.O_NOFOLLOW ?? 0));

export async function openRuntimeLog(logPath = path.join(runtimeDir, 'service.log')) {
  if (process.platform === 'win32') {
    try {
      if ((await lstat(logPath)).isSymbolicLink()) throw new Error('runtime log must not be a symbolic link');
    } catch (error) {
      if (error?.code !== 'ENOENT') throw error;
    }
  }
  const handle = await open(logPath, logFlags, 0o600);
  try {
    const stats = await handle.stat();
    if (!stats.isFile() || stats.nlink !== 1) throw new Error('runtime log must be a single-link regular file');
    await handle.chmod(0o600);
    return handle;
  } catch (error) {
    await handle.close().catch(() => {});
    throw error;
  }
}

export async function publishState(value, target = statePath) {
  const temporary = path.join(path.dirname(target), `.${path.basename(target)}.${process.pid}.${randomBytes(12).toString('hex')}.tmp`);
  let handle;
  try {
    handle = await open(temporary, 'wx', 0o600);
    await handle.writeFile(JSON.stringify(value), 'utf8');
    await handle.chmod(0o600);
    await handle.close();
    handle = undefined;
    await rename(temporary, target);
  } finally {
    if (handle) await handle.close().catch(() => {});
    await unlink(temporary).catch(() => {});
  }
}

function validateState(state, requireMacProtectedRoots) {
  if (!state || typeof state !== 'object' || Array.isArray(state)
      || !Number.isInteger(state.port) || state.port < 1024 || state.port > 65535
      || typeof state.token !== 'string' || !/^[a-f0-9]{64}$/.test(state.token)) {
    throw new Error('Invalid service state');
  }
  if (state.protected_mail_roots !== undefined
      && (!Array.isArray(state.protected_mail_roots) || state.protected_mail_roots.length === 0
          || state.protected_mail_roots.some(root => (
            typeof root !== 'string' || root.includes('\0') || !path.isAbsolute(root)
          )))) {
    throw new Error('Invalid protected macOS mail roots in service state');
  }
  if (process.platform === 'darwin' && requireMacProtectedRoots && state.protected_mail_roots === undefined) {
    throw Object.assign(new Error('Service state lacks protected macOS mail roots; restart service'), {
      code: 'SERVICE_RESTART_REQUIRED'
    });
  }
  return state;
}

export async function readState({ target = statePath, requireMacProtectedRoots = false } = {}) {
  try {
    if (process.platform !== 'win32') {
      const directory = canonicalPath(path.dirname(target));
      const directoryStats = await lstat(directory);
      verifyPosixLeafOwner(directoryStats);
      if (!trustedPosixDirectory(directoryStats)) {
        throw runtimePathError('runtime state directory is not trusted');
      }
      await verifyPosixParentChain(directory);
    }
    return validateState(await readStateFile(target), requireMacProtectedRoots);
  } catch (cause) {
    const restartRequired = cause?.code === 'SERVICE_RESTART_REQUIRED';
    throw Object.assign(new Error(restartRequired ? cause.message : 'Service state unavailable. Run: outlook-local start', { cause }), {
      code: restartRequired ? cause.code : 'SERVICE_UNAVAILABLE'
    });
  }
}
export async function request(endpoint, body, { timeoutMs = process.platform === 'darwin' ? 185000 : 55000, state: pinnedState } = {}) {
  const state = pinnedState ?? await readState();
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
export async function call(method, params = {}, options = {}) {
  return (await request('/rpc', { method, params }, options)).result;
}
