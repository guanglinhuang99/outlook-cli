import { readdir, realpath, stat } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';

export const MAC_PATH_FORBIDDEN = 'MAC_PATH_FORBIDDEN';

// Resolve the nearest existing ancestor first. This keeps symlink handling in
// place for destinations whose final file (or directory) does not exist yet.
async function resolveForWrite(value) {
  let current = path.resolve(value);
  const tail = [];

  while (true) {
    try {
      const existing = await realpath(current);
      return path.join(existing, ...tail);
    } catch (error) {
      if (error?.code !== 'ENOENT' && error?.code !== 'ENOTDIR') throw error;
      const parent = path.dirname(current);
      if (parent === current) return path.join(current, ...tail);
      tail.unshift(path.basename(current));
      current = parent;
    }
  }
}

function isWithin(target, root) {
  const relative = path.relative(root, target);
  return relative === '' || (relative !== '..' && !relative.startsWith(`..${path.sep}`) && !path.isAbsolute(relative));
}

function protectedPathError() {
  return Object.assign(new Error('Path is inside a protected macOS mail cache directory'), {
    code: MAC_PATH_FORBIDDEN
  });
}

function profileRoot(profileDir, homeDir) {
  if (typeof profileDir !== 'string') {
    throw Object.assign(new Error('OUTLOOK_MAC_PROFILE must be an absolute path or start with ~/'), { code: 'INVALID_ARGUMENT' });
  }
  if (profileDir === '~' || profileDir.startsWith('~/')) return path.join(homeDir, profileDir.slice(2));
  if (path.isAbsolute(profileDir)) return profileDir;
  throw Object.assign(new Error('OUTLOOK_MAC_PROFILE must be an absolute path or start with ~/'), { code: 'INVALID_ARGUMENT' });
}

async function discoverMacProfile(homeDir) {
  const root = path.join(homeDir, 'Library', 'Group Containers', 'UBF8T346G9.Office', 'Outlook');
  let profileGroups;
  try {
    profileGroups = await readdir(root, { withFileTypes: true });
  } catch (error) {
    if (error?.code === 'ENOENT') return undefined;
    throw error;
  }
  const matches = [];
  for (const group of profileGroups.filter(entry => entry.name.endsWith(' Profiles'))) {
    let profiles;
    try {
      profiles = await readdir(path.join(root, group.name), { withFileTypes: true });
    } catch (error) {
      if (error?.code === 'ENOENT' || error?.code === 'ENOTDIR') continue;
      throw error;
    }
    for (const profile of profiles) {
      const candidate = path.join(root, group.name, profile.name, 'HxStore.hxd');
      try {
        if ((await stat(candidate)).isFile()) matches.push(path.dirname(candidate));
      } catch (error) {
        if (error?.code !== 'ENOENT' && error?.code !== 'ENOTDIR') throw error;
      }
    }
  }
  return matches.length === 1 ? resolveForWrite(matches[0]) : undefined;
}

function validateProtectedRoots(protectedRoots) {
  if (!Array.isArray(protectedRoots) || protectedRoots.some(root => (
    typeof root !== 'string' || root.includes('\0') || !path.isAbsolute(root)
  ))) {
    throw Object.assign(new Error('protected_mail_roots must be an array of absolute paths'), { code: 'INVALID_ARGUMENT' });
  }
}

function profileUnavailableError() {
  return Object.assign(new Error('Could not resolve a unique macOS Outlook profile; set OUTLOOK_MAC_PROFILE to an absolute or ~/ path and restart the service'), {
    code: 'MAC_PROFILE_UNAVAILABLE'
  });
}

export async function canonicalMacProtectedRoots({
  homeDir = os.homedir(),
  profileDir = process.env.OUTLOOK_MAC_PROFILE,
  protectedRoots = [],
  autoDiscover = true
} = {}) {
  validateProtectedRoots(protectedRoots);
  const roots = [
    path.join(homeDir, 'Library', 'Mail'),
    path.join(homeDir, 'Library', 'Group Containers', 'UBF8T346G9.Office', 'Outlook'),
    ...protectedRoots
  ];
  const resolvedProfile = profileDir
    ? await resolveForWrite(profileRoot(profileDir, homeDir))
    : profileDir === null || !autoDiscover ? undefined : await discoverMacProfile(homeDir);
  if (resolvedProfile) roots.push(resolvedProfile);
  return [...new Set(await Promise.all(roots.map(resolveForWrite)))];
}

export async function resolveMacProfile({
  homeDir = os.homedir(),
  profileDir = process.env.OUTLOOK_MAC_PROFILE,
  strict = false
} = {}) {
  const resolvedProfile = profileDir
    ? await resolveForWrite(profileRoot(profileDir, homeDir))
    : profileDir === null ? undefined : await discoverMacProfile(homeDir);
  if (strict) {
    let valid = Boolean(resolvedProfile);
    if (valid) {
      try {
        valid = (await stat(path.join(resolvedProfile, 'HxStore.hxd'))).isFile();
      } catch (error) {
        if (error?.code !== 'ENOENT' && error?.code !== 'ENOTDIR') throw error;
        valid = false;
      }
    }
    if (!valid) throw profileUnavailableError();
  }
  return resolvedProfile;
}

/**
 * Resolve a path for writing and reject the read-only macOS mail cache roots.
 *
 * `homeDir` and `profileDir` are injectable so tests do not need to mutate
 * HOME. Callers on macOS should use the returned path for the actual write.
 */
export async function assertSafeMacPath(target, {
  homeDir = os.homedir(),
  profileDir = process.env.OUTLOOK_MAC_PROFILE,
  protectedRoots = []
} = {}) {
  const canonicalTarget = await resolveForWrite(target);
  const canonicalRoots = await canonicalMacProtectedRoots({ homeDir, profileDir, protectedRoots });
  if (canonicalRoots.some(root => isWithin(canonicalTarget, root))) throw protectedPathError();
  return canonicalTarget;
}
