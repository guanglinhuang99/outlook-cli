import assert from 'node:assert/strict';
import { mkdtemp, mkdir, realpath, rm, symlink, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { assertSafeMacPath, canonicalMacProtectedRoots, MAC_PATH_FORBIDDEN, resolveMacProfile } from '../src/mac-paths.mjs';

async function tempFixture() {
  const base = await mkdtemp(path.join(os.tmpdir(), 'outlook-mac-paths-'));
  const homeDir = path.join(base, 'home');
  const profileDir = path.join(base, 'profile');
  await mkdir(homeDir, { recursive: true });
  await mkdir(profileDir, { recursive: true });
  return { base, homeDir, profileDir };
}

test('rejects injected protected roots and descendants without changing HOME', async t => {
  const fixture = await tempFixture();
  t.after(() => rm(fixture.base, { recursive: true, force: true }));
  const roots = [
    path.join(fixture.homeDir, 'Library', 'Mail'),
    path.join(fixture.homeDir, 'Library', 'Group Containers', 'UBF8T346G9.Office', 'Outlook'),
    fixture.profileDir
  ];
  await Promise.all(roots.map(root => mkdir(root, { recursive: true })));
  const originalHome = process.env.HOME;

  for (const root of roots) {
    for (const target of [root, path.join(root, 'nested', 'new-file')]) {
      await assert.rejects(
        assertSafeMacPath(target, { homeDir: fixture.homeDir, profileDir: fixture.profileDir }),
        error => error.code === MAC_PATH_FORBIDDEN
      );
    }
  }
  assert.equal(process.env.HOME, originalHome);
});

test('resolves symlinked ancestors before checking protected roots', async t => {
  const fixture = await tempFixture();
  t.after(() => rm(fixture.base, { recursive: true, force: true }));
  const mailRoot = path.join(fixture.homeDir, 'Library', 'Mail');
  const realMail = path.join(fixture.base, 'mail-real');
  await mkdir(path.dirname(mailRoot), { recursive: true });
  await mkdir(realMail, { recursive: true });
  await symlink(realMail, mailRoot, process.platform === 'win32' ? 'junction' : 'dir');

  const alias = path.join(fixture.base, 'mail-alias');
  await symlink(mailRoot, alias, process.platform === 'win32' ? 'junction' : 'dir');
  await assert.rejects(
    assertSafeMacPath(path.join(alias, 'future', 'message.bin'), {
      homeDir: fixture.homeDir,
      profileDir: fixture.profileDir
    }),
    error => error.code === MAC_PATH_FORBIDDEN
  );
});

test('returns the canonical destination for a safe symlinked write path', async t => {
  const fixture = await tempFixture();
  t.after(() => rm(fixture.base, { recursive: true, force: true }));
  const safeRoot = path.join(fixture.base, 'safe');
  const safeAlias = path.join(fixture.base, 'safe-alias');
  await mkdir(safeRoot, { recursive: true });
  await symlink(safeRoot, safeAlias, process.platform === 'win32' ? 'junction' : 'dir');
  const result = await assertSafeMacPath(path.join(safeAlias, 'new.bin'), {
    homeDir: fixture.homeDir,
    profileDir: fixture.profileDir
  });
  assert.equal(result, path.join(await realpath(safeRoot), 'new.bin'));
});

test('matches Python tilde expansion for custom profiles and rejects ambiguous relative profiles', async t => {
  const fixture = await tempFixture();
  t.after(() => rm(fixture.base, { recursive: true, force: true }));
  const target = path.join(fixture.homeDir, 'custom-profile', 'new.bin');
  await assert.rejects(assertSafeMacPath(target, {
    homeDir: fixture.homeDir, profileDir: '~/custom-profile'
  }), { code: MAC_PATH_FORBIDDEN });
  for (const profileDir of ['relative-profile', '~someone/profile']) {
    await assert.rejects(assertSafeMacPath(target, {
      homeDir: fixture.homeDir, profileDir
    }), { code: 'INVALID_ARGUMENT' });
  }
});

test('canonical protected roots keep a custom profile blocked without current env', async t => {
  const fixture = await tempFixture();
  t.after(() => rm(fixture.base, { recursive: true, force: true }));
  const customProfile = path.join(fixture.base, 'custom-profile');
  const currentProfile = path.join(fixture.base, 'current-profile');
  await mkdir(customProfile);
  await mkdir(currentProfile);
  const roots = await canonicalMacProtectedRoots({
    homeDir: fixture.homeDir,
    profileDir: customProfile
  });
  assert.ok(roots.includes(await realpath(customProfile)));
  await assert.rejects(assertSafeMacPath(path.join(customProfile, 'download.bin'), {
    homeDir: fixture.homeDir,
    profileDir: null,
    protectedRoots: roots
  }), error => error.code === MAC_PATH_FORBIDDEN);
  await assert.rejects(assertSafeMacPath(path.join(currentProfile, 'download.bin'), {
    homeDir: fixture.homeDir,
    profileDir: currentProfile,
    protectedRoots: roots
  }), error => error.code === MAC_PATH_FORBIDDEN);
});

test('canonical protected roots include an auto-discovered profile symlinked outside Outlook', async t => {
  const fixture = await tempFixture();
  t.after(() => rm(fixture.base, { recursive: true, force: true }));
  const outlookRoot = path.join(fixture.homeDir, 'Library', 'Group Containers', 'UBF8T346G9.Office', 'Outlook');
  const profileTarget = path.join(fixture.base, 'external-profile');
  const profileLink = path.join(outlookRoot, 'Outlook 15 Profiles', 'Default Profile');
  await mkdir(profileTarget);
  await mkdir(path.dirname(profileLink), { recursive: true });
  await mkdir(path.join(profileTarget, 'Files'));
  await symlink(profileTarget, profileLink, process.platform === 'win32' ? 'junction' : 'dir');
  await writeFile(path.join(profileTarget, 'HxStore.hxd'), 'synthetic cache');

  const roots = await canonicalMacProtectedRoots({ homeDir: fixture.homeDir });
  const canonicalProfile = await realpath(profileTarget);
  assert.ok(roots.includes(canonicalProfile));
  await assert.rejects(assertSafeMacPath(path.join(profileTarget, 'download.bin'), {
    homeDir: fixture.homeDir,
    profileDir: null,
    protectedRoots: roots
  }), error => error.code === MAC_PATH_FORBIDDEN);
});

test('strict profile discovery rejects a missing profile', async t => {
  const fixture = await tempFixture();
  t.after(() => rm(fixture.base, { recursive: true, force: true }));
  await assert.rejects(resolveMacProfile({ homeDir: fixture.homeDir, strict: true }), {
    code: 'MAC_PROFILE_UNAVAILABLE',
    message: /set OUTLOOK_MAC_PROFILE.*restart the service/
  });
});

test('strict profile discovery rejects an explicitly configured profile without HxStore', async t => {
  const fixture = await tempFixture();
  t.after(() => rm(fixture.base, { recursive: true, force: true }));
  const profile = path.join(fixture.base, 'missing-profile');
  await mkdir(profile);
  await assert.rejects(resolveMacProfile({ homeDir: fixture.homeDir, profileDir: profile, strict: true }), {
    code: 'MAC_PROFILE_UNAVAILABLE'
  });
});

test('strict profile discovery rejects ambiguous profiles', async t => {
  const fixture = await tempFixture();
  t.after(() => rm(fixture.base, { recursive: true, force: true }));
  const profilesRoot = path.join(fixture.homeDir, 'Library', 'Group Containers', 'UBF8T346G9.Office', 'Outlook', 'Outlook 15 Profiles');
  for (const name of ['First Profile', 'Second Profile']) {
    const profile = path.join(profilesRoot, name);
    await mkdir(profile, { recursive: true });
    await writeFile(path.join(profile, 'HxStore.hxd'), 'synthetic cache');
  }
  await assert.rejects(resolveMacProfile({ homeDir: fixture.homeDir, strict: true }), {
    code: 'MAC_PROFILE_UNAVAILABLE'
  });
});
