import assert from 'node:assert/strict';
import { mkdtemp, readdir, rm, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { buildSwiftHelper } from '../src/swift-helper.mjs';

test('Swift helpers compile once per source version and are reused afterwards', async t => {
  if (process.platform === 'win32') return t.skip('Swift helpers are macOS-only');
  const base = await mkdtemp(path.join(os.tmpdir(), 'outlook-swift-'));
  t.after(() => rm(base, { recursive: true, force: true }));
  const source = path.join(base, 'embed.swift');
  await writeFile(source, 'print("v1")\n');
  const stateDir = path.join(base, 'helpers');
  const commands = [];
  const run = async (command, args) => {
    commands.push(command);
    if (command.endsWith('swiftc')) await writeFile(args[args.indexOf('-o') + 1], '#!/bin/sh\n');
  };
  const fail = reason => new Error(reason);
  const first = await buildSwiftHelper({ stateDir, source, name: 'embed', fail, run });
  const second = await buildSwiftHelper({ stateDir, source, name: 'embed', fail, run });
  assert.equal(second, first);
  assert.deepEqual(commands, ['/usr/bin/swiftc'], 'the cached executable is reused without recompiling');

  await writeFile(source, 'print("v2")\n');
  const updated = await buildSwiftHelper({ stateDir, source, name: 'embed', fail, run, sign: 'example.helper' });
  assert.notEqual(updated, first);
  assert.deepEqual(commands, ['/usr/bin/swiftc', '/usr/bin/swiftc', '/usr/bin/codesign']);
  assert.equal((await readdir(stateDir)).filter(name => name.startsWith('build-')).length, 0);
});
