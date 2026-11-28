import { execFile } from 'node:child_process';
import { createHash } from 'node:crypto';
import { chmod, lstat, mkdtemp, readFile, realpath, rename, rm, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { promisify } from 'node:util';
import { prepareRuntime } from './runtime.mjs';

const exec = promisify(execFile);

// Compiles a bundled Swift source once per source version into a private state
// directory and returns the cached executable. `files` are written next to the
// source and included in the version; `flags(build)` adds compiler arguments.
export async function buildSwiftHelper({ stateDir, source, name, files = {}, flags = () => [], sign, fail, run = exec }) {
  await prepareRuntime(stateDir);
  const directory = await realpath(stateDir);
  const sourceText = await readFile(source, 'utf8');
  const version = createHash('sha256').update(JSON.stringify([process.arch, sourceText, files])).digest('hex');
  const helper = path.join(directory, `${name}-${version}`);
  const info = await lstat(helper).catch(error => {
    if (error?.code === 'ENOENT') return null;
    throw error;
  });
  if (info) {
    if (!info.isFile() || info.nlink !== 1 || info.uid !== process.getuid() || (info.mode & 0o777) !== 0o700) {
      throw fail('helper has an unsafe file type, owner, or permissions');
    }
    return helper;
  }
  const build = await mkdtemp(path.join(directory, 'build-'));
  try {
    for (const [file, content] of Object.entries(files)) await writeFile(path.join(build, file), content, { mode: 0o600 });
    const input = path.join(build, `${name}.swift`);
    const binary = path.join(build, name);
    await writeFile(input, sourceText, { mode: 0o600 });
    try {
      await run('/usr/bin/swiftc', ['-module-cache-path', path.join(build, 'modules'), input, '-o', binary, ...flags(build)],
        { timeout: 120000, maxBuffer: 16384 });
      if (sign) await run('/usr/bin/codesign', ['--force', '--sign', '-', '--identifier', sign, binary],
        { timeout: 10000, maxBuffer: 16384 });
    } catch {
      throw fail('building the local helper requires working macOS Swift command line tools');
    }
    await chmod(binary, 0o700);
    await rename(binary, helper);
    return helper;
  } finally {
    await rm(build, { force: true, recursive: true });
  }
}
