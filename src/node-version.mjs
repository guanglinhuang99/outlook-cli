// Windows source installs need Node 24: on Node 22 for Windows, file identity checks used
// by Graph sign-in and the knowledge archive fail, and built-in SQLite lacks FTS5 before 22.16.
// The Windows installer bundles Node 24.
export const WINDOWS_MINIMUM_NODE_MAJOR = 24;

export function nodeVersionError(platform = process.platform, version = process.versions.node) {
  const major = Number(version.split('.')[0]);
  if (platform !== 'win32' || major >= WINDOWS_MINIMUM_NODE_MAJOR) return null;
  return Object.assign(new Error(`On Windows this tool requires Node.js ${WINDOWS_MINIMUM_NODE_MAJOR}+ (this is ${version}). Use the installer's bundled Node or install Node.js ${WINDOWS_MINIMUM_NODE_MAJOR} LTS.`), {
    code: 'NODE_VERSION_UNSUPPORTED'
  });
}
