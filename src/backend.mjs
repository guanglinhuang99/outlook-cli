import { OutlookBridge } from './bridge.mjs';
import { MacCacheBridge } from './mac.mjs';
import { LocalFileBridge, decodeId, localError } from './offline.mjs';
import { isBundleWorkflow } from './tools.mjs';
import { formatDay, parseDay, singaporeDay, singaporeMidnight, singaporeTimezone } from './singapore-dates.mjs';
import { runWindowsBundle } from './windows-bundles.mjs';

const modes = ['auto', 'com', 'local', 'mac-cache'];
function withWindowsDateBasis(result) {
  return { ...result, date_basis: 'received_at' };
}

export class BackendRouter {
  constructor({ com, local, mac, mode = process.env.OUTLOOK_LOCAL_BACKEND ?? 'auto', platform = process.platform, now = () => Date.now() } = {}) {
    if (!modes.includes(mode)) throw new Error('OUTLOOK_LOCAL_BACKEND must be auto, com, local or mac-cache');
    this.mode = mode;
    this.platform = platform;
    this.now = now;
    if (platform === 'darwin') {
      this.mac = mac ?? new MacCacheBridge();
      this.com = com;
      this.local = local;
    } else if (platform === 'win32') {
      this.com = com ?? new OutlookBridge();
      this.local = local ?? new LocalFileBridge();
      this.mac = mac;
    } else {
      this.com = com;
      this.local = local;
      this.mac = mac;
    }
  }
  get bridges() { return [...new Set([this.com, this.local, this.mac].filter(Boolean))]; }
  get queue() { return this.bridges.flatMap(bridge => bridge.queue ?? []); }
  get active() { return this.bridges.find(bridge => bridge.active) ?? false; }
  async call(method, params = {}) {
    if (method === 'today') return this.today(params);
    const { backend, com_body_mode, ...args } = params;
    const mode = backend ?? this.mode;
    const id = decodeId(args.id ?? args.folder);
    if (!modes.includes(mode)) throw localError('INVALID_ARGUMENT', 'backend must be auto, com, local or mac-cache');
    if (this.platform !== 'darwin' && this.platform !== 'win32') throw localError('UNSUPPORTED_PLATFORM', 'This platform is not supported');
    if (this.platform === 'darwin') {
      if (mode === 'com') throw localError('UNSUPPORTED_BACKEND', 'The COM backend is only available on Windows');
      if (!['auto', 'local', 'mac-cache'].includes(mode)) throw localError('UNSUPPORTED_BACKEND', 'The requested backend is not available on macOS');
      if (id && id.backend !== 'mac-cache') throw localError('BACKEND_ID_MISMATCH', 'Windows IDs cannot be passed to the macOS cache backend. List messages again with backend=mac-cache.');
      if (!this.mac) throw localError('UNSUPPORTED_BACKEND', 'The macOS cache backend is unavailable');
      return this.mac.call(method, args);
    }
    if (mode === 'mac-cache') throw localError('UNSUPPORTED_BACKEND', 'The macOS cache backend is only available on macOS');
    if (id?.backend === 'mac-cache') throw localError('BACKEND_ID_MISMATCH', 'macOS cache IDs cannot be passed to Windows. List messages again with backend=com or local.');
    if (isBundleWorkflow(method)) {
      const bundleArgs = method === 'today_bundle' && !args.cursor && !args.date
        ? { ...args, date: singaporeDay(this.now()) }
        : args;
      return runWindowsBundle(method, { ...bundleArgs, backend },
        (operation, input) => this.call(operation, { ...input, backend }));
    }
    if (mode === 'local') {
      if (!this.local) throw localError('UNSUPPORTED_BACKEND', 'The local-file backend is unavailable');
      return withWindowsDateBasis(await this.local.call(method, args));
    }
    if (id?.backend === 'local') {
      if (mode === 'com') throw localError('BACKEND_ID_MISMATCH', 'Local IDs cannot be passed to COM. List messages with backend=com again.');
      if (!this.local) throw localError('UNSUPPORTED_BACKEND', 'The local-file backend is unavailable');
      const result = await this.local.call(method, args);
      return withWindowsDateBasis({ ...result, warnings: [...(result.warnings ?? []), {
        code: 'LOCAL_ID_ROUTING', message: 'This ID belongs to the local-file backend; reading cached data only. List messages/folders with backend=com to use Outlook.'
      }] });
    }
    if (!this.com) throw localError('UNSUPPORTED_BACKEND', 'The COM backend is unavailable');
    try {
      const comArgs = com_body_mode === undefined ? args : { ...args, com_body_mode };
      return withWindowsDateBasis({ ...(await this.com.call(method, comArgs)), backend: 'com', data_scope: 'outlook_session' });
    }
    catch (error) {
      // A busy Outlook or an unreadable item must not silently change backends.
      if (mode !== 'auto' || !['OUTLOOK_UNAVAILABLE', 'OUTLOOK_START_FAILED'].includes(error.code)) throw error;
      if (id) throw localError('BACKEND_ID_MISMATCH', `${error.code}: ${error.message} This is a COM ID and cannot be read by the local-file backend. Retry with Outlook available, or relist messages/folders with backend=local to obtain local-file IDs.`);
      if (!this.local) throw localError('UNSUPPORTED_BACKEND', 'The local-file backend is unavailable');
      let result;
      try { result = await this.local.call(method, args); }
      catch (localFailure) {
        throw localError('OUTLOOK_FALLBACK_FAILED', `COM unavailable (${error.code}: ${error.message}); local-file fallback also failed (${localFailure.code ?? 'OFFLINE_ERROR'}: ${localFailure.message}).`);
      }
      return withWindowsDateBasis({ ...result, backend: 'local', data_scope: 'cached_only',
        fallback_reason: error.code, fallback_message: error.message,
        warnings: [...(result.warnings ?? []), { code: 'LOCAL_FALLBACK',
          message: `Fell back to local OST/PST because ${error.code}: ${error.message} Only locally cached mail is available; no server synchronization is performed by this backend.`
        }] });
    }
  }
  async today(params = {}) {
    const date = params.date ?? singaporeDay(this.now());
    const start = parseDay(date);
    const next = formatDay(new Date(start.getTime() + 86400000));
    const args = { ...params, after: singaporeMidnight(date), before: singaporeMidnight(next, true) };
    delete args.date;
    const result = await this.call('messages', args);
    return { ...result, date, timezone: singaporeTimezone, date_basis: 'received_at' };
  }
  close() { for (const bridge of this.bridges) bridge.close(); }
}
