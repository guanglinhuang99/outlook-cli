import { OutlookBridge } from './bridge.mjs';
import { LocalFileBridge, decodeId, localError } from './offline.mjs';

export class BackendRouter {
  constructor({ com = new OutlookBridge(), local = new LocalFileBridge(), mode = process.env.OUTLOOK_LOCAL_BACKEND ?? 'auto' } = {}) {
    if (!['auto', 'com', 'local'].includes(mode)) throw new Error('OUTLOOK_LOCAL_BACKEND must be auto, com or local');
    this.com = com;
    this.local = local;
    this.mode = mode;
  }
  get queue() { return [...this.com.queue, ...this.local.queue]; }
  get active() { return this.com.active || this.local.active; }
  async call(method, params = {}) {
    const { backend, ...args } = params;
    let mode = backend ?? this.mode;
    const id = decodeId(args.id ?? args.folder);
    if (mode === 'auto' && id?.backend === 'local') mode = 'local';
    if (mode === 'local') return this.local.call(method, args);
    if (id?.backend === 'local') throw localError('BACKEND_ID_MISMATCH', 'Local IDs cannot be passed to COM. List messages with backend=com again.');
    try { return { ...(await this.com.call(method, args)), backend: 'com', data_scope: 'outlook_session' }; }
    catch (error) {
      // A busy Outlook or an unreadable item must not silently change backends.
      if (mode !== 'auto' || error.code !== 'OUTLOOK_UNAVAILABLE') throw error;
      if (id) throw localError('BACKEND_ID_MISMATCH', 'Outlook is closed and this is a COM ID. Relist messages/folders to obtain local-file IDs.');
      return { ...(await this.local.call(method, args)), fallback_reason: 'OUTLOOK_UNAVAILABLE' };
    }
  }
  close() { this.com.close(); this.local.close(); }
}
