import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { OutlookBridge } from './bridge.mjs';

export const macWorkerPath = fileURLToPath(new URL('../macos/common_worker.py', import.meta.url));

export class MacCacheBridge extends OutlookBridge {
  constructor({ timeoutMs = 180000, profileDir = process.env.OUTLOOK_MAC_PROFILE, spawnWorker } = {}) {
    const workerEnv = profileDir ? { ...process.env, OUTLOOK_MAC_PROFILE: profileDir } : process.env;
    super({
      timeoutMs,
      enableWindowsFulltext: false,
      spawnWorker: spawnWorker ?? (() => spawn(process.env.OUTLOOK_MAC_PYTHON ?? 'python3', [macWorkerPath], {
        env: workerEnv,
        stdio: ['pipe', 'pipe', 'pipe']
      }))
    });
  }
}
