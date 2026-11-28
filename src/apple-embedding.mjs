import { spawn } from 'node:child_process';
import { createHash } from 'node:crypto';
import { mkdtemp, rm } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const helper = fileURLToPath(new URL('../macos/embed.swift', import.meta.url));
const outputLimit = 8 * 1024 * 1024;

function providerError(message, code = 'EMBEDDING_PROVIDER_UNAVAILABLE') {
  return Object.assign(new Error(message), { code });
}

async function invoke(language, texts) {
  const moduleCache = await mkdtemp(path.join(os.tmpdir(), 'outlook-apple-embed-'));
  try {
    return await new Promise((resolve, reject) => {
      const swift = spawn('/usr/bin/swift', ['-module-cache-path', moduleCache, helper], {
        stdio: ['pipe', 'pipe', 'ignore']
      });
      const chunks = [];
      let size = 0;
      let failure;
      const timeout = setTimeout(() => {
        failure = providerError('Apple embedding provider timed out');
        swift.kill('SIGKILL');
      }, 45000);
      swift.once('error', () => { failure = providerError('Could not start the Apple Swift embedding helper'); });
      swift.stdin.on('error', () => {});
      swift.stdout.on('data', chunk => {
        size += chunk.length;
        if (size > outputLimit) {
          failure = providerError('Apple embedding result exceeds the output limit');
          swift.kill('SIGKILL');
        } else chunks.push(chunk);
      });
      swift.once('close', code => {
        clearTimeout(timeout);
        if (failure) { reject(failure); return; }
        if (code !== 0) {
          reject(providerError('Apple sentence embeddings require Swift and an installed language model'));
          return;
        }
        try { resolve(JSON.parse(Buffer.concat(chunks).toString('utf8'))); }
        catch { reject(providerError('Invalid result from the Apple embedding helper')); }
      });
      swift.stdin.end(JSON.stringify({ language, texts }));
    });
  } finally {
    await rm(moduleCache, { recursive: true, force: true });
  }
}

export async function createAppleEmbeddingProvider({ env = process.env } = {}) {
  if (process.platform !== 'darwin') throw providerError('Apple sentence embeddings require macOS');
  const language = env.OUTLOOK_APPLE_EMBEDDING_LANGUAGE ?? 'zh-Hans';
  if (!['zh-Hans', 'en'].includes(language)) {
    throw providerError('OUTLOOK_APPLE_EMBEDDING_LANGUAGE must be zh-Hans or en', 'INVALID_ARGUMENT');
  }
  const description = await invoke(language, []);
  if (!Number.isSafeInteger(description.revision) || !Number.isSafeInteger(description.dimension)
      || description.dimension <= 0) throw providerError('Invalid Apple model identity');
  return {
    model: `apple:${language}`,
    digest: createHash('sha256').update(JSON.stringify([
      'apple', language, os.release(), description.revision, description.dimension
    ])).digest('hex'),
    async embed(texts) {
      if (!Array.isArray(texts) || texts.length === 0 || texts.length > 25
          || texts.some(text => typeof text !== 'string' || Buffer.byteLength(text) > 100000)) {
        throw providerError('Apple embedding requires a bounded batch of 1..25 texts', 'INVALID_ARGUMENT');
      }
      const result = await invoke(language, texts);
      if (result.revision !== description.revision || result.dimension !== description.dimension
          || !Array.isArray(result.embeddings) || result.embeddings.length !== texts.length
          || result.embeddings.some(vector => vector !== null && (!Array.isArray(vector)
            || vector.length !== description.dimension
            || vector.some(value => typeof value !== 'number' || !Number.isFinite(value))))) {
        throw providerError('Apple embedding model changed or returned invalid vectors');
      }
      return result.embeddings;
    }
  };
}
