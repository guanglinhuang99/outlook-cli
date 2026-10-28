import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';

// Exercise the public pagination contract instead of assuming that one read
// contains the whole attachment. Keep the service's default chunk size.
export async function readVerifiedAttachment(call, params) {
  const chunks = [];
  let offset = 0;
  let first;
  for (;;) {
    const page = await call('read_attachment', { ...params, byte_offset: offset });
    first ??= page;
    const chunk = Buffer.from(page.data, 'base64');
    assert.equal(page.offset, offset);
    assert.equal(page.chunk_size, chunk.length);
    assert.equal(page.size_bytes, first.size_bytes);
    assert.equal(page.sha256, first.sha256);
    assert.ok(Number.isSafeInteger(page.size_bytes) && page.size_bytes >= 0 && page.size_bytes <= 10485760);
    assert.ok(offset + chunk.length <= page.size_bytes);
    chunks.push(chunk);
    offset += chunk.length;
    if (page.next_offset === null) break;
    assert.ok(chunk.length > 0 && offset < page.size_bytes);
    assert.equal(page.next_offset, offset);
  }
  const bytes = Buffer.concat(chunks);
  assert.equal(bytes.length, first.size_bytes);
  assert.equal(createHash('sha256').update(bytes).digest('hex'), first.sha256);
  return { bytes, sha256: first.sha256 };
}
