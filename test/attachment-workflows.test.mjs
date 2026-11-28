import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { mkdtemp, readFile, readdir, realpath, rm, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { downloadAllAttachments, readAttachmentText } from '../src/attachment-workflows.mjs';
import { assertSafeMacPath, MAC_PATH_FORBIDDEN } from '../src/mac-paths.mjs';

const id = 'fixture-message-id';
const state = { port: 17379, token: 'fixture', pid: 1, startedAt: 'before', protected_mail_roots: [] };
const fixture = (filename, data, extras = {}) => ({ filename, data: Buffer.from(data), ...extras });
const digest = bytes => createHash('sha256').update(bytes).digest('hex');

function fakeBackend(items, extraOptions = {}) {
  const calls = [];
  const { backendState = state, mutateChunk, ...overrides } = extraOptions;
  return { calls, ...overrides, platform: overrides.platform ?? 'linux',
    readState: overrides.readState ?? (() => backendState),
    call: async (method, input, options) => {
      assert.equal(options.state, backendState);
      assert.equal(input.id, id);
      calls.push({ method, input });
      const context = { backend: 'mac-cache', snapshot: 'test-snapshot', data_scope: 'cached_only' };
      if (method === 'attachments') {
        return { ...context, attachments_complete: false, attachments: items.map((item, i) => ({
          index: i + 1, filename: item.filename, size_bytes: item.data.length,
          type: 1, status: item.status ?? 'cached_at_snapshot'
        })) };
      }
      assert.equal(method, 'read_attachment');
      const item = items[input.index - 1];
      const end = Math.min(input.byte_offset + input.chunk_bytes, item.data.length);
      const data = item.data.subarray(input.byte_offset, end);
      const chunk = { ...context, filename: item.filename, offset: input.byte_offset, chunk_size: data.length,
        encoding: 'base64', size_bytes: item.data.length, next_offset: end < item.data.length ? end : null,
        sha256: item.hash ?? digest(item.data), data: data.toString('base64') };
      return mutateChunk?.(chunk, input) ?? chunk;
    }
  };
}

test('downloads verified chunks with safe names and per-item failure details', async t => {
  const directory = await mkdtemp(path.join(os.tmpdir(), 'outlook-attachments-'));
  t.after(() => rm(directory, { recursive: true, force: true }));
  const old = path.join(directory, 'report.txt');
  await writeFile(old, 'leave intact');
  const big = Buffer.alloc(1024 * 1024 + 1, 42);
  const backend = fakeBackend([
    fixture('../report.txt', big),
    fixture('..\\report.txt', 'second attachment'),
    fixture('pending.txt', 'absent', { status: 'not_cached' }),
    fixture('corrupt.txt', 'corrupted', { hash: '0'.repeat(64) })
  ]);
  const result = await downloadAllAttachments({ id, out_dir: directory }, backend);
  assert.equal(result.attachments_complete, false);
  assert.equal(result.complete, false);
  assert.deepEqual(result.downloaded.map(item => path.basename(item.path)), ['report-2.txt', 'report-3.txt']);
  assert.equal(result.downloaded[0].sha256, digest(big));
  assert.deepEqual(await readFile(result.downloaded[0].path), big);
  assert.equal(await readFile(old, 'utf8'), 'leave intact');
  assert.deepEqual(result.errors.map(item => item.code), ['ATTACHMENT_NOT_CACHED', 'ATTACHMENT_INTEGRITY_ERROR']);
  assert.deepEqual((await readdir(directory)).sort(), ['report-2.txt', 'report-3.txt', 'report.txt']);
  assert.deepEqual(backend.calls.filter(call => call.input.index === 1).map(call => call.input.byte_offset), [0, 1024 * 1024]);
});

test('state changes and malformed chunks never publish an output file', async t => {
  const directory = await mkdtemp(path.join(os.tmpdir(), 'outlook-attachments-'));
  t.after(() => rm(directory, { recursive: true, force: true }));
  let count = 0;
  const common = [fixture('report.txt', 'hello world')];
  const changed = fakeBackend(common, { readState: () => ++count <= 2 ? state : { ...state, pid: 2 } });
  const failed = await downloadAllAttachments({ id, out_dir: directory }, changed);
  assert.equal(failed.errors[0].code, 'SERVICE_RESTART_REQUIRED');
  const stalled = fakeBackend(common, {
    mutateChunk: chunk => ({ ...chunk, data: '', chunk_size: 0, next_offset: 0 })
  });
  const rejected = await downloadAllAttachments({ id, out_dir: directory }, stalled);
  assert.equal(rejected.errors[0].code, 'ATTACHMENT_INTEGRITY_ERROR');
  assert.deepEqual(await readdir(directory), []);
});

test('rejects a protected macOS directory before listing attachments', async t => {
  const directory = await mkdtemp(path.join(os.tmpdir(), 'outlook-attachments-protected-'));
  t.after(() => rm(directory, { recursive: true, force: true }));
  const canonical = await realpath(directory);
  const backend = fakeBackend([], {
    platform: 'darwin', backendState: { ...state, protected_mail_roots: [canonical] },
    assertSafePath: (target, options) => assertSafeMacPath(target, { ...options, homeDir: canonical, profileDir: null })
  });
  await assert.rejects(downloadAllAttachments({ id, out_dir: directory }, backend), { code: MAC_PATH_FORBIDDEN });
  assert.deepEqual(backend.calls, []);
});

test('decodes UTF-8 and UTF-16 text and truncates without splitting surrogates', async () => {
  const utf16 = Buffer.concat([Buffer.from([255, 254]), Buffer.from('业务数据', 'utf16le')]);
  const backend = fakeBackend([fixture('notes.txt', 'a😄b'), fixture('表格.csv', utf16),
    fixture('binary.log', Buffer.from([0, 1, 2])), fixture('unsupported.exe', 'hello')]);
  const short = await readAttachmentText({ id, index: 1, max_chars: 2 }, backend);
  assert.equal(short.text, 'a');
  assert.equal(short.text_chars, 1);
  assert.equal(short.truncated, true);
  const complete = await readAttachmentText({ id, index: 2 }, backend);
  assert.equal(complete.text, '业务数据');
  assert.equal(complete.truncated, false);
  await assert.rejects(readAttachmentText({ id, index: 3 }, backend), { code: 'ATTACHMENT_ENCODING_UNSUPPORTED' });
  await assert.rejects(readAttachmentText({ id, index: 4 }, backend), { code: 'ATTACHMENT_FORMAT_UNSUPPORTED' });
});

function crc32(bytes) {
  let crc = 0xffffffff;
  for (const byte of bytes) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit++) crc = (crc >>> 1) ^ (crc & 1 ? 0xedb88320 : 0);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function storedZip(entries) {
  const files = [];
  const index = [];
  let offset = 0;
  for (const [key, value] of Object.entries(entries)) {
    const name = Buffer.from(key);
    const data = Buffer.from(value);
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50);
    local.writeUInt16LE(20, 4);
    local.writeUInt32LE(crc32(data), 14);
    local.writeUInt32LE(data.length, 18);
    local.writeUInt32LE(data.length, 22);
    local.writeUInt16LE(name.length, 26);
    const central = Buffer.alloc(46);
    central.writeUInt32LE(0x02014b50);
    central.writeUInt16LE(20, 4);
    central.writeUInt16LE(20, 6);
    central.writeUInt32LE(crc32(data), 16);
    central.writeUInt32LE(data.length, 20);
    central.writeUInt32LE(data.length, 24);
    central.writeUInt16LE(name.length, 28);
    central.writeUInt32LE(offset, 42);
    files.push(local, name, data);
    index.push(central, name);
    offset += local.length + name.length + data.length;
  }
  const directory = Buffer.concat(index);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50);
  end.writeUInt16LE(index.length / 2, 8);
  end.writeUInt16LE(index.length / 2, 10);
  end.writeUInt32LE(directory.length, 12);
  end.writeUInt32LE(offset, 16);
  return Buffer.concat([...files, directory, end]);
}

const python = process.env.OUTLOOK_MAC_PYTHON ?? (process.platform === 'win32' ? 'python' : 'python3');
const pythonUnavailable = spawnSync(python, ['--version']).status !== 0;
test('extracts Word paragraphs and worksheet values and rejects embedded entity declarations',
  { skip: pythonUnavailable }, async () => {
    const docx = storedZip({
      'word/document.xml': '<w:document xmlns:w="urn:word"><w:body><w:p><w:r><w:t>销售</w:t><w:tab/><w:t>计划</w:t></w:r></w:p><w:p><w:r><w:t>Next row</w:t></w:r></w:p></w:body></w:document>'
    });
    const excel = storedZip({
      'xl/workbook.xml': '<workbook xmlns:r="urn:relationships"><sheets><sheet name="2026" r:id="r1"/></sheets></workbook>',
      'xl/_rels/workbook.xml.rels': '<Relationships><Relationship Id="r1" Type="/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
      'xl/sharedStrings.xml': '<sst><si><t>收入</t></si></sst>',
      'xl/worksheets/sheet1.xml': '<worksheet><sheetData><row><c t="s"><v>0</v></c><c><v>123.4</v></c><c t="inlineStr"><is><t>人民币</t></is></c></row></sheetData></worksheet>'
    });
    const unsafe = storedZip({
      'word/document.xml': '<!DOCTYPE document [<!ENTITY external SYSTEM "file:///etc/passwd">]><document><p><t>&external;</t></p></document>'
    });
    const backend = fakeBackend([fixture('doc.docx', docx), fixture('book.xlsx', excel), fixture('malformed.docx', unsafe)]);
    assert.equal((await readAttachmentText({ id, index: 1 }, backend)).text, '销售\t计划\nNext row');
    assert.equal((await readAttachmentText({ id, index: 2 }, backend)).text, '2026\n收入\t123.4\t人民币');
    await assert.rejects(readAttachmentText({ id, index: 3 }, backend), { code: 'ATTACHMENT_TEXT_FAILED' });
  });

function pdfFixture() {
  const stream = 'BT /F1 12 Tf 72 700 Td (Local PDF attachment) Tj ET';
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>',
    '<< /Length ' + stream.length + ' >>\nstream\n' + stream + '\nendstream',
    '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>'
  ];
  let source = '%PDF-1.4\n';
  const offsets = ['0000000000 65535 f '];
  for (const [i, object] of objects.entries()) {
    offsets.push(String(source.length).padStart(10, '0') + ' 00000 n ');
    source += (i + 1) + ' 0 obj\n' + object + '\nendobj\n';
  }
  const xref = source.length;
  source += 'xref\n0 6\n' + offsets.join('\n') + '\ntrailer\n<< /Size 6 /Root 1 0 R >>\nstartxref\n' + xref + '\n%%EOF\n';
  return source;
}

test('reports a missing local PDF extractor', async () => {
  const backend = fakeBackend([fixture('report.pdf', pdfFixture())], { pdfCommand: 'test-missing-pdf-extractor' });
  await assert.rejects(readAttachmentText({ id, index: 1 }, backend), { code: 'ATTACHMENT_EXTRACTOR_UNAVAILABLE' });
});

test('extracts text from a PDF through local pdftotext', { skip: spawnSync('pdftotext', ['-v']).status !== 0 },
  async () => {
    const backend = fakeBackend([fixture('report.pdf', pdfFixture())]);
    const result = await readAttachmentText({ id, index: 1 }, backend);
    assert.equal(result.format, 'pdf');
    assert.equal(result.text, 'Local PDF attachment');
    assert.equal(result.truncated, false);
  });
