import { constants } from 'node:fs';
import { chmod, lstat, open as openFile, realpath, stat } from 'node:fs/promises';
import path from 'node:path';

import { prepareRuntime, runtimeDir } from './runtime.mjs';

const DEFAULT_MAX_BYTES = 512 * 1024 * 1024;
const DEFAULT_BODY_CODE_UNITS = 16_000_000;
const DEFAULT_BATCH_BYTES = 128 * 1024 * 1024;
const DEFAULT_DEADLINE_MS = 30_000;
const DEFAULT_BUSY_TIMEOUT_MS = 5_000;
const DEFAULT_CANDIDATE_WINDOW = 50;
const DB_VERSION = 1;
const DB_NAME = 'fulltext.sqlite3';
const JOURNAL_SUFFIX = '-journal';
const FORBIDDEN_SIDECARS = ['-wal', '-shm'];
const DOCUMENT_COLUMNS = ['doc_key', 'id', 'signature', 'text', 'normalized', 'bytes', 'updated_at'];
const MISSING_STATUSES = new Set(['missing', 'truncated', 'unsupported']);

function errorWithCode(message, code, retryable = false) {
  const error = new Error(message);
  error.code = code;
  if (retryable) error.retryable = true;
  return error;
}
function invalid(message) {
  return errorWithCode(message, 'INVALID_ARGUMENT');
}
function isWindows() {
  return process.platform === 'win32';
}

function isPrivateRegularFile(stats, label) {
  if (stats.isSymbolicLink()) throw errorWithCode(`${label} must not be a symbolic link`, 'FULLTEXT_UNSAFE_PATH');
  if (!stats.isFile()) throw errorWithCode(`${label} must be a regular file`, 'FULLTEXT_UNSAFE_PATH');
  if (stats.nlink !== 1) throw errorWithCode(`${label} must not be a hard link`, 'FULLTEXT_UNSAFE_PATH');
  if (!isWindows()) {
    if (typeof stats.uid === 'number' && stats.uid !== process.getuid()) {
      throw errorWithCode(`${label} must be owned by the current user`, 'FULLTEXT_UNSAFE_PATH');
    }
    if ((stats.mode & 0o077) !== 0) {
      throw errorWithCode(`${label} must have private permissions`, 'FULLTEXT_UNSAFE_PATH');
    }
  }
}

async function inspectPrivateFile(file, label, { allowMissing = true } = {}) {
  let stats;
  try {
    stats = await lstat(file);
  } catch (error) {
    if (allowMissing && error?.code === 'ENOENT') return null;
    throw errorWithCode(`${label} cannot be inspected`, 'FULLTEXT_UNSAFE_PATH');
  }
  isPrivateRegularFile(stats, label);
  return stats;
}

function sqliteMessage(error) {
  return String(error?.message ?? error ?? 'SQLite operation failed');
}

function mapSqliteError(error) {
  if (error?.code === 'FULLTEXT_BUSY' || error?.code === 'FULLTEXT_CORRUPT'
      || error?.code === 'FULLTEXT_CAP') return error;
  const message = sqliteMessage(error);
  const lower = message.toLowerCase();
  const numericCode = Number(error?.errcode ?? error?.sqliteErrorCode ?? 0);
  if (lower.includes('busy') || lower.includes('locked') || numericCode === 5 || numericCode === 6) {
    return errorWithCode('Full-text SQLite cache is busy; retry the search.', 'FULLTEXT_BUSY', true);
  }
  if (/database or disk is full|sqlite_full|page limit exceeded|too large/.test(lower) || numericCode === 13) {
    return errorWithCode('Full-text SQLite cache reached its capacity; retry with a narrower search.', 'FULLTEXT_CAP', true);
  }
  if (lower.includes('malformed') || lower.includes('corrupt') || lower.includes('no such table')
      || lower.includes('fts') || lower.includes('schema') || lower.includes('not a database') || lower.includes('quick_check failed')) {
    return errorWithCode(`Full-text SQLite cache is corrupt or unsupported: ${message}`, 'FULLTEXT_CORRUPT', true);
  }
  return errorWithCode(`Full-text SQLite operation failed: ${message}`, 'FULLTEXT_ERROR', true);
}

function sqlPlaceholders(count) {
  return Array.from({ length: count }, () => '?').join(',');
}

function pairPredicate(pairs) {
  if (!pairs.length) return { sql: '0', values: [] };
  const values = [];
  const terms = pairs.map(pair => {
    values.push(pair.id, pair.signature);
    return `(d.id=? AND d.signature=?)`;
  });
  return { sql: terms.join(' OR '), values };
}

function lower(value) {
  return value.toLowerCase();
}

function codePointLength(value) {
  return Array.from(value).length;
}

function sourceRange(text, folded, start, end) {
  if (text.length === folded.length) return { start, end };
  let foldedPosition = 0;
  let sourceStart = null;
  let sourceEnd = null;
  for (let position = 0; position < text.length;) {
    const codePoint = text.codePointAt(position);
    const width = codePoint > 0xffff ? 2 : 1;
    const character = text.slice(position, position + width);
    const next = foldedPosition + character.toLowerCase().length;
    if (sourceStart === null && start < next) sourceStart = position;
    if (sourceEnd === null && end <= next) {
      sourceEnd = position + width;
      break;
    }
    foldedPosition = next;
    position += width;
  }
  return sourceStart === null || sourceEnd === null ? null : { start: sourceStart, end: sourceEnd };
}

function bodySnippet(text, query) {
  const foldedQuery = lower(query);
  if (!foldedQuery) return '';
  const foldedText = lower(text);
  const matchStart = foldedText.indexOf(foldedQuery);
  if (matchStart < 0) return '';
  const range = sourceRange(text, foldedText, matchStart, matchStart + foldedQuery.length);
  if (!range) return '';

  const markerLength = 1;
  const maxLength = 240;
  const context = 96;
  let begin = Math.max(0, range.start - context);
  let end = Math.min(text.length, range.end + context);
  if (begin > 0 && begin < text.length && /[\uDC00-\uDFFF]/u.test(text[begin])) begin -= 1;
  if (end < text.length && end > 0 && /[\uD800-\uDBFF]/u.test(text[end - 1])) end += 1;

  let prefix = begin > 0 ? '…' : '';
  let suffix = end < text.length ? '…' : '';
  let value = text.slice(begin, end);
  const available = maxLength - prefix.length - suffix.length;
  if (value.length > available) {
    const matchLength = Math.min(range.end - range.start, available);
    const roomBefore = Math.max(0, Math.floor((available - matchLength) / 2));
    begin = Math.max(0, range.start - roomBefore);
    end = Math.min(text.length, begin + available);
    if (end < range.end) {
      end = Math.min(text.length, range.end);
      begin = Math.max(0, end - available);
    }
    if (begin > 0 && /[\uDC00-\uDFFF]/u.test(text[begin])) begin -= 1;
    if (end < text.length && end > 0 && /[\uD800-\uDBFF]/u.test(text[end - 1])) end += 1;
    prefix = begin > 0 ? '…' : '';
    suffix = end < text.length ? '…' : '';
    value = text.slice(begin, end);
  }
  value = value.replace(/\s+/gu, ' ').trim();
  value = `${prefix}${value}${suffix}`;
  if (value.length > maxLength) value = `${value.slice(0, maxLength - markerLength)}…`;
  return value;
}

function deadlineFrom(value, now, defaultMs) {
  if (value === undefined || value === null) return now + defaultMs;
  if (value instanceof Date) {
    if (!Number.isFinite(value.getTime())) throw invalid('deadline must be a valid time');
    return value.getTime();
  }
  if (typeof value !== 'number' || !Number.isFinite(value)) throw invalid('deadline must be a finite time');
  // Absolute wall-clock milliseconds are convenient for service callers. Small
  // values are treated as a relative budget so tests can request 0 immediately.
  return value > 1_000_000_000_000 ? value : now + Math.max(0, value);
}

export class WindowsFulltextIndex {
  constructor({
    directory = path.join(runtimeDir, 'fulltext'),
    maxBytes = DEFAULT_MAX_BYTES,
    bodyCapCodeUnits = DEFAULT_BODY_CODE_UNITS,
    singleBatchBytes = DEFAULT_BATCH_BYTES,
    budgetMs = DEFAULT_DEADLINE_MS,
    busyTimeoutMs = DEFAULT_BUSY_TIMEOUT_MS,
    now = Date.now
  } = {}) {
    if (typeof directory !== 'string' || !directory || directory.includes('\0') || !path.isAbsolute(directory)) {
      throw invalid('full-text directory must be an absolute path');
    }
    if (!Number.isSafeInteger(maxBytes) || maxBytes <= 0) throw invalid('maxBytes must be a positive safe integer');
    if (!Number.isSafeInteger(bodyCapCodeUnits) || bodyCapCodeUnits <= 0) throw invalid('bodyCapCodeUnits must be positive');
    if (!Number.isSafeInteger(singleBatchBytes) || singleBatchBytes <= 0) throw invalid('singleBatchBytes must be positive');
    if (!Number.isSafeInteger(budgetMs) || budgetMs < 0) throw invalid('budgetMs must be a non-negative safe integer');
    if (!Number.isSafeInteger(busyTimeoutMs) || busyTimeoutMs < 0) throw invalid('busyTimeoutMs must be a non-negative safe integer');
    if (typeof now !== 'function') throw invalid('now must be a function');
    this.directory = path.resolve(directory);
    this.maxBytes = maxBytes;
    this.bodyCapCodeUnits = bodyCapCodeUnits;
    this.singleBatchBytes = singleBatchBytes;
    this.budgetMs = budgetMs;
    this.busyTimeoutMs = busyTimeoutMs;
    this.now = now;
    this.databasePath = path.join(this.directory, DB_NAME);
    this._preparePromise = null;
    this._sqlitePromise = null;
    this._integrityChecked = false;
  }

  async _ensureDirectory() {
    if (!this._preparePromise) {
      this._preparePromise = (async () => {
        try {
          const requested = await lstat(this.directory);
          if (requested.isSymbolicLink()) {
            throw errorWithCode('full-text directory must not be a symbolic link', 'FULLTEXT_UNSAFE_PATH');
          }
        } catch (error) {
          if (error?.code !== 'ENOENT') throw error;
        }
        let prepared = false;
        for (let attempt = 0; attempt < 4 && !prepared; attempt += 1) {
          try {
            await prepareRuntime(this.directory);
            prepared = true;
          } catch (error) {
            // Two service requests can create the managed leaf together. The
            // loser may observe it before the winner publishes its marker.
            const transient = /not managed|marker already exists|not recognized/i.test(String(error?.message ?? ''));
            if (!transient || attempt === 3) throw error;
            await new Promise(resolve => setTimeout(resolve, 5 * (attempt + 1)));
          }
        }
        const actual = await realpath(this.directory);
        const requested = await lstat(this.directory);
        if (requested.isSymbolicLink()) throw errorWithCode('full-text directory must not be a symbolic link', 'FULLTEXT_UNSAFE_PATH');
        const stats = await lstat(actual);
        if (!stats.isDirectory() || stats.isSymbolicLink()) {
          throw errorWithCode('full-text directory must be a real directory', 'FULLTEXT_UNSAFE_PATH');
        }
        if (!isWindows() && (stats.mode & 0o077) !== 0) {
          throw errorWithCode('full-text directory must have private permissions', 'FULLTEXT_UNSAFE_PATH');
        }
        this.directory = actual;
        this.databasePath = path.join(actual, DB_NAME);
      })();
    }
    try {
      await this._preparePromise;
    } catch (error) {
      this._preparePromise = null;
      throw error;
    }
  }

  async _sqlite() {
    if (!this._sqlitePromise) this._sqlitePromise = import('node:sqlite').catch(() => {
      throw errorWithCode('Windows full-text search requires Node.js 22.13+ with built-in SQLite; use the bundled Node 24 runtime.', 'FULLTEXT_UNAVAILABLE');
    });
    return this._sqlitePromise;
  }

  async _guardSidecars() {
    await inspectPrivateFile(this.databasePath, 'full-text database');
    const journal = await inspectPrivateFile(`${this.databasePath}${JOURNAL_SUFFIX}`, 'full-text journal');
    if (journal && journal.size > this.maxBytes) {
      throw errorWithCode('full-text journal exceeds the configured size limit', 'FULLTEXT_CAP', true);
    }
    for (const suffix of FORBIDDEN_SIDECARS) {
      const sidecar = `${this.databasePath}${suffix}`;
      const stats = await inspectPrivateFile(sidecar, `full-text ${suffix} sidecar`);
      if (stats) throw errorWithCode(`full-text ${suffix} sidecar exists; WAL mode is unsupported`, 'FULLTEXT_UNSAFE_PATH');
    }
  }

  async _createDatabaseFile() {
    await this._guardSidecars();
    try {
      const handle = await openFile(this.databasePath, constants.O_WRONLY | constants.O_CREAT | constants.O_EXCL, 0o600);
      await handle.close();
    } catch (error) {
      if (error?.code !== 'EEXIST') throw error;
    }
    await this._guardSidecars();
  }

  _columns(db, table) {
    return db.prepare(`PRAGMA table_info(${table})`).all().map(row => String(row.name));
  }

  _ensureSchema(db) {
    const readState = () => ({
      version: Number(db.prepare('PRAGMA user_version').get()?.user_version ?? 0),
      documents: db.prepare("SELECT name FROM sqlite_master WHERE type='table' AND name='documents'").get(),
      fts: db.prepare("SELECT name,sql FROM sqlite_master WHERE type='table' AND name='body_fts'").get()
    });
    const validate = state => {
      if (state.version !== 0 && state.version !== DB_VERSION) {
        throw new Error(`unsupported full-text schema version ${state.version}`);
      }
      if (!state.documents && !state.fts) return false;
      if (!state.documents || !state.fts
          || this._columns(db, 'documents').join('\0') !== DOCUMENT_COLUMNS.join('\0')
          || !String(state.fts.sql ?? '').toLowerCase().includes('trigram')
          || !String(state.fts.sql ?? '').toLowerCase().includes("content='documents'")) {
        throw new Error('full-text database schema mismatch');
      }
      return true;
    };

    // Schema creation is serialized so two service processes cannot each see
    // an empty database and race through CREATE TABLE/CREATE VIRTUAL TABLE.
    if (validate(readState())) return;
    db.exec('BEGIN IMMEDIATE');
    try {
      const state = readState();
      if (!validate(state)) {
        db.exec(`
          CREATE TABLE documents (
            doc_key INTEGER PRIMARY KEY,
            id TEXT NOT NULL UNIQUE,
            signature TEXT NOT NULL,
            text TEXT NOT NULL,
            normalized TEXT NOT NULL,
            bytes INTEGER NOT NULL,
            updated_at INTEGER NOT NULL
          );
          CREATE VIRTUAL TABLE body_fts USING fts5(
            id UNINDEXED,
            signature UNINDEXED,
            normalized,
            content='documents',
            content_rowid='doc_key',
            tokenize='trigram'
          );
          PRAGMA user_version=${DB_VERSION};
        `);
      }
      db.exec('COMMIT');
    } catch (error) {
      try { db.exec('ROLLBACK'); } catch {}
      throw error;
    }
  }

  _checkIntegrity(db) {
    const quick = db.prepare('PRAGMA quick_check').get();
    if (String(quick?.quick_check ?? '').toLowerCase() !== 'ok') throw new Error('SQLite quick_check failed');
    const result = db.prepare("INSERT INTO body_fts(body_fts,rank) VALUES('integrity-check',1)").all();
    if (result.length) throw new Error('SQLite FTS integrity-check failed');
  }

  async _openDatabase() {
    await this._createDatabaseFile();
    const { DatabaseSync } = await this._sqlite();
    let db;
    try {
      db = new DatabaseSync(this.databasePath, { timeout: this.busyTimeoutMs });
      db.exec(`PRAGMA trusted_schema=OFF; PRAGMA foreign_keys=ON; PRAGMA busy_timeout=${this.busyTimeoutMs};`);
      const journalMode = db.prepare('PRAGMA journal_mode=DELETE').get();
      if (String(journalMode?.journal_mode ?? '').toLowerCase() !== 'delete') throw new Error('unexpected SQLite journal mode');
      const pageSize = Number(db.prepare('PRAGMA page_size').get()?.page_size ?? 0);
      const pageCount = Number(db.prepare('PRAGMA page_count').get()?.page_count ?? 0);
      if (!Number.isSafeInteger(pageSize) || pageSize <= 0 || !Number.isSafeInteger(pageCount) || pageCount < 0) {
        throw new Error('invalid SQLite page metadata');
      }
      const maxPages = Math.max(1, Math.floor(this.maxBytes / pageSize));
      const configured = Number(db.prepare(`PRAGMA max_page_count=${maxPages}`).get()?.max_page_count ?? 0);
      if (!configured || configured > maxPages || pageCount > maxPages) throw new Error('SQLite page limit exceeded');
      this._ensureSchema(db);
      if (!this._integrityChecked) {
        this._checkIntegrity(db);
        this._integrityChecked = true;
      }
      await chmod(this.databasePath, 0o600).catch(error => {
        if (!isWindows()) throw error;
      });
      await this._guardSidecars();
      const databaseStats = await stat(this.databasePath);
      if (databaseStats.size > this.maxBytes) throw errorWithCode('full-text database exceeds the configured size limit', 'FULLTEXT_CAP', true);
      return db;
    } catch (error) {
      try { db?.close(); } catch {}
      if (error?.code?.startsWith?.('FULLTEXT_')) throw error;
      throw mapSqliteError(error);
    }
  }

  _validateCandidates(candidates) {
    if (!Array.isArray(candidates)) throw invalid('full-text candidates must be an array');
    const seen = new Set();
    return candidates.map((candidate, index) => {
      if (!candidate || typeof candidate !== 'object' || typeof candidate.id !== 'string' || !candidate.id || candidate.id.includes('\0')) {
        throw invalid(`full-text candidate ${index} has an invalid id`);
      }
      if (candidate.signature !== null && typeof candidate.signature !== 'string') {
        throw invalid(`full-text candidate ${candidate.id} has an invalid signature`);
      }
      if (seen.has(candidate.id)) throw invalid(`full-text candidates contain duplicate id ${candidate.id}`);
      seen.add(candidate.id);
      return { ...candidate };
    });
  }

  _existingRows(db, ids) {
    if (!ids.length) return new Map();
    const placeholders = sqlPlaceholders(ids.length);
    const rows = db.prepare(`SELECT d.id,d.signature FROM documents d WHERE d.id IN (${placeholders})`).all(...ids);
    return new Map(rows.map(row => [String(row.id), { id: String(row.id), signature: String(row.signature) }]));
  }

  _removeDocument(db, id) {
    const row = db.prepare('SELECT doc_key FROM documents WHERE id=?').get(id);
    if (!row) return;
    db.prepare('DELETE FROM body_fts WHERE rowid=?').run(row.doc_key);
    db.prepare('DELETE FROM documents WHERE doc_key=?').run(row.doc_key);
  }

  _invalidateCurrent(db, candidates, existing) {
    for (const candidate of candidates) {
      const old = existing.get(candidate.id);
      if (old && (candidate.signature === null || old.signature !== candidate.signature)) {
        this._removeDocument(db, candidate.id);
      }
    }
  }

  _prepareCache(db, candidates) {
    db.exec('BEGIN IMMEDIATE');
    try {
      // Read the rows while holding the write lock. Otherwise another process
      // can replace/delete a version between the read and invalidation.
      const existing = this._existingRows(db, candidates.map(candidate => candidate.id));
      this._invalidateCurrent(db, candidates, existing);
      db.exec('COMMIT');
      return existing;
    } catch (error) {
      try { db.exec('ROLLBACK'); } catch {}
      throw error;
    }
  }

  _storeReady(db, candidate, text) {
    const normalized = lower(text);
    const bytes = Buffer.byteLength(text, 'utf8');
    db.exec('BEGIN IMMEDIATE');
    try {
      this._removeDocument(db, candidate.id);
      const inserted = db.prepare(
        'INSERT INTO documents(id,signature,text,normalized,bytes,updated_at) VALUES(?,?,?,?,?,?)'
      ).run(candidate.id, candidate.signature, text, normalized, bytes, this.now());
      db.prepare(
        'INSERT INTO body_fts(rowid,id,signature,normalized) SELECT doc_key,id,signature,normalized FROM documents WHERE doc_key=?'
      ).run(inserted.lastInsertRowid);
      db.exec('COMMIT');
    } catch (error) {
      try { db.exec('ROLLBACK'); } catch {}
      throw error;
    }
  }

  _queryRows(db, pairs, query) {
    if (!pairs.length || !query) return [];
    const predicate = pairPredicate(pairs);
    const foldedQuery = lower(query);
    const useFts = !foldedQuery.includes('\0') && codePointLength(foldedQuery) >= 3;
    const source = useFts ? 'body_fts f JOIN documents d ON d.doc_key=f.rowid AND d.id=f.id AND d.signature=f.signature' : 'documents d';
    const condition = useFts ? 'f.normalized MATCH ?' : 'instr(d.normalized,?)>0';
    const sql = `SELECT d.id,d.signature FROM ${source} WHERE ${condition} AND (${predicate.sql})`;
    const argument = useFts ? `"${foldedQuery.replaceAll('"', '""')}"` : foldedQuery;
    return db.prepare(sql).all(argument, ...predicate.values);
  }

  _queryHits(db, pairs, ephemeralHits, query) {
    db.exec('BEGIN');
    try {
      const predicate = pairPredicate(pairs);
      const present = pairs.length
        ? db.prepare(`SELECT d.id,d.signature FROM documents d WHERE ${predicate.sql}`).all(...predicate.values)
        : [];
      const presentKeys = new Set(present.map(row => `${String(row.id)}\0${String(row.signature)}`));
      const missingPairs = pairs.filter(pair => !presentKeys.has(`${pair.id}\0${pair.signature}`));
      const hits = new Map(ephemeralHits);
      if (query) {
        const rows = this._queryRows(db, pairs, query);
        const textStatement = db.prepare('SELECT text FROM documents WHERE id=? AND signature=?');
        const foldedQuery = lower(query);
        for (const row of rows) {
          const body = textStatement.get(row.id, row.signature)?.text;
          if (typeof body !== 'string' || !lower(body).includes(foldedQuery)) continue;
          const snippet = bodySnippet(body, query);
          if (snippet) hits.set(String(row.id), snippet);
        }
      }
      db.exec('COMMIT');
      return { hits, missingPairs };
    } catch (error) {
      try { db.exec('ROLLBACK'); } catch {}
      throw error;
    }
  }

  async evaluate(candidates, { query = '', loadBody, deadline } = {}) {
    const allCandidates = this._validateCandidates(candidates);
    if (typeof query !== 'string') throw invalid('full-text query must be text');
    if (typeof loadBody !== 'function') throw invalid('loadBody must be a function');
    const activeCandidates = allCandidates.slice(0, DEFAULT_CANDIDATE_WINDOW);
    const statuses = new Map();
    const indexedIds = new Set();
    const indexedSources = new Map();
    const ephemeralHits = new Map();
    let updated = 0;
    let reused = 0;
    let batchBytesUsed = 0;
    const end = deadlineFrom(deadline, this.now(), this.budgetMs);

    await this._ensureDirectory();
    const db = await this._openDatabase();
    try {
      const existing = this._prepareCache(db, activeCandidates);
      for (const candidate of activeCandidates) {
        if (candidate.signature !== null && existing.get(candidate.id)?.signature === candidate.signature) {
          statuses.set(candidate.id, 'indexed');
          indexedIds.add(candidate.id);
          indexedSources.set(candidate.id, 'reused');
          reused += 1;
        }
      }

      for (let index = 0; index < activeCandidates.length; index += 1) {
        const candidate = activeCandidates[index];
        if (statuses.has(candidate.id)) continue;
        const remainingMs = Math.max(0, end - this.now());
        if (remainingMs <= 0 || batchBytesUsed >= this.singleBatchBytes) {
          statuses.set(candidate.id, 'pending');
          for (const rest of activeCandidates.slice(index + 1)) {
            if (!statuses.has(rest.id)) statuses.set(rest.id, 'pending');
          }
          break;
        }
        let body;
        try {
          body = await loadBody(candidate, remainingMs);
        } catch {
          statuses.set(candidate.id, 'error');
          continue;
        }
        // A source adapter may resolve after its requested timeout. Do not
        // commit a body that arrived outside this evaluation's budget.
        if (this.now() >= end) {
          statuses.set(candidate.id, 'pending');
          for (const rest of activeCandidates.slice(index + 1)) {
            if (!statuses.has(rest.id)) statuses.set(rest.id, 'pending');
          }
          break;
        }
        const status = body?.status;
        if (MISSING_STATUSES.has(status)) {
          statuses.set(candidate.id, 'missing');
          continue;
        }
        if (status === 'changed') {
          statuses.set(candidate.id, 'error');
          continue;
        }
        if (status !== 'ready' || typeof body.text !== 'string' || body.signature !== candidate.signature) {
          statuses.set(candidate.id, 'error');
          continue;
        }
        const bytes = Buffer.byteLength(body.text, 'utf8');
        // An individual body beyond either hard cap is unavailable for this
        // search. It is a cache miss and should be retried when the source can
        // provide a bounded representation, rather than holding pagination in
        // a permanent pending state.
        if (body.text.length > this.bodyCapCodeUnits || bytes > this.singleBatchBytes) {
          statuses.set(candidate.id, 'missing');
          continue;
        }
        // The remaining batch budget is shared by the candidate window. The
        // current body may fit alone but later candidates must wait for the
        // next query once this window is full.
        if (batchBytesUsed + bytes > this.singleBatchBytes) {
          statuses.set(candidate.id, 'pending');
          for (const rest of activeCandidates.slice(index + 1)) {
            if (!statuses.has(rest.id)) statuses.set(rest.id, 'pending');
          }
          break;
        }
        batchBytesUsed += bytes;
        statuses.set(candidate.id, 'indexed');
        indexedIds.add(candidate.id);
        indexedSources.set(candidate.id, 'updated');
        updated += 1;
        if (candidate.signature === null) {
          const snippet = bodySnippet(body.text, query);
          if (snippet) ephemeralHits.set(candidate.id, snippet);
        } else {
          this._storeReady(db, candidate, body.text);
        }
        if (this.now() >= end) {
          for (const rest of activeCandidates.slice(index + 1)) {
            if (!statuses.has(rest.id)) statuses.set(rest.id, 'pending');
          }
          break;
        }
      }
      for (const candidate of allCandidates.slice(activeCandidates.length)) statuses.set(candidate.id, 'pending');
      for (const candidate of allCandidates) statuses.set(candidate.id, statuses.get(candidate.id) ?? 'pending');

      const pairs = activeCandidates
        .filter(candidate => candidate.signature !== null && indexedIds.has(candidate.id))
        .map(candidate => ({ id: candidate.id, signature: candidate.signature }));
      const queried = this._queryHits(db, pairs, ephemeralHits, query);
      // A second process may have replaced or removed a candidate after the
      // body read/store. Keep that row out of the result and force a retry;
      // never claim complete based on a stale pair snapshot.
      for (const pair of queried.missingPairs) {
        if (statuses.get(pair.id) !== 'indexed') continue;
        const source = indexedSources.get(pair.id);
        if (source === 'reused') reused -= 1;
        if (source === 'updated') updated -= 1;
        statuses.set(pair.id, 'pending');
        indexedIds.delete(pair.id);
      }

      const indexed = [...statuses.values()].filter(value => value === 'indexed').length;
      const missing = [...statuses.values()].filter(value => value === 'missing').length;
      const errors = [...statuses.values()].filter(value => value === 'error').length;
      const pending = [...statuses.values()].filter(value => value === 'pending').length;
      // Missing/truncated/unsupported bodies are expected cache misses. They
      // remain incomplete and are retried on the next query, while only an
      // in-flight deadline or a transient/error result holds pagination.
      const retryRequired = Boolean(errors || pending);
      const fulltext = {
        candidates: allCandidates.length,
        indexed,
        missing,
        errors,
        pending,
        updated,
        reused,
        complete: indexed === allCandidates.length && !missing && !errors && !pending,
        engine: 'sqlite-fts5-trigram'
      };
      return {
        hits: queried.hits,
        fulltext,
        retry_required: retryRequired,
        body_search_incomplete: Boolean(missing || errors || pending)
      };
    } catch (error) {
      if (error?.code?.startsWith?.('FULLTEXT_')) throw error;
      throw mapSqliteError(error);
    } finally {
      try { db.close(); } catch {}
    }
  }

  close() {
    // Connections are intentionally per-call, so there is no persistent handle to close.
  }
}
