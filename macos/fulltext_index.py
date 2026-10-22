"""Bounded, incremental full-text indexing for verified Outlook bodies.

The index owns only private SQLite state.  Outlook cache files are opened by
``outlook_local``'s verification helpers and are never written here.
"""

from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path
import sqlite3
import stat
import time

try:  # The macOS modules are also imported as top-level modules by the tests.
    import outlook_local as core
except ImportError:  # pragma: no cover - package import fallback
    from . import outlook_local as core


MAX_DB_BYTES = 512 * 1024**2
BODY_SOURCE_BYTES = 128 * 1024**2
BODY_VERIFY_BYTES = 128 * 1024**2
SEARCH_SECONDS = 30.0
QUERY_SECONDS = 30.0
DB_WAIT_SECONDS = 5.0
DB_VERSION = 2
SIDECARS = ("-journal", "-wal", "-shm")
MISSING_BODY_STATUSES = {"not_cached", "not_linked", "unavailable"}


class FullTextIndexError(ValueError):
    """The private full-text cache is unavailable or fails validation."""


class FullTextIndexBusy(FullTextIndexError):
    """Another agent is indexing; retry without rebuilding the cache."""
    retryable = True


class FullTextQueryTimeout(FullTextIndexError):
    """The query needs a narrower candidate scope."""


def _digest(value) -> str:
    try:
        encoded = json.dumps(value, ensure_ascii=False, sort_keys=True,
                             separators=(",", ":"), allow_nan=False).encode("utf-8")
    except (TypeError, ValueError) as error:
        raise ValueError("full-text row contains non-serializable signature data") from error
    return hashlib.sha256(encoded).hexdigest()


def _private_file(path: Path, label: str, *, allow_missing: bool = True) -> None:
    try:
        value = path.lstat()
    except FileNotFoundError:
        if allow_missing:
            return
        raise FullTextIndexError(f"{label} is missing")
    except OSError as error:
        raise FullTextIndexError(f"{label} cannot be inspected") from error
    if stat.S_ISLNK(value.st_mode):
        raise FullTextIndexError(f"{label} must not be a symbolic link")
    if not stat.S_ISREG(value.st_mode):
        raise FullTextIndexError(f"{label} must be a regular file")
    if int(getattr(value, "st_nlink", 0)) != 1:
        raise FullTextIndexError(f"{label} must not be a hard link")
    uid = getattr(os, "getuid", lambda: None)()
    if uid is not None and int(getattr(value, "st_uid", -1)) != int(uid):
        raise FullTextIndexError(f"{label} owner is not the current user")
    if stat.S_IMODE(value.st_mode) & 0o077:
        raise FullTextIndexError(f"{label} does not have private permissions")


def _schema_columns(connection: sqlite3.Connection, table: str) -> tuple[str, ...]:
    try:
        return tuple(str(row[1]) for row in connection.execute(
            f"PRAGMA table_info({table})"))
    except sqlite3.Error as error:
        raise FullTextIndexError("full-text database schema cannot be read") from error


class FullTextIndex:
    """Incrementally index verified plaintext for a current scoped index."""

    _DOCUMENT_COLUMNS = ("doc_key", "id", "signature", "status", "body_status",
                         "text", "folded_text", "error")
    _FTS_COLUMNS = ("id", "signature", "folded_text")

    def __init__(self, path):
        raw = Path(path).expanduser()
        if raw.is_symlink():
            raise FullTextIndexError("full-text database must not be a symbolic link")
        if raw.parent.exists():
            try:
                core._private_directory(raw.parent, "full-text state directory")
            except (OSError, ValueError) as error:
                raise FullTextIndexError(str(error)) from error
        self.path = raw.resolve(strict=False)
        # The service is constructed before its state directory is created.
        # Validate already-existing paths now, and defer SQLite opening until
        # the first search after state_root has established the private root.
        if self.path.parent.exists():
            try:
                core._private_directory(self.path.parent, "full-text state directory")
            except (OSError, ValueError) as error:
                raise FullTextIndexError(str(error)) from error
            self._guard_paths()
        self._integrity_checked = False

    def _guard_paths(self) -> None:
        _private_file(self.path, "full-text database")
        journal = Path(str(self.path) + "-journal")
        _private_file(journal, "full-text journal")
        # DELETE journaling is intentional.  WAL and shared-memory files are
        # state SQLite could interpret from another writer, so refuse them.
        for suffix in SIDECARS[1:]:
            sidecar = Path(str(self.path) + suffix)
            if sidecar.is_symlink():
                raise FullTextIndexError(
                    f"full-text sidecar {suffix} must not be a symbolic link")
            if sidecar.exists():
                _private_file(sidecar, f"full-text sidecar {suffix}")
                raise FullTextIndexError(f"full-text sidecar {suffix} exists")
        if journal.exists():
            try:
                if journal.stat().st_size > MAX_DB_BYTES:
                    raise FullTextIndexError("full-text journal exceeds size limit")
            except OSError as error:
                raise FullTextIndexError("full-text journal cannot be inspected") from error

    @staticmethod
    def _database_error(error: Exception) -> FullTextIndexError:
        code = getattr(error, "sqlite_errorcode", 0) & 0xff
        if code in (sqlite3.SQLITE_BUSY, sqlite3.SQLITE_LOCKED):
            return FullTextIndexBusy("Another full-text search is updating the cache; retry shortly")
        if code == sqlite3.SQLITE_FULL:
            return FullTextIndexError("Full-text cache capacity or disk space exhausted")
        return FullTextIndexError(f"full-text database is corrupt or unsupported: {error}")

    def _connect(self, *, check_integrity: bool | None = None) -> sqlite3.Connection:
        if check_integrity is None:
            check_integrity = not self._integrity_checked
        try:
            core._private_directory(self.path.parent, "full-text state directory")
        except (OSError, ValueError) as error:
            raise FullTextIndexError(str(error)) from error
        self._guard_paths()
        connection: sqlite3.Connection | None = None
        try:
            existed = self.path.exists()
            connection = sqlite3.connect(str(self.path), timeout=DB_WAIT_SECONDS)
            # SQLite follows the path while opening; check the resulting file
            # before accepting any state from it.
            if not existed:
                os.chmod(self.path, 0o600)
            _private_file(self.path, "full-text database", allow_missing=False)
            if self.path.stat().st_size > MAX_DB_BYTES:
                raise sqlite3.DatabaseError("SQLite database exceeds size limit")
            connection.execute("PRAGMA trusted_schema=OFF")
            connection.execute("PRAGMA foreign_keys=ON")
            connection.execute(f"PRAGMA busy_timeout={int(DB_WAIT_SECONDS * 1000)}")
            journal_mode = connection.execute("PRAGMA journal_mode=DELETE").fetchone()
            if not journal_mode or str(journal_mode[0]).casefold() != "delete":
                raise sqlite3.DatabaseError("unexpected SQLite journal mode")
            page_size_row = connection.execute("PRAGMA page_size").fetchone()
            page_count_row = connection.execute("PRAGMA page_count").fetchone()
            if not page_size_row or not page_count_row:
                raise sqlite3.DatabaseError("SQLite page metadata unavailable")
            page_size, page_count = int(page_size_row[0]), int(page_count_row[0])
            if page_size <= 0 or page_count < 0:
                raise sqlite3.DatabaseError("invalid SQLite page metadata")
            max_pages = max(1, MAX_DB_BYTES // page_size)
            configured = connection.execute(f"PRAGMA max_page_count={max_pages}").fetchone()
            if not configured or int(configured[0]) > max_pages or page_count > max_pages:
                raise sqlite3.DatabaseError("SQLite page limit exceeded")
            if check_integrity:
                quick_check = connection.execute("PRAGMA quick_check").fetchone()
                if quick_check != ("ok",):
                    raise sqlite3.DatabaseError("SQLite quick_check failed")
            self._ensure_schema(connection)
            if check_integrity:
                fts_check = connection.execute(
                    "INSERT INTO body_fts(body_fts,rank) VALUES('integrity-check',1)"
                ).fetchall()
                if fts_check:
                    raise sqlite3.DatabaseError("SQLite FTS integrity-check failed")
            connection.commit()
            self._guard_paths()
            if check_integrity:
                self._integrity_checked = True
            return connection
        except (OSError, sqlite3.Error, ValueError, TypeError) as error:
            if connection is not None:
                try:
                    connection.rollback()
                    connection.close()
                except sqlite3.Error:
                    pass
            raise self._database_error(error) from error

    def _ensure_schema(self, connection: sqlite3.Connection) -> None:
        version = connection.execute("PRAGMA user_version").fetchone()
        if not version or int(version[0]) not in (0, DB_VERSION):
            raise sqlite3.DatabaseError("unsupported full-text database version")
        documents = connection.execute(
            "SELECT 1 FROM sqlite_master WHERE type='table' AND name='documents'"
        ).fetchone()
        fts = connection.execute(
            "SELECT 1 FROM sqlite_master WHERE type='table' AND name='body_fts'"
        ).fetchone()
        if not documents and not fts:
            connection.execute(
                "CREATE TABLE documents ("
                "doc_key INTEGER PRIMARY KEY, id TEXT NOT NULL UNIQUE, "
                "signature TEXT NOT NULL, status TEXT NOT NULL, body_status TEXT, "
                "text TEXT, folded_text TEXT, error TEXT)"
            )
            connection.execute(
                "CREATE VIRTUAL TABLE body_fts USING fts5("
                "id UNINDEXED, signature UNINDEXED, folded_text, "
                "content='documents', content_rowid='doc_key', tokenize='trigram')"
            )
            connection.execute(f"PRAGMA user_version={DB_VERSION}")
            return
        if not documents or not fts:
            raise sqlite3.DatabaseError("full-text database schema mismatch")
        if _schema_columns(connection, "documents") != self._DOCUMENT_COLUMNS:
            raise sqlite3.DatabaseError("full-text documents schema mismatch")
        if _schema_columns(connection, "body_fts") != self._FTS_COLUMNS:
            raise sqlite3.DatabaseError("full-text FTS schema mismatch")
        sql_row = connection.execute(
            "SELECT sql FROM sqlite_master WHERE type='table' AND name='body_fts'"
        ).fetchone()
        if not sql_row or "trigram" not in str(sql_row[0]).casefold():
            raise sqlite3.DatabaseError("full-text FTS tokenizer mismatch")
        if int(version[0]) == 0:
            connection.execute(f"PRAGMA user_version={DB_VERSION}")

    @staticmethod
    def _row_signature(profile, row) -> str:
        return _digest({
            "profile": str(profile) if isinstance(profile, Path) else profile,
            "id": row.get("id"),
            "message_id": row.get("message_id"),
            "preview": row.get("preview"),
            "body": row.get("body"),
            "evidence": row.get("evidence"),
        })

    @staticmethod
    def _snippet(text: str, query: str) -> str:
        query_folded = query.casefold()
        if not query_folded:
            return ""
        # Casefold once, then map only the matching range back to source
        # characters.  A position list would multiply memory for large bodies.
        folded = text.casefold()
        start = folded.find(query_folded)
        if start < 0:
            return ""
        if len(text) == len(folded):
            source_start, source_end = start, start + len(query_folded)
        else:
            source_start = source_end = None
            folded_position = 0
            final_folded_position = start + len(query_folded) - 1
            for position, character in enumerate(text):
                next_position = folded_position + len(character.casefold())
                if source_start is None and start < next_position:
                    source_start = position
                if source_end is None and final_folded_position < next_position:
                    source_end = position + 1
                    break
                folded_position = next_position
            if source_start is None or source_end is None:
                return ""
        context = 48
        begin = max(0, source_start - context)
        end = min(len(text), source_end + context, begin + 240)
        value = text[begin:end]
        value = " ".join(value.split())
        if begin:
            value = "…" + value
        if end < len(text):
            value += "…"
        return value

    @staticmethod
    def _body_text(directory, index, row, budget):
        # Keep the entire verified UTF-16 payload; preview is metadata only.
        with core.source_read_budget(budget):
            body = core._resolve_deferred_body(directory, index, row)
            if not isinstance(body, dict) or body.get("status") != "cached_html":
                return None, body
            core._verified_raw_payload(directory, body)
            units = body.get("text_utf16_units")
            if type(units) is not int or units < 0:
                raise ValueError("verified body has invalid UTF-16 length")
            text = core._read_text_units(directory, body, 0, units)
            if not isinstance(text, str):
                raise ValueError("verified body text is not text")
            return text, body

    @staticmethod
    def _fts_delete(connection: sqlite3.Connection, doc_key: int, old):
        if old is not None:
            # External-content FTS5 rows must be removed while their content
            # row still exists.  DELETE lets FTS5 read that row, and also
            # keeps missing-body rows (which have a NULL indexed column) in
            # sync with documents for integrity-check.
            connection.execute("DELETE FROM body_fts WHERE rowid=?", (doc_key,))

    @classmethod
    def _remove_document(cls, connection: sqlite3.Connection, old) -> None:
        if old is None:
            return
        cls._fts_delete(connection, int(old["doc_key"]), old)
        connection.execute("DELETE FROM documents WHERE doc_key=?", (old["doc_key"],))

    @staticmethod
    def _store_document(connection: sqlite3.Connection, identifier: str, signature: str,
                        status: str, body_status: str | None, text: str | None) -> None:
        connection.execute("SAVEPOINT fulltext_row")
        try:
            inserted = connection.execute(
                "INSERT INTO documents(id,signature,status,body_status,text,folded_text,error) "
                "VALUES(?,?,?,?,?,?,NULL)",
                (identifier, signature, status, body_status, text,
                 text.casefold() if isinstance(text, str) else None),
            )
            # Keep one external-content row for every document, including
            # cached-missing bodies whose folded_text is NULL.  FTS5's
            # integrity-check requires document and FTS rowids to agree;
            # search filters missing rows by status.
            doc_key = inserted.lastrowid
            connection.execute(
                "INSERT INTO body_fts(rowid,id,signature,folded_text) "
                "SELECT doc_key,id,signature,folded_text FROM documents WHERE doc_key=?",
                (doc_key,),
            )
        except Exception:
            # SQLITE_FULL can already have rolled back the entire transaction.
            # Cleanup must never hide its original capacity error.
            try:
                connection.execute("ROLLBACK TO fulltext_row")
                connection.execute("RELEASE fulltext_row")
            except sqlite3.Error:
                pass
            raise
        connection.execute("RELEASE fulltext_row")

    @staticmethod
    def _query_hits(connection: sqlite3.Connection, by_id: dict[str, dict],
                    signatures: dict[str, str], query: str) -> dict[str, str]:
        """Stream bodies under a deadline, retaining only compact hit snippets."""
        folded_query = query.casefold()
        if not folded_query:
            return {}
        deadline = time.monotonic() + QUERY_SECONDS

        def expired():
            return time.monotonic() >= deadline

        def timeout():
            return FullTextQueryTimeout(
                "Full-text query deadline exceeded; narrow received date, folder or person filters")

        connection.set_progress_handler(expired, 1000)
        try:
            if expired():
                raise timeout()
            if "\0" not in query and len(folded_query) >= 3:
                escaped = folded_query.replace('"', '""')
                result = connection.execute(
                    "SELECT d.id,d.text FROM body_fts f "
                    "JOIN documents d ON d.doc_key=f.rowid AND d.status='indexed' "
                    "JOIN eligible e ON e.id=d.id AND e.signature=d.signature "
                    "WHERE f.body_fts MATCH ? AND f.signature=e.signature",
                    (f'"{escaped}"',),
                )
            else:
                result = connection.execute(
                    "SELECT d.id,d.text FROM documents d "
                    "JOIN eligible e ON e.id=d.id AND e.signature=d.signature "
                    "WHERE d.status='indexed' AND instr(d.folded_text,?)>0",
                    (folded_query,),
                )
            hits = {}
            for identifier, text in result:
                if expired():
                    raise timeout()
                if isinstance(text, str):
                    # Literal verification also maps Unicode casefold expansions.
                    snippet = FullTextIndex._snippet(text, query)
                    if snippet:
                        hits[identifier] = snippet
            if expired():
                raise timeout()
            return hits
        except sqlite3.Error as error:
            if getattr(error, 'sqlite_errorcode', 0) == sqlite3.SQLITE_INTERRUPT:
                raise timeout() from error
            raise
        finally:
            connection.set_progress_handler(None, 0)

    def search(self, directory, index, rows, query):
        """Index current candidate rows and return ``(hits, stats)``."""
        if not isinstance(query, str):
            raise ValueError("full-text query must be text")
        profile = index.get("profile") if isinstance(index, dict) else None
        by_id: dict[str, dict] = {}
        signatures: dict[str, str] = {}
        for row in rows:
            identifier = row.get("id") if isinstance(row, dict) else None
            if not isinstance(identifier, str) or not identifier:
                raise ValueError("full-text candidate has an invalid id")
            if identifier in by_id:
                raise ValueError("full-text candidates contain duplicate ids")
            by_id[identifier] = row
            signatures[identifier] = self._row_signature(profile, row)
        scope_ids: set[str] = set()
        for scope_row in (index.get("items", []) if isinstance(index, dict) else []):
            scope_id = scope_row.get("id") if isinstance(scope_row, dict) else None
            if not isinstance(scope_id, str) or not scope_id:
                raise ValueError("full-text scope contains an invalid id")
            if scope_id in scope_ids:
                raise ValueError("full-text scope contains duplicate ids")
            scope_ids.add(scope_id)

        connection = self._connect()
        deadline = time.monotonic() + SEARCH_SECONDS
        budget = {"remaining": BODY_SOURCE_BYTES,
                  "verification_remaining": BODY_VERIFY_BYTES}
        effective: dict[str, str] = {}
        errors = 0
        pending = 0
        updated = 0
        reused = 0
        try:
            connection.execute("BEGIN IMMEDIATE")
            connection.execute(
                "CREATE TEMP TABLE eligible(id TEXT PRIMARY KEY, signature TEXT NOT NULL)"
            )
            connection.execute(
                "CREATE TEMP TABLE scope_ids(id TEXT PRIMARY KEY)"
            )
            if signatures:
                connection.executemany(
                    "INSERT INTO eligible(id,signature) VALUES(?,?)",
                    signatures.items(),
                )
            if scope_ids:
                connection.executemany(
                    "INSERT INTO scope_ids(id) VALUES(?)",
                    ((identifier,) for identifier in scope_ids),
                )
            existing_rows = connection.execute(
                "SELECT d.doc_key,d.id,d.signature,d.status,d.body_status,d.error "
                "FROM documents d JOIN eligible e ON e.id=d.id"
            ).fetchall()
            existing = {
                row[1]: {"doc_key": row[0], "id": row[1], "signature": row[2],
                         "status": row[3], "body_status": row[4], "error": row[5]}
                for row in existing_rows
            }
            # Remove rows outside the current scope in one set operation.
            connection.execute(
                "DELETE FROM body_fts WHERE rowid IN ("
                "SELECT doc_key FROM documents WHERE id NOT IN (SELECT id FROM scope_ids))"
            )
            connection.execute(
                "DELETE FROM documents WHERE id NOT IN (SELECT id FROM scope_ids)"
            )

            for identifier, row in by_id.items():
                signature = signatures[identifier]
                cached = existing.get(identifier)
                if (cached is not None and cached["signature"] == signature
                        and cached["status"] in {"indexed", "missing"}):
                    effective[identifier] = cached["status"]
                    reused += 1
                    continue
                if cached is not None:
                    self._remove_document(connection, cached)
                if time.monotonic() >= deadline:
                    effective[identifier] = "pending"
                    pending += 1
                    break
                try:
                    text, body = self._body_text(directory, index, row, budget)
                    if text is None:
                        body_status = body.get("status") if isinstance(body, dict) else None
                        if body_status not in MISSING_BODY_STATUSES:
                            raise ValueError(f"body resolution returned {body_status or 'no status'}")
                        status = "missing"
                        self._store_document(connection, identifier, signature, status,
                                             body_status, None)
                        effective[identifier] = status
                    else:
                        body_status = body.get("status") if isinstance(body, dict) else "cached_html"
                        self._store_document(connection, identifier, signature, "indexed",
                                             body_status, text)
                        effective[identifier] = "indexed"
                    updated += 1
                except Exception as error:
                    if isinstance(error, sqlite3.Error):
                        # A cache write/read failure is a database failure, not
                        # a body that happens to be absent.  Abort this
                        # transaction so callers never receive a false
                        # complete/no-hit result from a damaged cache.
                        raise self._database_error(error) from error
                    source_budget_error = getattr(core, "SourceBudgetExceeded", ())
                    if source_budget_error and isinstance(error, source_budget_error):
                        effective[identifier] = "pending"
                        pending += 1
                        break
                    errors += 1
                    effective[identifier] = "error"
                if time.monotonic() >= deadline:
                    # A completed row is accounted for; leave the next row
                    # explicitly pending so a later call resumes deterministically.
                    remaining = len(by_id) - len(effective)
                    if remaining:
                        pending += remaining
                        for next_id in list(by_id)[len(effective):]:
                            effective[next_id] = "pending"
                    break
            for identifier in by_id:
                effective.setdefault(identifier, "pending")
            pending = sum(value == "pending" for value in effective.values())
            indexed = sum(value == "indexed" for value in effective.values())
            missing = sum(value == "missing" for value in effective.values())
            # Errors are retryable and deliberately have no durable row.  They
            # are kept in the revision for this call, while the next call
            # retries them.
            revision = _digest(sorted((identifier, signatures[identifier], effective[identifier])
                                      for identifier in by_id))
            complete = bool(indexed == len(by_id) and not missing and not errors and not pending)
            stats = {"candidates": len(by_id), "indexed": indexed, "missing": missing,
                     "errors": errors, "pending": pending, "updated": updated,
                     "reused": reused, "complete": complete, "revision": revision,
                     "engine": "sqlite-fts5-trigram"}
            hits = {}
            if query and indexed:
                try:
                    # Keep the temp eligible table and write transaction alive
                    # through the hit query.  A second connection could see a
                    # different candidate signature between stats and hits.
                    hits = self._query_hits(connection, by_id, signatures, query)
                except FullTextQueryTimeout:
                    # Preserve completed indexing so a narrower retry can reuse it.
                    connection.commit()
                    raise
                except sqlite3.Error as error:
                    raise FullTextIndexError(f"full-text query failed: {error}") from error
            connection.commit()
        except Exception as error:
            try:
                connection.rollback()
            except sqlite3.Error:
                pass
            if isinstance(error, sqlite3.Error):
                raise self._database_error(error) from error
            raise
        finally:
            connection.close()
        return hits, stats
