#!/usr/bin/env python3
"""Read-only MCP stdio server for the Outlook local prototype.

The transport is deliberately small: one JSON-RPC object per input line and
one JSON-RPC response per request that is not a notification.  The Outlook
service is injected by tests and imported lazily by the CLI so initialization
and tool discovery never scan a mailbox.
"""

from __future__ import annotations

import argparse
from copy import deepcopy
import json
from pathlib import Path
import sys
from typing import Any, Callable, Mapping, TextIO


SUPPORTED_PROTOCOL_VERSIONS = (
    "2024-11-05",
    "2025-03-26",
    "2025-06-18",
    "2025-11-25",
)
LATEST_PROTOCOL_VERSION = SUPPORTED_PROTOCOL_VERSIONS[-1]
MAX_REQUEST_BYTES = 1_048_576
DEFAULT_STATE = Path.home() / "Library/Caches/outlook-local-mcp"
DEFAULT_FROM_DATE = "2000-01-01"
DEFAULT_REFRESH_SECONDS = 300
SERVER_NAME = "outlook-local-mcp"
SERVER_VERSION = "0.2.0"
MAX_QUERY_LENGTH = 10_000
MAX_ID_LENGTH = 512
MAX_SNAPSHOT_LENGTH = 128
MAX_CURSOR_LENGTH = 4_096
MAX_DATE_LENGTH = 10
_LOCAL_CACHE_NOTE = (
    " Local cache state may be refreshed or pruned, and attachment operations may create local exports; "
    "the original Outlook mailbox is never modified."
)


def _string(description: str, *, enum: tuple[str, ...] | None = None,
           default: str | None = None, max_length: int | None = None) -> dict[str, Any]:
    value: dict[str, Any] = {"type": "string", "description": description}
    if enum is not None:
        value["enum"] = list(enum)
    if default is not None:
        value["default"] = default
    if max_length is not None:
        value["maxLength"] = max_length
    return value


def _boolean(description: str, *, default: bool | None = None) -> dict[str, Any]:
    value: dict[str, Any] = {"type": "boolean", "description": description}
    if default is not None:
        value["default"] = default
    return value


def _integer(description: str, minimum: int, maximum: int, *, default: int | None = None) -> dict[str, Any]:
    value: dict[str, Any] = {
        "type": "integer",
        "description": description,
        "minimum": minimum,
        "maximum": maximum,
    }
    if default is not None:
        value["default"] = default
    return value


def _tool(name: str, description: str, properties: Mapping[str, dict[str, Any]],
          required: tuple[str, ...] = (), *, read_only_hint: bool = True,
          destructive_hint: bool = False) -> dict[str, Any]:
    if not read_only_hint:
        description += _LOCAL_CACHE_NOTE
    return {
        "name": name,
        "description": description,
        "inputSchema": {
            "type": "object",
            "properties": dict(properties),
            "required": list(required),
            "additionalProperties": False,
        },
        "annotations": {
            "readOnlyHint": read_only_hint,
            "destructiveHint": destructive_hint,
            "openWorldHint": False,
        },
    }


_SEARCH_PROPERTIES = {
    "query": _string(
        "Text to search for in Inbox and all descendants, Sent Items and its known descendants; Deleted is excluded. Empty searches the scoped mailbox.",
        max_length=MAX_QUERY_LENGTH,
    ),
    "person": _string("Optional literal name or email substring in verified sender/to/cc metadata; combined with query using AND. Missing participant evidence is reported, never guessed.", max_length=500),
    "person_role": _string("Which participant role to search.", enum=("any", "sender", "to", "cc"), default="any"),
    "field": _string("Search field.", enum=("subject", "sender", "preview", "all"), default="subject"),
    "from_date": _string(
        "Inclusive received date in YYYY-MM-DD form, interpreted in Asia/Singapore.",
        max_length=MAX_DATE_LENGTH,
    ),
    "to_date": _string(
        "Inclusive received date upper bound in YYYY-MM-DD form, interpreted in Asia/Singapore; the whole day is included.",
        max_length=MAX_DATE_LENGTH,
    ),
    "folder_id": _string(
        "Optional folder identifier within Inbox/Sent Items and their known descendants; Deleted is excluded.",
        max_length=MAX_ID_LENGTH,
    ),
    "include_subfolders": _boolean("Include descendants of folder_id.", default=True),
    "limit": _integer("Maximum results per page.", 1, 100, default=20),
    "cursor": _string("Cursor returned by a previous search.", max_length=MAX_CURSOR_LENGTH),
}


TOOL_SCHEMAS = (
    _tool(
        "outlook_local_status",
        "Return status for the read-only snapshot of Inbox and all descendants, Sent Items and its known descendants; Deleted is excluded. Cache may be incomplete; dates are received dates in Asia/Singapore.",
        {},
    ),
    _tool(
        "outlook_local_folders",
        "List Inbox and all descendants, Sent Items and its known descendants present in the local snapshot; Deleted is excluded. Cache may be incomplete.",
        {},
        read_only_hint=False,
        destructive_hint=True,
    ),
    _tool(
        "outlook_local_refresh",
        "Refresh the read-only snapshot for Inbox and all descendants, Sent Items and its known descendants; Deleted is excluded. Cache may be incomplete; from_date is a received date in Asia/Singapore.",
        {
            "from_date": _string(
                "Optional inclusive received date in YYYY-MM-DD form, interpreted in Asia/Singapore.",
                max_length=MAX_DATE_LENGTH,
            )
        },
        read_only_hint=False,
        destructive_hint=True,
    ),
    _tool(
        "outlook_local_search",
        "Search messages only in Inbox and all descendants, Sent Items and its known descendants; Deleted is excluded. Cache may be incomplete; date filters use the received date in Asia/Singapore and to_date is inclusive.",
        _SEARCH_PROPERTIES,
        read_only_hint=False,
        destructive_hint=True,
    ),
    _tool(
        "outlook_local_today",
        "List today's locally cached mail metadata from Inbox/Sent Items and known descendants; Deleted is excluded. Today uses verified last delivery time in Asia/Singapore. Missing receipt times are excluded and counted. Returns IDs, subjects, previews and body status; use read or messages_bundle for bodies. Cache may be incomplete. A cursor keeps the first page's date across midnight.",
        {
            "folder_id": _SEARCH_PROPERTIES["folder_id"],
            "include_subfolders": _SEARCH_PROPERTIES["include_subfolders"],
            "limit": _integer("Maximum results per page.", 1, 100, default=100),
            "cursor": _string("Cursor returned by outlook_local_today.", max_length=MAX_CURSOR_LENGTH),
        },
        read_only_hint=False,
        destructive_hint=True,
    ),
    _tool(
        "outlook_local_read",
        "Read one message from Inbox/Sent Items or a known descendant folder; Deleted is excluded. Cache may be incomplete; snapshot pins a stable read.",
        {
            "id": _string("Message identifier.", max_length=MAX_ID_LENGTH),
            "snapshot": _string("Optional snapshot identifier for a stable read.", max_length=MAX_SNAPSHOT_LENGTH),
            "body_max_chars": _integer("Maximum returned body characters.", 1, 100_000, default=20_000),
        },
        required=("id",),
        read_only_hint=False,
        destructive_hint=True,
    ),
    _tool(
        "outlook_local_digest",
        "Read compact summary material from the verified Inbox/Sent scope, filtered by received date. "
        "This does not generate an AI summary. Follow every next_cursor and collect page_ids; only after "
        "all pages succeed may the caller compare IDs with its previous run for removals. Each changed item "
        "has a version; reuse a saved summary only for unchanged versions. Bodies are deduplicated using "
        "verified full-text identity and returned prefix. Keep referenced bodies when supplying known_bodies. "
        "Missing/empty/truncated bodies stay explicit. Attachments are listed, not expanded. "
        "Refresh first when new mail must be checked immediately; otherwise the normal refresh interval applies.",
        {
            **_SEARCH_PROPERTIES,
            "limit": _integer("Maximum messages per page; reduce on source budget errors.", 1, 25, default=25),
            "body_max_chars": _integer("Maximum returned characters per body; truncation is explicit.", 1, 20_000, default=12_000),
            "known_versions": {
                "type": "array", "description": "Versions whose summary material the caller still retains. Never acknowledge failed or unsummarized items.",
                "items": {"type": "string", "maxLength": 64, "pattern": "^[a-f0-9]{64}$"},
                "minItems": 0, "maxItems": 5000,
            },
            "known_bodies": {
                "type": "array", "description": "Content IDs whose text the caller still retains, including earlier pages. Omit to receive all bodies needed by changed items.",
                "items": {"type": "string", "maxLength": 64, "pattern": "^[a-f0-9]{64}$"},
                "minItems": 0, "maxItems": 5000,
            },
        },
        read_only_hint=False,
        destructive_hint=True,
    ),
    _tool(
        "outlook_local_messages_bundle",
        "Read a bounded bundle of messages from Inbox/Sent Items and known descendant folders; Deleted is excluded. Cache may be incomplete; snapshot pins a stable read for the bundle.",
        {
            "ids": {
                "type": "array",
                "description": "Message identifiers, from 1 through 25 items.",
                "items": {"type": "string", "maxLength": MAX_ID_LENGTH},
                "minItems": 1,
                "maxItems": 25,
            },
            "snapshot": _string("Optional snapshot identifier for a stable read.", max_length=MAX_SNAPSHOT_LENGTH),
            "body_max_chars": _integer("Maximum body characters per message.", 1, 100_000, default=12_000),
            "text_budget_chars": _integer("Maximum combined text characters.", 1_000, 200_000, default=60_000),
        },
        required=("ids",),
        read_only_hint=False,
        destructive_hint=True,
    ),
    _tool(
        "outlook_local_attachments",
        "List attachments for one message from Inbox/Sent Items or a known descendant folder; Deleted is excluded. Cache may be incomplete; snapshot pins a stable read.",
        {
            "id": _string("Message identifier.", max_length=MAX_ID_LENGTH),
            "snapshot": _string("Optional snapshot identifier for a stable read.", max_length=MAX_SNAPSHOT_LENGTH),
        },
        required=("id",),
        read_only_hint=False,
        destructive_hint=True,
    ),
    _tool(
        "outlook_local_attachment",
        "Copy one cached attachment from a message in Inbox/Sent Items or a known descendant folder into this MCP server's own exports directory and return its local path. Deleted is excluded and the caller cannot choose its destination path. Cache may be incomplete; snapshot pins a stable read.",
        {
            "id": _string("Message identifier.", max_length=MAX_ID_LENGTH),
            "attachment_id": _string("Attachment identifier.", max_length=MAX_ID_LENGTH),
            "snapshot": _string("Optional snapshot identifier for a stable read.", max_length=MAX_SNAPSHOT_LENGTH),
        },
        required=("id", "attachment_id"),
        read_only_hint=False,
        destructive_hint=True,
    ),
)

TOOL_BY_NAME = {tool["name"]: tool for tool in TOOL_SCHEMAS}


class ValidationError(ValueError):
    """Raised when a JSON-RPC tool call has invalid arguments."""


def _type_name(value: Any) -> str:
    if value is None:
        return "null"
    if type(value) is bool:
        return "boolean"
    if type(value) is int:
        return "integer"
    if type(value) is float:
        return "number"
    if type(value) is str:
        return "string"
    if type(value) is list:
        return "array"
    if type(value) is dict:
        return "object"
    return type(value).__name__


def _matches_type(value: Any, expected: str) -> bool:
    # JSON booleans are not integers for this API, even though bool subclasses
    # int in Python.
    if expected == "boolean":
        return type(value) is bool
    if expected == "integer":
        return type(value) is int
    if expected == "string":
        return type(value) is str
    if expected == "array":
        return type(value) is list
    if expected == "object":
        return type(value) is dict
    return False


def validate_tool_arguments(name: str, arguments: Any) -> dict[str, Any]:
    """Validate and default one tool's arguments without calling the service."""

    if name not in TOOL_BY_NAME:
        raise ValidationError(f"unknown tool: {name}")
    if type(arguments) is not dict:
        raise ValidationError("arguments must be an object")
    schema = TOOL_BY_NAME[name]["inputSchema"]
    properties = schema["properties"]
    unknown = sorted(set(arguments) - set(properties))
    if unknown:
        raise ValidationError(f"unknown argument(s): {', '.join(unknown)}")
    missing = [field for field in schema["required"] if field not in arguments]
    if missing:
        raise ValidationError(f"missing required argument(s): {', '.join(missing)}")

    validated: dict[str, Any] = {}
    for field, field_schema in properties.items():
        if field not in arguments:
            if "default" in field_schema:
                validated[field] = field_schema["default"]
            continue
        value = arguments[field]
        expected = field_schema["type"]
        if not _matches_type(value, expected):
            raise ValidationError(
                f"argument {field!r} must be {expected}, got {_type_name(value)}"
            )
        if expected == "string" and "maxLength" in field_schema and len(value) > field_schema["maxLength"]:
            raise ValidationError(
                f"argument {field!r} exceeds maximum length {field_schema['maxLength']}"
            )
        if "enum" in field_schema and value not in field_schema["enum"]:
            choices = ", ".join(field_schema["enum"])
            raise ValidationError(f"argument {field!r} must be one of: {choices}")
        if expected == "integer":
            if value < field_schema["minimum"] or value > field_schema["maximum"]:
                raise ValidationError(
                    f"argument {field!r} must be {field_schema['minimum']}..{field_schema['maximum']}"
                )
        elif expected == "array":
            if len(value) < field_schema["minItems"] or len(value) > field_schema["maxItems"]:
                raise ValidationError(
                    f"argument {field!r} must contain {field_schema['minItems']}..{field_schema['maxItems']} items"
                )
            item_schema = field_schema.get("items")
            if item_schema:
                for index, item in enumerate(value):
                    if not _matches_type(item, item_schema["type"]):
                        raise ValidationError(
                            f"argument {field!r}[{index}] must be {item_schema['type']}, got {_type_name(item)}"
                        )
                    if item_schema["type"] == "string" and "maxLength" in item_schema and len(item) > item_schema["maxLength"]:
                        raise ValidationError(
                            f"argument {field!r}[{index}] exceeds maximum length {item_schema['maxLength']}"
                        )
        validated[field] = value
    if name == "outlook_local_digest":
        for field in ("known_versions", "known_bodies"):
            if any(len(token) != 64 or any(char not in "0123456789abcdef" for char in token)
                   for token in validated.get(field, [])):
                raise ValidationError(f"{field} must contain lowercase SHA-256 tokens")
    if name in {"outlook_local_search", "outlook_local_digest"}:
        if "person" in validated:
            validated["person"] = validated["person"].strip()
            if not validated["person"]:
                raise ValidationError("person must not be blank")
        if validated.get("person_role", "any") != "any" and not validated.get("person"):
            raise ValidationError("person_role requires person")
    return validated


def _response(request_id: Any, result: dict[str, Any]) -> dict[str, Any]:
    return {"jsonrpc": "2.0", "id": request_id, "result": result}


def _rpc_error(request_id: Any, code: int, message: str) -> dict[str, Any]:
    return {"jsonrpc": "2.0", "id": request_id, "error": {"code": code, "message": message}}


def _tool_error(error_type: str, message: str) -> dict[str, Any]:
    error = {"type": error_type, "message": message}
    return {
        "isError": True,
        "structuredContent": {"error": error},
        "content": [{"type": "text", "text": json.dumps({"error": error}, ensure_ascii=False)}],
    }


def _tool_success(value: dict[str, Any]) -> dict[str, Any]:
    return {
        "isError": False,
        "structuredContent": value,
        "content": [{"type": "text", "text": json.dumps(value, ensure_ascii=False)}],
    }


def _valid_request_id(value: Any) -> bool:
    if value is None or type(value) is str or type(value) is int:
        return True
    if type(value) is float:
        return value == value and value not in (float("inf"), float("-inf"))
    return False


class McpServer:
    """Protocol dispatcher with a lazily constructed read-only service."""

    def __init__(self, service: Any = None, *, service_factory: Callable[[], Any] | None = None,
                 stderr: TextIO | None = None):
        if service is not None and service_factory is not None:
            raise ValueError("provide service or service_factory, not both")
        self._service = service
        self._service_factory = service_factory
        self._stderr = stderr if stderr is not None else sys.stderr

    def _service_instance(self):
        if self._service is None:
            if self._service_factory is None:
                raise RuntimeError("Outlook service is not configured")
            self._service = self._service_factory()
        return self._service

    def _log(self, message: str) -> None:
        print(message, file=self._stderr, flush=True)

    def _call_tool(self, params: Any) -> dict[str, Any]:
        if type(params) is not dict:
            return _tool_error("invalid_params", "tools/call params must be an object")
        unknown = sorted(set(params) - {"name", "arguments", "_meta"})
        if unknown:
            return _tool_error("invalid_params", f"unknown tools/call field(s): {', '.join(unknown)}")
        name = params.get("name")
        if type(name) is not str:
            return _tool_error("invalid_params", "tools/call requires string name")
        try:
            arguments = validate_tool_arguments(name, params.get("arguments", {}))
        except ValidationError as error:
            return _tool_error("invalid_params", str(error))
        try:
            value = self._service_instance().call(name, arguments)
        except Exception as error:  # service errors stay inside a successful JSON-RPC response
            self._log(f"outlook_local service error: {type(error).__name__}: {error}")
            return _tool_error(type(error).__name__, str(error) or "service call failed")
        if type(value) is not dict:
            return _tool_error("invalid_service_result", "service.call must return an object")
        return _tool_success(value)

    def handle(self, request: Any) -> dict[str, Any] | None:
        """Dispatch one decoded JSON-RPC value; return None for notifications."""

        if type(request) is not dict:
            return _rpc_error(None, -32600, "Invalid Request")
        request_id = request.get("id")
        notification = "id" not in request
        if request.get("jsonrpc") != "2.0" or type(request.get("method")) is not str:
            return None if notification else _rpc_error(None, -32600, "Invalid Request")
        if "id" in request and not _valid_request_id(request_id):
            return _rpc_error(None, -32600, "Invalid Request")
        params = request.get("params", {})
        if type(params) is not dict:
            return None if notification else _rpc_error(request_id, -32600, "Invalid Request")

        method = request["method"]
        if method == "notifications/initialized":
            return None
        if method == "initialize":
            proposed = params.get("protocolVersion")
            version = proposed if proposed in SUPPORTED_PROTOCOL_VERSIONS else LATEST_PROTOCOL_VERSION
            result = {
                "protocolVersion": version,
                "capabilities": {"tools": {"listChanged": False}},
                "serverInfo": {"name": SERVER_NAME, "version": SERVER_VERSION},
            }
        elif method == "ping":
            result = {}
        elif method == "tools/list":
            result = {"tools": deepcopy(TOOL_SCHEMAS)}
        elif method == "tools/call":
            result = self._call_tool(params)
        else:
            return None if notification else _rpc_error(request_id, -32601, "Method not found")
        return None if notification else _response(request_id, result)


MCPServer = McpServer


def _line_bytes(value: str | bytes) -> int:
    if isinstance(value, bytes):
        content = value.rstrip(b"\r\n")
        return len(content)
    content = value.rstrip("\r\n")
    return len(content.encode("utf-8", errors="surrogatepass"))


def _terminated(value: str | bytes) -> bool:
    return value.endswith(b"\n") or value.endswith(b"\r") if isinstance(value, bytes) else value.endswith(("\n", "\r"))


def _readline_limited(stream: Any) -> tuple[str | bytes | None, bool]:
    """Return (line, oversized), draining an oversized NDJSON line."""

    line = stream.readline(MAX_REQUEST_BYTES + 2)
    if not line:
        return None, False
    oversized = _line_bytes(line) > MAX_REQUEST_BYTES
    if oversized and not _terminated(line):
        while True:
            chunk = stream.readline(MAX_REQUEST_BYTES + 2)
            if not chunk or _terminated(chunk):
                break
    return line, oversized


def _write_json(stream: Any, value: dict[str, Any]) -> None:
    # Escape all non-ASCII code points at the protocol boundary.  Besides
    # keeping the stdio transport independent of the stream's encoding, this
    # makes escaped but unpaired UTF-16 surrogates safe to write as JSON.
    text = json.dumps(value, ensure_ascii=True, separators=(",", ":")) + "\n"
    try:
        stream.write(text)
    except TypeError:
        stream.write(text.encode("utf-8"))
    stream.flush()


def run_stdio(stdin: Any = None, stdout: Any = None, stderr: TextIO | None = None,
              *, service: Any = None, service_factory: Callable[[], Any] | None = None) -> int:
    """Run the bounded NDJSON transport until stdin closes."""

    stdin = sys.stdin if stdin is None else stdin
    stdout = sys.stdout if stdout is None else stdout
    server = McpServer(service, service_factory=service_factory, stderr=stderr)
    while True:
        line, oversized = _readline_limited(stdin)
        if line is None:
            break
        if oversized:
            _write_json(stdout, _rpc_error(None, -32600, "Request exceeds 1 MiB"))
            continue
        try:
            request = json.loads(line)
        except (TypeError, json.JSONDecodeError, UnicodeDecodeError):
            _write_json(stdout, _rpc_error(None, -32700, "Parse error"))
            continue
        response = server.handle(request)
        if response is not None:
            _write_json(stdout, response)
    return 0


def _nonnegative_int(value: str) -> int:
    number = int(value)
    if number < 0:
        raise argparse.ArgumentTypeError("must be non-negative")
    return number


def _service_factory(state: str, profile: str | None, from_date: str, refresh_seconds: int) -> Callable[[], Any]:
    def create():
        # Keep the production import lazy: initialize and tools/list must be
        # available even if the service module is unavailable or expensive.
        from outlook_service import MailService

        return MailService(
            state=state,
            profile=profile,
            from_date=from_date,
            refresh_seconds=refresh_seconds,
        )

    return create


def main(argv: list[str] | None = None, *, stdin: Any = None, stdout: Any = None,
         stderr: TextIO | None = None, service: Any = None,
         service_factory: Callable[[], Any] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--state", default=str(DEFAULT_STATE))
    parser.add_argument("--profile")
    parser.add_argument("--from-date", default=DEFAULT_FROM_DATE)
    parser.add_argument("--refresh-seconds", type=_nonnegative_int, default=DEFAULT_REFRESH_SECONDS)
    args = parser.parse_args(argv)
    if service is None and service_factory is None:
        service_factory = _service_factory(args.state, args.profile, args.from_date, args.refresh_seconds)
    return run_stdio(
        stdin,
        stdout,
        stderr,
        service=service,
        service_factory=service_factory,
    )


if __name__ == "__main__":
    raise SystemExit(main())
