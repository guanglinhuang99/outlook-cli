"""Minimum Python version shared by every macOS entry point."""
from __future__ import annotations

import json
import sys

MINIMUM = (3, 11)


def unsupported_message(version=None):
    """Return a user-facing error when this interpreter is too old, otherwise None."""
    current = tuple(sys.version_info[:2] if version is None else version[:2])
    if current >= MINIMUM:
        return None
    return (f"Python {MINIMUM[0]}.{MINIMUM[1]}+ is required, but this is Python {current[0]}.{current[1]}. "
            "Install Python 3.11 or newer; for the shared service set OUTLOOK_MAC_PYTHON to it and run restart, "
            "for the native MCP server rerun install_mcp.py with it.")


def answer_unsupported(stdin, stdout, message):
    """Answer each JSON-lines request with an explicit error instead of failing mid-operation."""
    while True:
        line = stdin.readline(1024 * 1024 + 1)
        if not line:
            return
        try:
            request = json.loads(line)
            request_id = request.get("id") if isinstance(request, dict) else None
        except ValueError:
            request_id = None
        stdout.write(json.dumps({"id": request_id, "error": {
            "code": "PYTHON_VERSION_UNSUPPORTED", "message": message}}) + "\n")
        stdout.flush()
