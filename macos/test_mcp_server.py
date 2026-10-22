import io
import json
from pathlib import Path
import subprocess
import sys
import unittest


MODULE_DIR = Path(__file__).parent
sys.path.insert(0, str(MODULE_DIR))
import mcp_server  # noqa: E402


def request(method, params=None, request_id=1):
    value = {"jsonrpc": "2.0", "id": request_id, "method": method}
    if params is not None:
        value["params"] = params
    return value


def lines(value):
    return [json.loads(line) for line in value.getvalue().splitlines()]


class FakeService:
    def __init__(self):
        self.calls = []

    def call(self, name, arguments):
        self.calls.append((name, arguments))
        return {"called": name, "arguments": arguments}


class ErrorService:
    def call(self, name, arguments):
        raise ValueError("fake domain failure")


class McpServerTests(unittest.TestCase):
    def test_tool_schemas_are_closed_and_annotate_local_cache_writes(self):
        expected = {
            "outlook_local_status",
            "outlook_local_folders",
            "outlook_local_refresh",
            "outlook_local_search",
            "outlook_local_today",
            "outlook_local_read",
            "outlook_local_messages_bundle",
            "outlook_local_digest",
            "outlook_local_attachments",
            "outlook_local_attachment",
        }
        actual = {tool["name"] for tool in mcp_server.TOOL_SCHEMAS}
        self.assertEqual(actual, expected)
        for tool in mcp_server.TOOL_SCHEMAS:
            self.assertFalse(tool["inputSchema"]["additionalProperties"])
            if tool["name"] == "outlook_local_status":
                self.assertEqual(tool["annotations"], {
                    "readOnlyHint": True,
                    "destructiveHint": False,
                    "openWorldHint": False,
                })
            else:
                self.assertEqual(tool["annotations"], {
                    "readOnlyHint": False,
                    "destructiveHint": True,
                    "openWorldHint": False,
                })
                self.assertIn("original Outlook mailbox is never modified", tool["description"])
        self.assertNotIn("output", str(mcp_server.TOOL_SCHEMAS))

    def test_initialize_negotiates_supported_versions_without_service_call(self):
        service = FakeService()
        server = mcp_server.McpServer(service)
        for version in mcp_server.SUPPORTED_PROTOCOL_VERSIONS:
            result = server.handle(request("initialize", {"protocolVersion": version}, version))
            self.assertEqual(result["result"]["protocolVersion"], version)
            self.assertEqual(result["result"]["serverInfo"]["name"], mcp_server.SERVER_NAME)
        unknown = server.handle(request("initialize", {"protocolVersion": "2099-01-01"}, 99))
        self.assertEqual(unknown["result"]["protocolVersion"], mcp_server.LATEST_PROTOCOL_VERSION)
        self.assertEqual(service.calls, [])

    def test_initialized_notification_ping_and_tools_list(self):
        server = mcp_server.McpServer(FakeService())
        self.assertIsNone(server.handle({"jsonrpc": "2.0", "method": "notifications/initialized"}))
        ping = server.handle(request("ping"))
        self.assertEqual(ping["result"], {})
        listing = server.handle(request("tools/list", {}))
        self.assertEqual(len(listing["result"]["tools"]), 10)
        self.assertEqual(
            {tool["name"] for tool in listing["result"]["tools"]},
            {tool["name"] for tool in mcp_server.TOOL_SCHEMAS},
        )

    def test_tools_call_passes_complete_name_and_defaults(self):
        service = FakeService()
        server = mcp_server.McpServer(service)
        response = server.handle(
            request("tools/call", {"name": "outlook_local_search", "arguments": {}})
        )
        self.assertFalse(response["result"]["isError"])
        self.assertEqual(service.calls, [
            (
                "outlook_local_search",
                {"field": "subject", "include_subfolders": True, "limit": 20, "person_role": "any"},
            )
        ])

    def test_tools_call_forwards_all_explicit_arguments(self):
        service = FakeService()
        server = mcp_server.McpServer(service)
        calls = (
            (
                "outlook_local_refresh",
                {"from_date": "2026-10-20"},
            ),
            (
                "outlook_local_search",
                {
                    "query": "合规🧪",
                    "field": "all",
                    "person": "张三",
                    "person_role": "cc",
                    "from_date": "2026-10-01",
                    "to_date": "2026-10-20",
                    "folder_id": "sent/child",
                    "include_subfolders": False,
                    "limit": 7,
                    "cursor": "cursor-1",
                },
            ),
            (
                "outlook_local_today",
                {"folder_id": "inbox", "include_subfolders": False, "limit": 3, "cursor": "today-cursor"},
            ),
            (
                "outlook_local_read",
                {"id": "message-1", "snapshot": "snapshot-1", "body_max_chars": 321},
            ),
            (
                "outlook_local_messages_bundle",
                {
                    "ids": ["message-1", "message-2"],
                    "snapshot": "snapshot-1",
                    "body_max_chars": 654,
                    "text_budget_chars": 4321,
                },
            ),
            (
                "outlook_local_attachments",
                {"id": "message-1", "snapshot": "snapshot-1"},
            ),
            (
                "outlook_local_attachment",
                {
                    "id": "message-1",
                    "attachment_id": "attachment-1",
                    "snapshot": "snapshot-1",
                },
            ),
        )
        for name, arguments in calls:
            response = server.handle(
                request("tools/call", {"name": name, "arguments": arguments})
            )
            self.assertFalse(response["result"]["isError"], name)
        self.assertEqual(service.calls, list(calls))

    def test_person_arguments_reject_blank_and_role_without_person(self):
        for arguments in ({"person": " "}, {"person_role": "to"},
                          {"person": "a", "person_role": "bcc"}):
            with self.assertRaises(mcp_server.ValidationError):
                mcp_server.validate_tool_arguments("outlook_local_search", arguments)

    def test_today_defaults_and_rejects_date_override(self):
        service = FakeService()
        server = mcp_server.McpServer(service)
        result = server.handle(request("tools/call", {"name": "outlook_local_today"}))["result"]
        self.assertFalse(result["isError"])
        self.assertEqual(service.calls, [("outlook_local_today", {"include_subfolders": True, "limit": 100})])
        for arguments in ({"from_date": "2000-01-01"}, {"limit": True}, {"limit": 101}):
            result = server.handle(request("tools/call", {"name": "outlook_local_today", "arguments": arguments}))["result"]
            self.assertTrue(result["isError"])
        self.assertEqual(len(service.calls), 1)

    def test_digest_bounds_and_incremental_tokens(self):
        validate = lambda args: mcp_server.validate_tool_arguments("outlook_local_digest", args)
        result = validate({"known_versions": ["a" * 64], "known_bodies": []})
        self.assertEqual(result["limit"], 25)
        self.assertEqual(result["body_max_chars"], 12000)
        for args in ({"limit": 26}, {"limit": True}, {"body_max_chars": 20001},
                     {"known_versions": ["a" * 64] * 5001}, {"known_bodies": [None]},
                     {"known_versions": ["x" * 64]}, {"known_versions": ["a" * 63]},
                     {"known_versions": ["A" * 64]}, {"person_role": "to"}, {"person": " "}):
            with self.subTest(args=str(args)[:90]), self.assertRaises(mcp_server.ValidationError):
                validate(args)

    def test_tools_call_rejects_unknown_arguments_and_bool_as_integer(self):
        service = FakeService()
        server = mcp_server.McpServer(service)
        cases = (
            {"name": "outlook_local_status", "arguments": {"output": "/tmp/x"}},
            {"name": "outlook_local_search", "arguments": {"limit": True}},
            {"name": "outlook_local_messages_bundle", "arguments": {"ids": ["a"], "text_budget_chars": 1}},
            {"name": "outlook_local_nope", "arguments": {}},
        )
        for params in cases:
            result = server.handle(request("tools/call", params))["result"]
            self.assertTrue(result["isError"])
            self.assertIn("error", result["structuredContent"])
        self.assertEqual(service.calls, [])

    def test_tools_call_validates_required_and_bounded_arrays(self):
        server = mcp_server.McpServer(FakeService())
        missing = server.handle(request("tools/call", {"name": "outlook_local_read", "arguments": {}}))
        self.assertTrue(missing["result"]["isError"])
        too_many = server.handle(
            request(
                "tools/call",
                {"name": "outlook_local_messages_bundle", "arguments": {"ids": [str(i) for i in range(26)]}},
            )
        )
        self.assertTrue(too_many["result"]["isError"])

    def test_service_domain_error_stays_in_tool_result(self):
        response = mcp_server.McpServer(ErrorService(), stderr=io.StringIO()).handle(
            request("tools/call", {"name": "outlook_local_status", "arguments": {}})
        )
        result = response["result"]
        self.assertTrue(result["isError"])
        self.assertEqual(result["structuredContent"]["error"]["type"], "ValueError")
        self.assertEqual(result["structuredContent"]["error"]["message"], "fake domain failure")

    def test_json_rpc_errors_and_notifications(self):
        output = io.StringIO()
        mcp_server.run_stdio(
            io.StringIO("not-json\n[]\n" + json.dumps(request("ping")) + "\n"),
            output,
            io.StringIO(),
            service=FakeService(),
        )
        responses = lines(output)
        self.assertEqual([response["error"]["code"] for response in responses[:2]], [-32700, -32600])
        self.assertEqual(responses[2]["result"], {})
        self.assertNotIn("result", responses[0])

        unknown = mcp_server.McpServer(FakeService()).handle(request("no/such/method"))
        self.assertEqual(unknown["error"]["code"], -32601)

    def test_oversized_line_is_rejected_and_next_request_is_read(self):
        oversized = "x" * (mcp_server.MAX_REQUEST_BYTES + 1)
        output = io.StringIO()
        mcp_server.run_stdio(
            io.StringIO(oversized + "\n" + json.dumps(request("ping")) + "\n"),
            output,
            io.StringIO(),
            service=FakeService(),
        )
        responses = lines(output)
        self.assertEqual(responses[0]["error"]["code"], -32600)
        self.assertEqual(responses[1]["result"], {})

    def test_stdio_escapes_unpaired_surrogate_and_keeps_next_request_working(self):
        payload = (
            json.dumps(
                request(
                    "tools/call",
                    {
                        "name": "outlook_local_search",
                        "arguments": {"query": "中文\ud800"},
                    },
                ),
                ensure_ascii=True,
            )
            + "\n"
            + json.dumps(request("ping"))
            + "\n"
        )
        output = io.StringIO()
        mcp_server.run_stdio(
            io.StringIO(payload),
            output,
            io.StringIO(),
            service=FakeService(),
        )
        responses = lines(output)
        self.assertEqual(len(responses), 2)
        self.assertFalse(responses[0]["result"]["isError"])
        self.assertEqual(
            responses[0]["result"]["structuredContent"]["arguments"]["query"],
            "中文\ud800",
        )
        self.assertEqual(responses[1]["result"], {})

    def test_main_accepts_cli_config_and_constructs_service_lazily(self):
        created = []

        def factory():
            created.append(True)
            return FakeService()

        output = io.StringIO()
        code = mcp_server.main(
            [
                "--state", "/tmp/independent-outlook-state",
                "--profile", "/tmp/profile",
                "--from-date", "2000-01-01",
                "--refresh-seconds", "7",
            ],
            stdin=io.StringIO(json.dumps(request("initialize")) + "\n"),
            stdout=output,
            stderr=io.StringIO(),
            service_factory=factory,
        )
        self.assertEqual(code, 0)
        self.assertEqual(created, [])
        self.assertEqual(lines(output)[0]["result"]["protocolVersion"], mcp_server.LATEST_PROTOCOL_VERSION)

    def test_real_subprocess_entry_helper_has_protocol_only_stdout(self):
        code = (
            "import mcp_server\n"
            "class S:\n"
            "    def call(self, name, arguments):\n"
            "        return {'name': name, 'arguments': arguments}\n"
            "mcp_server.run_stdio(service=S())\n"
        )
        payload = (
            json.dumps(request("initialize", {"protocolVersion": "2025-11-25"}))
            + "\n"
            + json.dumps(request("tools/call", {"name": "outlook_local_status", "arguments": {}}))
            + "\n"
        )
        completed = subprocess.run(
            [sys.executable, "-c", code],
            cwd=MODULE_DIR,
            input=payload,
            text=True,
            capture_output=True,
            check=False,
        )
        self.assertEqual(completed.returncode, 0, completed.stderr)
        self.assertEqual(completed.stderr, "")
        output = [json.loads(line) for line in completed.stdout.splitlines()]
        self.assertEqual(len(output), 2)
        self.assertEqual(output[0]["result"]["protocolVersion"], "2025-11-25")
        self.assertEqual(output[1]["result"]["structuredContent"]["name"], "outlook_local_status")


if __name__ == "__main__":
    unittest.main()
