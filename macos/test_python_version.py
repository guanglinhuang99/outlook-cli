import io
import json
import unittest

import install_mcp
import python_version


class PythonVersionTests(unittest.TestCase):
    def test_minimum_is_accepted_and_older_versions_are_explained(self):
        self.assertIsNone(python_version.unsupported_message((3, 11, 0)))
        self.assertIsNone(python_version.unsupported_message((3, 13, 1)))
        message = python_version.unsupported_message((3, 9, 6))
        self.assertIn("Python 3.11+ is required, but this is Python 3.9", message)
        self.assertIn("OUTLOOK_MAC_PYTHON", message)

    def test_old_interpreter_answers_every_request_with_an_explicit_error(self):
        stdin = io.StringIO('{"id":1,"method":"health","params":{}}\nnot json\n{"id":7}\n')
        stdout = io.StringIO()
        python_version.answer_unsupported(stdin, stdout, "too old")
        responses = [json.loads(line) for line in stdout.getvalue().splitlines()]
        self.assertEqual([item["id"] for item in responses], [1, None, 7])
        self.assertTrue(all(item["error"] == {"code": "PYTHON_VERSION_UNSUPPORTED", "message": "too old"}
                            for item in responses))

    def test_native_installer_ships_the_version_guard(self):
        self.assertIn("python_version.py", install_mcp.REQUIRED_FILES)


if __name__ == "__main__":
    unittest.main()
