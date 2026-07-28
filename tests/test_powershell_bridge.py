from __future__ import annotations

import importlib.util
import io
import json
import queue
import sys
import threading
import unittest
from pathlib import Path
from unittest.mock import Mock, patch


SOURCE = (
    Path(__file__).resolve().parents[1]
    / "cross-platform"
    / "graph_app_role_manager.py"
)
SPEC = importlib.util.spec_from_file_location("graph_app_role_manager", SOURCE)
if SPEC is None or SPEC.loader is None:
    raise RuntimeError("Unable to load graph_app_role_manager.py")

MODULE = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = MODULE
SPEC.loader.exec_module(MODULE)


class PowerShellBridgeTests(unittest.TestCase):
    def setUp(self) -> None:
        self.events: list[tuple[str, object]] = []
        self.bridge = MODULE.PowerShellBridge(
            Path("/tmp/graph_backend.ps1"),
            lambda event, payload: self.events.append((event, payload)),
        )

    def test_call_sends_protocol_request_and_returns_data(self) -> None:
        stdin = io.StringIO()
        process = Mock()
        process.poll.return_value = None
        process.stdin = stdin
        self.bridge.process = process

        def respond() -> None:
            while not stdin.getvalue():
                pass
            request = json.loads(stdin.getvalue())
            with self.bridge.lock:
                response_queue = self.bridge.pending[request["requestId"]]
            response_queue.put(
                {
                    "ok": True,
                    "requestId": request["requestId"],
                    "data": {"backend": "PowerShell"},
                }
            )

        thread = threading.Thread(target=respond)
        thread.start()
        result = self.bridge.call("ping", timeout=2)
        thread.join(timeout=2)

        request = json.loads(stdin.getvalue())
        self.assertEqual(request["command"], "ping")
        self.assertEqual(result, {"backend": "PowerShell"})
        self.assertEqual(self.bridge.pending, {})

    def test_call_surfaces_backend_error(self) -> None:
        process = Mock()
        process.poll.return_value = None
        process.stdin = io.StringIO()
        self.bridge.process = process

        response_queue: queue.Queue[dict[str, object]] = queue.Queue()

        def queue_factory(*_args, **_kwargs):
            response_queue.put({"ok": False, "error": "Insufficient privileges"})
            return response_queue

        with patch.object(MODULE.queue, "Queue", side_effect=queue_factory):
            with self.assertRaisesRegex(RuntimeError, "Insufficient privileges"):
                self.bridge.call("assignRole", timeout=1)

    def test_find_pwsh_uses_system_path(self) -> None:
        with patch.object(MODULE.shutil, "which", return_value="/usr/local/bin/pwsh") as which:
            self.assertEqual(
                self.bridge.find_pwsh(),
                "/usr/local/bin/pwsh",
            )
        which.assert_called_once_with("pwsh")

    def test_backend_event_is_forwarded_without_completing_request(self) -> None:
        device_code = {
            "verificationUri": "https://login.microsoft.com/device",
            "userCode": "ABCD1234",
        }
        self.bridge.process = Mock(
            stdout=io.StringIO(
                f'{MODULE.PROTOCOL_PREFIX}'
                f'{{"ok":true,"event":"deviceCode","data":{json.dumps(device_code)}}}\n'
            )
        )

        self.bridge._read_stdout()

        self.assertEqual(self.events, [("deviceCode", device_code)])
        self.assertEqual(self.bridge.pending, {})

    def test_ready_event_marks_bridge_ready(self) -> None:
        self.bridge.process = Mock(
            stdout=io.StringIO(
                f'{MODULE.PROTOCOL_PREFIX}{{"ok":true,"event":"ready"}}\n'
            )
        )

        self.bridge._read_stdout()

        self.assertTrue(self.bridge.ready.is_set())

    def test_backend_path_is_resolved_next_to_gui(self) -> None:
        backend = SOURCE.with_name("graph_backend.ps1")
        self.assertTrue(backend.is_file())

    def test_windows_backend_uses_device_code_authentication(self) -> None:
        backend = SOURCE.with_name("graph_backend.ps1").read_text(encoding="utf-8")
        self.assertIn("if ($IsWindows)", backend)
        self.assertIn("$params.UseDeviceCode = $true", backend)
        self.assertIn("event = 'deviceCode'", backend)


if __name__ == "__main__":
    unittest.main()
