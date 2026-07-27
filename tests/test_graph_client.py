from __future__ import annotations

import importlib.util
import sys
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


class GraphClientTests(unittest.TestCase):
    def setUp(self) -> None:
        self.client = MODULE.GraphClient()
        self.client.access_token = "test-token-not-a-secret"
        self.client.graph_sp = {"id": "graph-resource-id"}

    def test_assign_app_role_uses_target_and_graph_resource(self) -> None:
        response = Mock(ok=True, content=b"{}", json=lambda: {})

        with patch.object(MODULE.requests, "post", return_value=response) as post:
            self.client.assign_app_role("target-id", "role-id")

        self.assertTrue(
            post.call_args.args[0].endswith(
                "/servicePrincipals/target-id/appRoleAssignments"
            )
        )
        self.assertEqual(
            post.call_args.kwargs["json"],
            {
                "principalId": "target-id",
                "resourceId": "graph-resource-id",
                "appRoleId": "role-id",
            },
        )

    def test_remove_app_role_assignment_uses_assignment_id(self) -> None:
        response = Mock(ok=True, content=b"")

        with patch.object(MODULE.requests, "delete", return_value=response) as delete:
            self.client.remove_app_role_assignment("target-id", "assignment-id")

        self.assertTrue(
            delete.call_args.args[0].endswith(
                "/servicePrincipals/target-id/appRoleAssignments/assignment-id"
            )
        )
        self.assertEqual(delete.call_args.kwargs["timeout"], 60)

    def test_graph_error_does_not_expose_request_headers(self) -> None:
        response = Mock(
            ok=False,
            status_code=403,
            text="Forbidden",
            json=lambda: {"error": {"message": "Insufficient privileges"}},
        )

        with self.assertRaisesRegex(
            RuntimeError,
            "Microsoft Graph HTTP 403: Insufficient privileges",
        ):
            self.client._raise_for_graph_error(response)


if __name__ == "__main__":
    unittest.main()
