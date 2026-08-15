import os
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from reference_exporter import ReferenceExporter


class ReferenceExporterTests(unittest.TestCase):
    def setUp(self):
        self.environment = patch.dict(
            os.environ,
            {
                "SISC_API_URL": "https://sisc.example/api",
                "SISC_TOKEN": "",
                "ACTIONS_ID_TOKEN_REQUEST_URL": "https://token.actions.example",
            },
            clear=False,
        )
        self.environment.start()
        self.addCleanup(self.environment.stop)

    @patch("reference_exporter.time.sleep", return_value=None)
    @patch("reference_exporter.requests.post")
    def test_refreshes_oidc_and_retries_transient_response(self, post, _sleep):
        unavailable = Mock(status_code=502, text="Bad Gateway")
        accepted = Mock(status_code=200, text="")
        accepted.json.return_value = {
            "status": "COMPLETED",
            "records": 12,
            "coverage_years": [2026],
        }
        post.side_effect = [unavailable, accepted]
        exporter = ReferenceExporter()

        with patch.object(exporter, "_github_oidc_token", side_effect=["oidc-1", "oidc-2"]) as oidc:
            with patch.object(exporter, "_build_payload", return_value={"records": [{}]}):
                result = exporter.export_file(Path("HURTO PERSONAS.xlsx"), "2026-07-23")

        self.assertTrue(result)
        self.assertEqual(oidc.call_count, 2)
        self.assertEqual(post.call_count, 2)
        self.assertEqual(post.call_args_list[0].kwargs["headers"]["Authorization"], "Bearer oidc-1")
        self.assertEqual(post.call_args_list[1].kwargs["headers"]["Authorization"], "Bearer oidc-2")
        self.assertEqual(exporter.failures, [])

    @patch("reference_exporter.requests.post")
    def test_preserves_actionable_rejection_detail(self, post):
        post.return_value = Mock(status_code=422, text='{"detail":"periodo invalido"}')
        exporter = ReferenceExporter()

        with patch.object(exporter, "_github_oidc_token", return_value="oidc"):
            with patch.object(exporter, "_build_payload", return_value={"records": [{}]}):
                result = exporter.export_file(Path("EXTORSION.xlsx"), "2026-07-23")

        self.assertFalse(result)
        self.assertIn("HTTP 422", exporter.failures[0])
        self.assertIn("periodo invalido", exporter.failures[0])


if __name__ == "__main__":
    unittest.main()
