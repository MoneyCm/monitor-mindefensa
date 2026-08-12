import json
import unittest
from datetime import datetime, timezone
from pathlib import Path

from sisc_heartbeat import build_payload, send_heartbeat


class MindefensaHeartbeatTests(unittest.TestCase):
    def test_builds_historical_payload_from_summary(self):
        summary = Path.cwd() / ".test-mindefensa-summary.json"
        state = Path.cwd() / ".test-mindefensa-state.json"
        self.addCleanup(summary.unlink, missing_ok=True)
        self.addCleanup(state.unlink, missing_ok=True)
        summary.write_text(
            json.dumps(
                {
                    "indicadores": [
                        {"delito": "Homicidios", "actual": 68, "ultimo_registro": "30/06/2026"},
                        {"delito": "Hurtos", "actual": 177, "ultimo_registro": "26/06/2026"},
                    ]
                }
            ),
            encoding="utf-8",
        )
        state.write_text(json.dumps({"archivos": {"HOMICIDIO.xlsx": {}}}), encoding="utf-8")
        payload = build_payload(
            summary,
            state,
            data_changed=True,
            now=datetime(2026, 8, 12, 12, 17, tzinfo=timezone.utc),
        )

        self.assertEqual(payload["connector_code"], "MINDEFENSA")
        self.assertEqual(payload["status"], "CURRENT")
        self.assertEqual(payload["source_cutoff_date"], "2026-06-30")
        self.assertEqual(payload["record_count"], 245)
        self.assertEqual(payload["details"]["source_role"], "historical-backup")

    def test_failure_and_missing_identity_fail_closed(self):
        payload = build_payload(outcome="failure")
        self.assertEqual(payload["status"], "ERROR")
        self.assertFalse(send_heartbeat(payload, service_key="", oidc_token=""))


if __name__ == "__main__":
    unittest.main()
