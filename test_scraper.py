import unittest
from unittest.mock import Mock, patch

from scraper import MinDefensaScraper


class MinDefensaScraperTests(unittest.TestCase):
    def test_asset_from_item_uses_top_level_metadata(self):
        item = {
            "id": "CONT123",
            "name": "HOMICIDIO INTENCIONAL.xlsx",
            "updatedDate": {"value": "2026-07-21T10:00:00Z", "timezone": "UTC"},
            "fields": {},
        }

        asset = MinDefensaScraper._asset_from_item(item)

        self.assertEqual(asset["nombre"], "HOMICIDIO INTENCIONAL.xlsx")
        self.assertEqual(asset["id"], "CONT123")
        self.assertIn("/CONT123/native", asset["url"])
        self.assertEqual(asset["fecha"]["value"], "2026-07-21T10:00:00Z")

    def test_discover_via_api_deduplicates_assets(self):
        response = Mock()
        response.raise_for_status.return_value = None
        response.json.return_value = {
            "hasMore": False,
            "items": [
                {
                    "id": "CONT123",
                    "name": "HOMICIDIO INTENCIONAL.xlsx",
                    "updatedDate": {"value": "2026-07-21T10:00:00Z"},
                    "fields": {},
                }
            ],
        }
        session = Mock()
        session.get.return_value = response

        with patch("scraper.requests.Session", return_value=session):
            scraper = MinDefensaScraper.__new__(MinDefensaScraper)
            scraper.cfg = {"umbrales": {"timeout_playwright": 1000}}
            assets = scraper._discover_via_api()

        self.assertEqual(len(assets), 1)
        self.assertEqual(session.get.call_count, len(scraper.FILE_CATEGORIES))


if __name__ == "__main__":
    unittest.main()
