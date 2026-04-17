from __future__ import annotations

from datetime import datetime
import unittest

from mss_ai_ppt_sample_assets.backend.services.data_ingestion.transformers.event_transformer import (
    enrich_event_doc,
)


class EventTransformerTestCase(unittest.TestCase):
    def test_enrich_event_doc_adds_duration_features(self) -> None:
        event_doc = {
            "event_grading_tag": "重大事件",
            "create_time": datetime(2026, 1, 30, 11, 2, 19),
            "wechat_push_time": datetime(2026, 1, 30, 11, 3, 54),
            "dispose_time": datetime(2026, 1, 30, 11, 3, 51),
            "latest_time": datetime(2026, 3, 2, 9, 55, 9),
            "contain_time": datetime(2026, 3, 2, 9, 55, 9),
            "finished_time": datetime(2026, 3, 2, 9, 55, 9),
            "checkout_time": datetime(2026, 1, 30, 11, 1, 13),
        }

        enriched = enrich_event_doc(event_doc)

        self.assertEqual(enriched["识别时长"], 1.1)
        self.assertEqual(enriched["访问时长"], 1.58)
        self.assertEqual(enriched["遏制时间"], 1.58)
        self.assertEqual(enriched["处置时长"], 44572.83)
        self.assertEqual(enriched["闭环时长"], 44572.83)


if __name__ == "__main__":
    unittest.main()
