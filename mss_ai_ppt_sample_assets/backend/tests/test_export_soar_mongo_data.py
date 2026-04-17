from __future__ import annotations

from datetime import datetime
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from openpyxl import load_workbook

from mss_ai_ppt_sample_assets.backend.scripts.export_soar_mongo_data import _write_output


class ExportSoarMongoDataTestCase(unittest.TestCase):
    def test_write_output_creates_alarm_and_event_sheets(self) -> None:
        payload = {
            "meta": {"alarm_count": 1, "event_count": 1},
            "alarm": [
                {
                    "alarm_name": "test alarm",
                    "create_time": datetime(2026, 1, 30, 11, 2, 19),
                }
            ],
            "event": [
                {
                    "event_grading_tag": "重大事件",
                    "识别时长": 1.1,
                    "create_time": datetime(2026, 1, 30, 11, 2, 19),
                }
            ],
        }

        with TemporaryDirectory() as tmpdir:
            output_path = Path(tmpdir) / "soar_export.xlsx"
            _write_output(payload, output_path)

            workbook = load_workbook(output_path)
            self.assertEqual(workbook.sheetnames, ["告警表", "事件表"])
            self.assertEqual(workbook["告警表"]["A1"].value, "alarm_name")
            self.assertEqual(workbook["告警表"]["A2"].value, "test alarm")
            self.assertEqual(workbook["事件表"]["A1"].value, "event_grading_tag")
            self.assertEqual(workbook["事件表"]["B1"].value, "识别时长")
            self.assertEqual(workbook["事件表"]["B2"].value, 1.1)


if __name__ == "__main__":
    unittest.main()
