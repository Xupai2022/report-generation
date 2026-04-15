from __future__ import annotations

import argparse
from datetime import date, datetime
import json
from pathlib import Path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Export raw alarm/event data from the SOAR MongoDB source."
    )
    parser.add_argument("--company-id", required=True, help="SOAR company_id filter")
    parser.add_argument(
        "--start-time",
        required=True,
        help="Inclusive start time in ISO-8601 format, for example 2025-04-10T19:28:55+08:00",
    )
    parser.add_argument(
        "--end-time",
        required=True,
        help="Inclusive end time in ISO-8601 format, for example 2026-04-14T19:28:55+08:00",
    )
    parser.add_argument(
        "--output",
        default="mss_ai_ppt_sample_assets/backend/outputs/soar_raw_export.json",
        help="Output JSON file path",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    from mss_ai_ppt_sample_assets.backend.services.data_ingestion.collectors import (
        SOARMongoCollector,
    )
    from mss_ai_ppt_sample_assets.backend.services.data_ingestion.models import (
        MongoIngestionRequest,
    )

    request = MongoIngestionRequest(
        company_id=args.company_id,
        start_time=args.start_time,
        end_time=args.end_time,
    )
    collector = SOARMongoCollector()
    payload = collector.collect(request)

    output_path = Path(args.output).resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(
        json.dumps(payload, default=_json_default, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    print(f"alarm_count={payload['meta']['alarm_count']}")
    print(f"event_count={payload['meta']['event_count']}")
    print(f"output={output_path}")

def _json_default(value):
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    raise TypeError(f"Object of type {type(value).__name__} is not JSON serializable")


if __name__ == "__main__":
    main()
