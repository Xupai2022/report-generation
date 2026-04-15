from __future__ import annotations

import argparse
from datetime import date, datetime, time, timedelta, timezone
import json
from pathlib import Path

from mss_ai_ppt_sample_assets.backend import config

BEIJING_TZ = timezone(timedelta(hours=8))


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Export raw alarm/event data from the SOAR MongoDB source."
    )
    parser.add_argument(
        "--company-id",
        default=config.settings.soar_default_company_id,
        help="SOAR company_id filter",
    )
    parser.add_argument(
        "--date-range",
        default=config.settings.soar_default_date_range,
        help="Beijing date range in the form YYYY-MM-DD~YYYY-MM-DD",
    )
    parser.add_argument(
        "--start-time",
        help="Optional override for inclusive start time in ISO-8601 format",
    )
    parser.add_argument(
        "--end-time",
        help="Optional override for inclusive end time in ISO-8601 format",
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

    start_time, end_time = _resolve_time_range(args.date_range, args.start_time, args.end_time)
    request = MongoIngestionRequest(company_id=args.company_id, start_time=start_time, end_time=end_time)
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
    print(f"company_id={args.company_id}")
    print(f"start_time={request.start_time.isoformat()}")
    print(f"end_time={request.end_time.isoformat()}")
    print(f"output={output_path}")


def _resolve_time_range(
    date_range: str | None,
    start_time: str | None,
    end_time: str | None,
) -> tuple[datetime, datetime]:
    if start_time or end_time:
        if not (start_time and end_time):
            raise ValueError("start_time and end_time must be provided together")
        return _parse_datetime(start_time), _parse_datetime(end_time)

    if not date_range:
        raise ValueError("date_range is required when start_time/end_time are not provided")

    return _parse_beijing_date_range(date_range)


def _parse_beijing_date_range(date_range: str) -> tuple[datetime, datetime]:
    try:
        start_raw, end_raw = [part.strip() for part in date_range.split("~", 1)]
        start_day = datetime.strptime(start_raw, "%Y-%m-%d").date()
        end_day = datetime.strptime(end_raw, "%Y-%m-%d").date()
    except ValueError as exc:
        raise ValueError(
            "date_range must be in the form YYYY-MM-DD~YYYY-MM-DD"
        ) from exc

    if end_day < start_day:
        raise ValueError("date_range end day must be greater than or equal to start day")

    start_dt = datetime.combine(start_day, time.min, tzinfo=BEIJING_TZ)
    end_dt = datetime.combine(end_day, time.max, tzinfo=BEIJING_TZ)
    return start_dt, end_dt


def _parse_datetime(value: str) -> datetime:
    parsed = datetime.fromisoformat(value)
    if parsed.tzinfo is None:
        return parsed.replace(tzinfo=BEIJING_TZ)
    return parsed.astimezone(BEIJING_TZ)


def _json_default(value):
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    raise TypeError(f"Object of type {type(value).__name__} is not JSON serializable")


if __name__ == "__main__":
    main()
