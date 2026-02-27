"""Excel file upload and processing module.

This module handles:
- File validation (format, size, MIME type)
- Excel data extraction to JSON
- Session-based storage for concurrent requests

Current parser target:
- `data.xlsx` workbook layout with sheet `数据统计`
"""

from __future__ import annotations

import json
import logging
import re
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, List, Set

import openpyxl
from openpyxl.utils.exceptions import InvalidFileException

from mss_ai_ppt_sample_assets.backend.exceptions import (
    DataValidationError,
    FileValidationError,
)

logger = logging.getLogger(__name__)


class ExcelValidator:
    """Validates Excel file uploads for security and format compliance."""

    ALLOWED_EXTENSIONS: Set[str] = {".xlsx"}
    FORBIDDEN_EXTENSIONS: Set[str] = {".xlsm", ".xls", ".xlsb", ".csv"}
    ALLOWED_MIME_TYPES: Set[str] = {
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "application/zip",
    }

    FORMAT_TIPS = {
        ".xlsm": "This format may contain macros. Please convert to .xlsx.",
        ".xls": "Legacy Excel format. Please save as .xlsx.",
        ".xlsb": "Binary Excel format. Please save as .xlsx.",
        ".csv": "CSV is not supported. Please save as .xlsx.",
    }

    def __init__(self, max_size_mb: int = 10):
        self.max_size_bytes = max_size_mb * 1024 * 1024

    def validate_extension(self, filename: str) -> None:
        file_ext = Path(filename).suffix.lower()
        if file_ext not in self.ALLOWED_EXTENSIONS:
            if file_ext in self.FORBIDDEN_EXTENSIONS:
                detail_msg = self.FORMAT_TIPS.get(file_ext, "Please convert to .xlsx.")
                raise FileValidationError(
                    filename=filename,
                    reason=f"Forbidden file extension {file_ext}: {detail_msg}",
                )
            raise FileValidationError(
                filename=filename,
                reason=f"Unsupported file extension {file_ext}, only .xlsx is allowed.",
            )

    def validate_mime_type(self, filename: str, content_type: str) -> None:
        if content_type not in self.ALLOWED_MIME_TYPES:
            logger.warning("MIME mismatch: %s - %s", filename, content_type)
            raise FileValidationError(
                filename=filename,
                reason=f"MIME verification failed: got {content_type}, expected .xlsx MIME.",
            )

    def validate_size(self, filename: str, size_bytes: int) -> None:
        if size_bytes > self.max_size_bytes:
            max_mb = self.max_size_bytes / 1024 / 1024
            actual_mb = round(size_bytes / 1024 / 1024, 2)
            raise FileValidationError(
                filename=filename,
                reason=f"File size {actual_mb}MB exceeds {max_mb}MB limit.",
            )


class ExcelDataExtractor:
    """Extracts structured data from Excel files."""

    CLASSIC_REQUIRED_SHEET = "数据统计"

    @staticmethod
    def _unwrap_cell(value: Any) -> tuple[Any, str]:
        if hasattr(value, "value") and hasattr(value, "number_format"):
            return value.value, str(value.number_format or "")
        return value, ""

    @staticmethod
    def _has_value(value: Any) -> bool:
        raw_value, _ = ExcelDataExtractor._unwrap_cell(value)
        if raw_value is None:
            return False
        if isinstance(raw_value, str):
            return bool(raw_value.strip())
        return True

    @staticmethod
    def _decimal_places_from_number_format(number_format: str) -> int | None:
        if not number_format:
            return None

        fmt = number_format.split(";")[0].strip()
        if not fmt or fmt.lower() == "general":
            return None

        # Remove literal/text parts and Excel fill/alignment directives.
        fmt = re.sub(r'"[^"]*"', "", fmt)
        fmt = re.sub(r"\\.", "", fmt)
        fmt = re.sub(r"_.", "", fmt)
        fmt = re.sub(r"\*.", "", fmt)

        decimal_match = re.search(r"\.([0#]+)", fmt)
        if decimal_match:
            return len(decimal_match.group(1))
        if re.search(r"[0#]", fmt):
            return 0
        return None

    @staticmethod
    def _format_numeric_for_text(value: float, number_format: str) -> str:
        decimals = ExcelDataExtractor._decimal_places_from_number_format(number_format)
        is_percent = "%" in number_format
        normalized = ExcelDataExtractor._normalize_number(value)

        if is_percent:
            if decimals is None:
                decimals = 2
            decimals = min(decimals, 2)
            percent_value = ExcelDataExtractor._normalize_number(float(normalized) * 100)
            if isinstance(percent_value, int):
                return f"{percent_value}%"
            return f"{percent_value:.{decimals}f}%"

        if decimals is not None:
            decimals = min(decimals, 2)
            if decimals == 0:
                return str(int(round(float(normalized))))
            return f"{float(normalized):.{decimals}f}"

        if isinstance(normalized, int):
            return str(normalized)
        return f"{float(normalized):.2f}".rstrip("0").rstrip(".")

    @staticmethod
    def _normalize_number(value: Any) -> Any:
        """Normalize numeric value: keep integers, round non-integers to max 2 decimals."""
        if isinstance(value, bool):
            return value
        if not isinstance(value, (int, float)):
            return value

        rounded = round(float(value), 2)
        if rounded.is_integer():
            return int(rounded)
        return rounded

    @staticmethod
    def _format_numbers_for_output(value: Any) -> Any:
        """Recursively format numeric payload for JSON output.

        Rule:
        - Integers remain numeric.
        - Non-integer floats are rendered as fixed 2-decimal strings.
        """
        if isinstance(value, dict):
            return {k: ExcelDataExtractor._format_numbers_for_output(v) for k, v in value.items()}
        if isinstance(value, list):
            return [ExcelDataExtractor._format_numbers_for_output(v) for v in value]
        if isinstance(value, bool):
            return value
        if isinstance(value, int):
            return value
        if isinstance(value, float):
            rounded = round(value, 2)
            if rounded.is_integer():
                return int(rounded)
            return f"{rounded:.2f}"
        return value

    @staticmethod
    def _to_text(value: Any, default: str = "") -> str:
        raw_value, number_format = ExcelDataExtractor._unwrap_cell(value)

        if raw_value is None:
            return default
        if isinstance(raw_value, (datetime, date)):
            return raw_value.strftime("%Y-%m-%d")
        if isinstance(raw_value, (int, float)) and not isinstance(raw_value, bool):
            return ExcelDataExtractor._format_numeric_for_text(float(raw_value), number_format)

        text = str(raw_value).strip()
        return text if text else default

    @staticmethod
    def _to_number(value: Any, default: float = 0) -> float:
        raw_value, _ = ExcelDataExtractor._unwrap_cell(value)

        if raw_value is None:
            return default
        if isinstance(raw_value, (int, float)):
            return ExcelDataExtractor._normalize_number(raw_value)

        text = str(raw_value).strip()
        if text in {"", "None", "#DIV/0!", "#N/A"}:
            return default

        if text.endswith("%"):
            try:
                return ExcelDataExtractor._normalize_number(float(text[:-1]) / 100)
            except Exception:
                return default

        try:
            if "." in text:
                return ExcelDataExtractor._normalize_number(float(text))
            return ExcelDataExtractor._normalize_number(float(int(text)))
        except Exception:
            return default

    @staticmethod
    def _to_pct(value: Any) -> str:
        raw_value, number_format = ExcelDataExtractor._unwrap_cell(value)
        if isinstance(raw_value, (int, float)) and "%" in number_format:
            decimals = ExcelDataExtractor._decimal_places_from_number_format(number_format)
            if decimals is None:
                decimals = 2
            return f"{raw_value * 100:.{decimals}f}%"

        n = ExcelDataExtractor._to_number(value, 0)
        if n <= 1:
            return f"{round(n * 100, 2)}%"
        return f"{round(n, 2)}%"

    @staticmethod
    def _put_text(target: Dict[str, Any], key: str, raw_value: Any) -> None:
        if ExcelDataExtractor._has_value(raw_value):
            target[key] = ExcelDataExtractor._to_text(raw_value)

    @staticmethod
    def _put_pct(target: Dict[str, Any], key: str, raw_value: Any) -> None:
        if ExcelDataExtractor._has_value(raw_value):
            target[key] = ExcelDataExtractor._to_pct(raw_value)

    @staticmethod
    def _add_section(root: Dict[str, Any], section_key: str, section: Dict[str, Any]) -> None:
        if section:
            root[section_key] = section

    @staticmethod
    def _read_month_columns(ws, header_row: int, col_start: int, col_end: int) -> List[int]:
        cols: List[int] = []
        for c in range(col_start, col_end + 1):
            if ExcelDataExtractor._has_value(ws.cell(header_row, c).value):
                cols.append(c)
        return cols

    @staticmethod
    def _read_labeled_pairs(ws, start_row: int, end_row: int, label_col: int, value_col: int) -> Dict[str, List[Any]]:
        labels: List[str] = []
        values: List[float] = []
        for r in range(start_row, end_row + 1):
            raw_label = ws.cell(r, label_col).value
            raw_value = ws.cell(r, value_col).value
            if not ExcelDataExtractor._has_value(raw_label) and not ExcelDataExtractor._has_value(raw_value):
                continue
            labels.append(ExcelDataExtractor._to_text(raw_label))
            values.append(ExcelDataExtractor._to_number(raw_value, 0))
        return {"labels": labels, "values": values}

    @staticmethod
    def _remove_ai_generated_fields(data: Dict[str, Any]) -> Dict[str, Any]:
        """Remove keys mapped to ai_generate placeholders from classic descriptor."""
        descriptor_dir = Path(__file__).resolve().parent.parent / "data" / "templates"
        descriptor = None

        for path in descriptor_dir.glob("*_descriptor.json"):
            try:
                content = json.loads(path.read_text(encoding="utf-8"))
            except Exception:
                continue
            if content.get("template_id") == "mss_classic_ops":
                descriptor = content
                break

        if not descriptor:
            return data

        for slide in descriptor.get("slides", []):
            slide_key = slide.get("slide_key")
            if not slide_key:
                continue

            section = data.get(slide_key)
            if not isinstance(section, dict):
                continue

            ai_tokens = {
                ph.get("token")
                for ph in slide.get("placeholders", [])
                if ph.get("ai_generate") and ph.get("token")
            }
            for token in ai_tokens:
                section.pop(token, None)

            if not section:
                data.pop(slide_key, None)

        return data

    @staticmethod
    def _extract_classic_ops(ws) -> Dict[str, Any]:
        output: Dict[str, Any] = {
            "schema_version": "1.0",
            "template_id": "mss_classic_ops",
        }

        raw_period_start = ws["L1"]
        raw_period_end = ws["M1"]
        period: Dict[str, Any] = {}

        if ExcelDataExtractor._has_value(raw_period_start):
            start = ExcelDataExtractor._to_text(raw_period_start)
            period["start"] = start
            period["start_month"] = start[:7] if len(start) >= 7 else start

        if ExcelDataExtractor._has_value(raw_period_end):
            end = ExcelDataExtractor._to_text(raw_period_end)
            period["end"] = end
            period["end_month"] = end[:7] if len(end) >= 7 else end

        if period:
            output["period"] = period

        cover: Dict[str, Any] = {}
        if period.get("start"):
            cover["PERIOD_start"] = period["start"]
        if period.get("end"):
            cover["PERIOD_end"] = period["end"]
        ExcelDataExtractor._add_section(output, "cover", cover)

        architecture: Dict[str, Any] = {}
        ExcelDataExtractor._put_text(architecture, "AF_count", ws["D3"])
        ExcelDataExtractor._put_text(architecture, "STA_count", ws["D4"])
        ExcelDataExtractor._put_text(architecture, "EDR_count", ws["D5"])
        ExcelDataExtractor._put_text(architecture, "TSS_count", ws["D6"])
        ExcelDataExtractor._add_section(output, "Architecture", architecture)

        deliverables: Dict[str, Any] = {}
        deliverables_map = {
            "kickoff_ppt": "D9",
            "first_analysis": "D10",
            "vuln_evidence": "D11",
            "vuln_list": "D12",
            "asset_inventory": "D13",
            "incident_tracking": "D14",
            "incident_response": "D15",
            "attack_surface": "D16",
            "sec_ops_weekly": "G9",
            "analysis_monthly": "G10",
            "sec_ops_quarterly": "G11",
            "nationalday": "G13",
            "springfestival": "G14",
            "mayday": "G15",
        }
        for token, addr in deliverables_map.items():
            ExcelDataExtractor._put_text(deliverables, token, ws[addr])
        # P7: primary mapping is G16/G17. Fallback to G17/G18 for legacy-filled sheets.
        if ExcelDataExtractor._has_value(ws["G16"]):
            ExcelDataExtractor._put_text(deliverables, "biweekly_threat", ws["G16"])
        elif ExcelDataExtractor._has_value(ws["G17"]):
            ExcelDataExtractor._put_text(deliverables, "biweekly_threat", ws["G17"])

        if ExcelDataExtractor._has_value(ws["G18"]):
            ExcelDataExtractor._put_text(deliverables, "phishing_poster", ws["G18"])
        elif ExcelDataExtractor._has_value(ws["G17"]):
            ExcelDataExtractor._put_text(deliverables, "phishing_poster", ws["G17"])
        ExcelDataExtractor._add_section(output, "deliverables", deliverables)

        coverage_summary: Dict[str, Any] = {}
        ExcelDataExtractor._put_pct(coverage_summary, "AF_coverage", ws["L22"])
        ExcelDataExtractor._put_pct(coverage_summary, "aes_coverage", ws["L24"])
        ExcelDataExtractor._put_pct(coverage_summary, "probe_coverage", ws["L23"])
        coverage_map = {
            "cybersecurity_incident": "C21",
            "proactive_protection": "D21",
            "incident_count1": "E21",
            "closure_rate1": "F21",
            "incident_count2": "G21",
            "closure_rate2": "H21",
            "response_time": "I21",
            "alert_manual": "L26",
            "alert_auto": "L25",
            "mss_risk": "J27",
            "nonmss_risk": "J28",
            "alert_inner": "J23",
            "incident_inner": "J24",
            "handling_inner": "J25",
            "response_inner": "J26",
            "alert_outer": "J28",
            "incident_outer": "J29",
            "handling_outer": "J30",
            "response_outer": "J31",
            "business_system": "D28",
            "server_asset": "D29",
            "pc_asset": "D30",
            "log_count": "G22",
            "log_rate": "G23",
            "alert_count": "G24",
            "alert_rate": "G25",
            "incident_count": "G26",
            "asset_count": "G27",
            "AF_count": "D3",
            "STA_count": "D4",
            "EDR_count": "D5",
            "TSS_count": "D6",
        }
        for token, addr in coverage_map.items():
            ExcelDataExtractor._put_text(coverage_summary, token, ws[addr])
        ExcelDataExtractor._add_section(output, "coverage_summary", coverage_summary)

        protection_overview: Dict[str, Any] = {}
        protection_map = {
            "af_block": "D35",
            "edr_risk": "D36",
            "policy_check": "D37",
            "vuln_protect": "D38",
            "surface": "D39",
            "scanning": "D40",
            "server": "D41",
            "pc": "D42",
            "vuln_high": "D43",
            "surface_risk": "D44",
            "incident_closed": "D45",
            "vuln_closed": "D46",
            "weekly": "D47",
            "monthly": "D48",
            "alert_judgment": "D34",
            "alert_response": "E34",
            "threat_contain": "F34",
            "xdr_log": "G35",
            "xdr_alert": "G36",
            "xdr_incident": "G37",
            "mss_push": "G38",
            "mss_latency": "G39",
            "xdr_auto": "G40",
            "mss_handle": "G41",
            "emergency_handle": "G42",
            "incident_contain": "G43",
        }
        for token, addr in protection_map.items():
            ExcelDataExtractor._put_text(protection_overview, token, ws[addr])
        if "xdr_auto" not in protection_overview and ExcelDataExtractor._has_value(ws["D124"]):
            ExcelDataExtractor._put_text(protection_overview, "xdr_auto", ws["D124"])
        ExcelDataExtractor._add_section(output, "protection_overview", protection_overview)

        incident_effectiveness: Dict[str, Any] = {}
        incident_map = {
            "incident_total": "C51",
            "response_avg": "D51",
            "handle_avg": "E51",
            "incident_closed": "F51",
            "trust_assurance": "F51",
            "security_trust": "D51",
        }
        for token, addr in incident_map.items():
            ExcelDataExtractor._put_text(incident_effectiveness, token, ws[addr])

        response_timeliness = ExcelDataExtractor._read_labeled_pairs(ws, 53, 57, 6, 7)
        if response_timeliness["labels"]:
            incident_effectiveness["response_timeliness"] = response_timeliness
            incident_effectiveness["P11_bar"] = response_timeliness

        trend_cols = ExcelDataExtractor._read_month_columns(ws, 60, 3, 12)
        if trend_cols:
            response_trend_months = [ExcelDataExtractor._to_text(ws.cell(60, c)) for c in trend_cols]
            response_trend_values = [
                ExcelDataExtractor._to_number(ws.cell(61, c).value, 0) for c in trend_cols
            ]
            response_trend = {
                "months": response_trend_months,
                "avg_response_minutes": response_trend_values,
            }
            incident_effectiveness["response_trend"] = response_trend
            incident_effectiveness["P11_line"] = response_trend

        incident_distribution = ExcelDataExtractor._read_labeled_pairs(ws, 53, 57, 3, 4)
        if incident_distribution["labels"]:
            incident_dist_obj = {
                "categories": incident_distribution["labels"],
                "values": incident_distribution["values"],
            }
            incident_effectiveness["incident_distribution"] = incident_dist_obj
            incident_effectiveness["P11_pie"] = incident_dist_obj

        ExcelDataExtractor._add_section(output, "incident_effectiveness", incident_effectiveness)

        asset_management: Dict[str, Any] = {}
        asset_map = {
            "intranet_server": "D70",
            "intranet_network": "D72",
            "intranet_iot": "D73",
            "intranet_mss": "D74",
            "rootdomain_internet": "G70",
            "subdomain_internet": "G71",
            "web_asset": "G72",
            "nonweb_asset": "G73",
            "login_entry": "G74",
            "server_total": "D64",
            "pc_total": "D65",
            "internet_entry": "D66",
            "internet_port": "D67",
            "asset_identify": "D68",
        }
        for token, addr in asset_map.items():
            ExcelDataExtractor._put_text(asset_management, token, ws[addr])
        ExcelDataExtractor._add_section(output, "asset_management", asset_management)

        vulnerability_effectiveness: Dict[str, Any] = {}
        vuln_map = {
            "vuln_high": "C77",
            "internet_closed": "D77",
            "admin_weak": "E77",
            "vuln_closed": "F77",
            "scanning": "D40",
        }
        for token, addr in vuln_map.items():
            ExcelDataExtractor._put_text(vulnerability_effectiveness, token, ws[addr])

        vuln_dist = ExcelDataExtractor._read_labeled_pairs(ws, 80, 82, 3, 4)
        if vuln_dist["labels"]:
            vuln_dist_obj = {"categories": vuln_dist["labels"], "values": vuln_dist["values"]}
            vulnerability_effectiveness["vulnerability_distribution"] = vuln_dist_obj
            vulnerability_effectiveness["P14_pie"] = vuln_dist_obj

        # Preserve additional P14 columns regardless of current downstream usage.
        vuln_closed_loop_counts = ExcelDataExtractor._read_labeled_pairs(ws, 80, 82, 3, 5)
        if vuln_closed_loop_counts["labels"]:
            vulnerability_effectiveness["closed_loop_counts"] = {
                "categories": vuln_closed_loop_counts["labels"],
                "values": vuln_closed_loop_counts["values"],
            }

        closed_loop_rate_labels: List[str] = []
        closed_loop_rate_values: List[str] = []
        for r in range(80, 83):
            raw_label = ws.cell(r, 3)
            raw_rate = ws.cell(r, 6)
            if not ExcelDataExtractor._has_value(raw_label) and not ExcelDataExtractor._has_value(raw_rate):
                continue
            closed_loop_rate_labels.append(ExcelDataExtractor._to_text(raw_label))
            closed_loop_rate_values.append(ExcelDataExtractor._to_pct(raw_rate))
        if closed_loop_rate_labels:
            vulnerability_effectiveness["closed_loop_rates"] = {
                "categories": closed_loop_rate_labels,
                "values": closed_loop_rate_values,
            }
        ExcelDataExtractor._add_section(output, "vulnerability_effectiveness", vulnerability_effectiveness)

        threat_effectiveness: Dict[str, Any] = {}
        # Always preserve raw monthly rows C97:N100 (12 values each) in input JSON.
        # Keep these under threat_trend for a single coherent trend payload.
        raw_cols = range(3, 15)  # C..N
        threat_trend: Dict[str, Any] = {
            "internal_lateral_attack_counts": [
                ExcelDataExtractor._to_number(ws.cell(97, c).value, 0) for c in raw_cols
            ],
            "alert_counts": [
                ExcelDataExtractor._to_number(ws.cell(98, c).value, 0) for c in raw_cols
            ],
            "valid_incident_counts": [
                ExcelDataExtractor._to_number(ws.cell(99, c).value, 0) for c in raw_cols
            ],
            "risk_host_counts": [
                ExcelDataExtractor._to_number(ws.cell(100, c).value, 0) for c in raw_cols
            ],
        }

        month_cols = ExcelDataExtractor._read_month_columns(ws, 94, 3, 13)
        if month_cols:
            months = [ExcelDataExtractor._to_text(ws.cell(94, c)) for c in month_cols]
            external_attacks = [ExcelDataExtractor._to_number(ws.cell(95, c).value, 0) for c in month_cols]
            malicious_outbound = [ExcelDataExtractor._to_number(ws.cell(96, c).value, 0) for c in month_cols]
            alerts_monthly = [ExcelDataExtractor._to_number(ws.cell(98, c).value, 0) for c in month_cols]
            incidents_monthly = [ExcelDataExtractor._to_number(ws.cell(99, c).value, 0) for c in month_cols]

            threat_trend["months"] = months
            threat_trend["external_attacks"] = external_attacks
            threat_trend["malicious_outbound"] = malicious_outbound

        threat_effectiveness["threat_trend"] = threat_trend
        threat_effectiveness["P15_line"] = threat_trend

        threat_map = {
            "xdr_attack": "D84",
            "threat_alert": "D85",
            "mss_ticket": "D86",
            "ticket_response": "D87",
            "policy_check": "D88",
            "policy_optimized": "D89",
            "threat_intel": "D90",
            "threat_asset": "D91",
        }
        for token, addr in threat_map.items():
            ExcelDataExtractor._put_text(threat_effectiveness, token, ws[addr])
        ExcelDataExtractor._add_section(output, "threat_effectiveness", threat_effectiveness)

        critical_assurance: Dict[str, Any] = {}
        critical_map = {
            "duty_critical": "D102",
            "incident_critical": "D103",
            "availability_assure": "D104",
        }
        for token, addr in critical_map.items():
            ExcelDataExtractor._put_text(critical_assurance, token, ws[addr])

        posture_rows = []
        for r in range(103, 110):
            raw_cat = ws.cell(r, 6).value
            raw_attack = ws.cell(r, 7).value
            raw_defense = ws.cell(r, 8).value
            if (
                not ExcelDataExtractor._has_value(raw_cat)
                and not ExcelDataExtractor._has_value(raw_attack)
                and not ExcelDataExtractor._has_value(raw_defense)
            ):
                continue
            posture_rows.append((raw_cat, raw_attack, raw_defense))

        if posture_rows:
            posture_comparison = {
                "categories": [ExcelDataExtractor._to_text(r[0]) for r in posture_rows],
                "attack_counts": [ExcelDataExtractor._to_number(r[1], 0) for r in posture_rows],
                "defense_rates": [ExcelDataExtractor._to_number(r[2], 0) for r in posture_rows],
            }
            critical_assurance["posture_comparison"] = posture_comparison
            critical_assurance["P16_combo"] = posture_comparison

        ExcelDataExtractor._add_section(output, "critical_assurance", critical_assurance)

        platform_effectiveness: Dict[str, Any] = {}
        platform_map = {
            "fw_attack": "D111",
            "fw_auto": "D112",
            "fw_coverage": "D113",
            "aes_risk": "D114",
            "host_handle": "D115",
            "server_coverage": "D116",
            "xdr_log": "D117",
            "alert_aggregate": "D118",
            "incident_intel": "D119",
            "component_offline": "D121",
            "component_log": "D122",
            "component_policy": "D123",
            "component_auto": "D124",
        }
        for token, addr in platform_map.items():
            ExcelDataExtractor._put_text(platform_effectiveness, token, ws[addr])

        platform_extra_map = {
            "aes_trust_risk_count": "F114",
            "agent_install_count": "F115",
            "asset_total_count": "F116",
            "xdr_avg_monthly_log_count": "F117",
            "xdr_avg_monthly_alert_count": "F118",
            "xdr_avg_monthly_incident_count": "F119",
        }
        for token, addr in platform_extra_map.items():
            ExcelDataExtractor._put_text(platform_effectiveness, token, ws[addr])

        ExcelDataExtractor._add_section(output, "platform_effectiveness", platform_effectiveness)

        output = ExcelDataExtractor._format_numbers_for_output(output)
        return ExcelDataExtractor._remove_ai_generated_fields(output)

    @staticmethod
    def extract_data(excel_path: Path) -> Dict[str, Any]:
        """Extract data from Excel file.

        Supported layout:
        - Classic ops workbook sheet `数据统计`
        """
        try:
            wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
            sheet_names = set(wb.sheetnames)

            if ExcelDataExtractor.CLASSIC_REQUIRED_SHEET in sheet_names:
                ws = wb[ExcelDataExtractor.CLASSIC_REQUIRED_SHEET]
                data = ExcelDataExtractor._extract_classic_ops(ws)
                wb.close()
                logger.info("Excel parsed as classic-ops format")
                return data

            wb.close()
            raise DataValidationError(
                field="worksheet",
                message=(
                    "Unsupported workbook layout. "
                    "Expected sheet: 数据统计."
                ),
            )
        except InvalidFileException as e:
            raise FileValidationError(
                filename=str(excel_path.name),
                reason=f"Invalid Excel file format: {e}",
            ) from e
        except (DataValidationError, FileValidationError):
            raise
        except Exception as e:
            raise DataValidationError(
                field="excel_data",
                message=f"Excel parsing failed: {e}",
            ) from e


class ExcelHandler:
    """High-level handler for Excel file uploads."""

    def __init__(self, max_size_mb: int = 10, chunk_size: int = 8192):
        self.validator = ExcelValidator(max_size_mb=max_size_mb)
        self.extractor = ExcelDataExtractor()
        self.chunk_size = chunk_size

    async def process_upload(
        self,
        file_content: bytes,
        filename: str,
        content_type: str,
        session_dir: Path,
    ) -> Dict[str, Any]:
        """Process uploaded Excel file and persist parsed JSON."""
        self.validator.validate_extension(filename)
        self.validator.validate_mime_type(filename, content_type)
        self.validator.validate_size(filename, len(file_content))

        session_dir.mkdir(parents=True, exist_ok=True)
        excel_path = session_dir / "uploaded.xlsx"

        try:
            excel_path.write_bytes(file_content)
            logger.info("File saved: %s (%.2f KB)", excel_path, len(file_content) / 1024)

            input_data = self.extractor.extract_data(excel_path)

            json_path = session_dir / "input.json"
            with json_path.open("w", encoding="utf-8") as f:
                json.dump(input_data, f, ensure_ascii=False, indent=2)

            logger.info("JSON saved: %s", json_path)
            return input_data
        except Exception:
            if excel_path.exists():
                try:
                    excel_path.unlink()
                except Exception:
                    pass
            raise

    def get_file_info(self, data: Dict[str, Any], file_size: int) -> Dict[str, Any]:
        """Generate file information summary for API response."""
        preview = {
            "period": data.get("period"),
            "template_id": data.get("template_id", "unknown"),
        }
        if "cover" in data:
            preview["cover"] = data.get("cover")

        return {
            "file_size_mb": round(file_size / 1024 / 1024, 2),
            "preview": preview,
        }
