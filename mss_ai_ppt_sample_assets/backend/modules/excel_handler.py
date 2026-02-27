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
    def _has_value(value: Any) -> bool:
        if value is None:
            return False
        if isinstance(value, str):
            return bool(value.strip())
        return True

    @staticmethod
    def _to_text(value: Any, default: str = "") -> str:
        if value is None:
            return default
        if isinstance(value, (datetime, date)):
            return value.strftime("%Y-%m-%d")
        text = str(value).strip()
        return text if text else default

    @staticmethod
    def _to_number(value: Any, default: float = 0) -> float:
        if value is None:
            return default
        if isinstance(value, (int, float)):
            return value

        text = str(value).strip()
        if text in {"", "None", "#DIV/0!", "#N/A"}:
            return default

        if text.endswith("%"):
            try:
                return float(text[:-1]) / 100
            except Exception:
                return default

        try:
            if "." in text:
                return float(text)
            return float(int(text))
        except Exception:
            return default

    @staticmethod
    def _to_pct(value: Any) -> str:
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

        raw_period_start = ws["L1"].value
        raw_period_end = ws["M1"].value
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
        ExcelDataExtractor._put_text(architecture, "AF_count", ws["D3"].value)
        ExcelDataExtractor._put_text(architecture, "STA_count", ws["D4"].value)
        ExcelDataExtractor._put_text(architecture, "EDR_count", ws["D5"].value)
        ExcelDataExtractor._put_text(architecture, "TSS_count", ws["D6"].value)
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
            ExcelDataExtractor._put_text(deliverables, token, ws[addr].value)
        # P7: primary mapping is G16/G17. Fallback to G17/G18 for legacy-filled sheets.
        if ExcelDataExtractor._has_value(ws["G16"].value):
            ExcelDataExtractor._put_text(deliverables, "biweekly_threat", ws["G16"].value)
        elif ExcelDataExtractor._has_value(ws["G17"].value):
            ExcelDataExtractor._put_text(deliverables, "biweekly_threat", ws["G17"].value)

        if ExcelDataExtractor._has_value(ws["G18"].value):
            ExcelDataExtractor._put_text(deliverables, "phishing_poster", ws["G18"].value)
        elif ExcelDataExtractor._has_value(ws["G17"].value):
            ExcelDataExtractor._put_text(deliverables, "phishing_poster", ws["G17"].value)
        ExcelDataExtractor._add_section(output, "deliverables", deliverables)

        coverage_summary: Dict[str, Any] = {}
        ExcelDataExtractor._put_pct(coverage_summary, "AF_coverage", ws["L22"].value)
        ExcelDataExtractor._put_pct(coverage_summary, "aes_coverage", ws["L24"].value)
        ExcelDataExtractor._put_pct(coverage_summary, "probe_coverage", ws["L23"].value)
        coverage_map = {
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
            ExcelDataExtractor._put_text(coverage_summary, token, ws[addr].value)
        ExcelDataExtractor._add_section(output, "coverage_summary", coverage_summary)

        protection_overview: Dict[str, Any] = {}
        protection_map = {
            "vuln_protect": "C34",
            "alert_judgment": "D34",
            "alert_response": "E34",
            "threat_contain": "F34",
            "af_block": "D95",
            "edr_risk": "D114",
            "policy_check": "D88",
            "surface": "D39",
            "asset": "D40",
            "server": "D41",
            "pc": "D42",
            "vuln_high": "D43",
            "surface_risk": "D44",
            "incident_closed": "D45",
            "vuln_closed": "D46",
            "weekly": "D47",
            "monthly": "D48",
            "xdr_log": "D117",
            "xdr_alert": "D118",
            "xdr_incident": "D119",
            "mss_push": "G38",
            "xdr_auto": "G40",
            "mss_handle": "G41",
            "emergency_handle": "G42",
            "incident_contain": "G43",
        }
        for token, addr in protection_map.items():
            ExcelDataExtractor._put_text(protection_overview, token, ws[addr].value)
        if ExcelDataExtractor._has_value(ws["D87"].value):
            ExcelDataExtractor._put_text(protection_overview, "mss_latency", ws["D87"].value)
        elif ExcelDataExtractor._has_value(ws["G39"].value):
            ExcelDataExtractor._put_text(protection_overview, "mss_latency", ws["G39"].value)
        if "xdr_auto" not in protection_overview and ExcelDataExtractor._has_value(ws["D124"].value):
            ExcelDataExtractor._put_text(protection_overview, "xdr_auto", ws["D124"].value)
        ExcelDataExtractor._add_section(output, "protection_overview", protection_overview)

        incident_effectiveness: Dict[str, Any] = {}
        incident_map = {
            "incident_total": "C51",
            "response_avg": "D51",
            "handle_avg": "E51",
            "incident_closed": "D45",
            "trust_assurance": "F51",
            "security_trust": "D51",
        }
        for token, addr in incident_map.items():
            ExcelDataExtractor._put_text(incident_effectiveness, token, ws[addr].value)

        response_timeliness = ExcelDataExtractor._read_labeled_pairs(ws, 53, 57, 6, 7)
        if response_timeliness["labels"]:
            incident_effectiveness["response_timeliness"] = response_timeliness
            incident_effectiveness["P11_bar"] = response_timeliness

        trend_cols = ExcelDataExtractor._read_month_columns(ws, 60, 3, 12)
        if trend_cols:
            response_trend_months = [ExcelDataExtractor._to_text(ws.cell(60, c).value) for c in trend_cols]
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
            ExcelDataExtractor._put_text(asset_management, token, ws[addr].value)
        ExcelDataExtractor._add_section(output, "asset_management", asset_management)

        vulnerability_effectiveness: Dict[str, Any] = {}
        vuln_map = {
            "vuln_high": "C77",
            "internet_closed": "D77",
            "admin_weak": "E77",
            "vuln_closed": "F77",
        }
        for token, addr in vuln_map.items():
            ExcelDataExtractor._put_text(vulnerability_effectiveness, token, ws[addr].value)
        if ExcelDataExtractor._has_value(ws["D64"].value) or ExcelDataExtractor._has_value(ws["D65"].value):
            asset_total = (
                ExcelDataExtractor._to_number(ws["D64"].value, 0)
                + ExcelDataExtractor._to_number(ws["D65"].value, 0)
            )
            vulnerability_effectiveness["asset"] = ExcelDataExtractor._to_text(asset_total)

        vuln_dist = ExcelDataExtractor._read_labeled_pairs(ws, 80, 82, 3, 4)
        if vuln_dist["labels"]:
            vuln_dist_obj = {"categories": vuln_dist["labels"], "values": vuln_dist["values"]}
            vulnerability_effectiveness["vulnerability_distribution"] = vuln_dist_obj
            vulnerability_effectiveness["P14_pie"] = vuln_dist_obj
        ExcelDataExtractor._add_section(output, "vulnerability_effectiveness", vulnerability_effectiveness)

        threat_effectiveness: Dict[str, Any] = {}
        month_cols = ExcelDataExtractor._read_month_columns(ws, 94, 3, 13)
        if month_cols:
            months = [ExcelDataExtractor._to_text(ws.cell(94, c).value) for c in month_cols]
            external_attacks = [ExcelDataExtractor._to_number(ws.cell(95, c).value, 0) for c in month_cols]
            malicious_outbound = [ExcelDataExtractor._to_number(ws.cell(96, c).value, 0) for c in month_cols]
            alerts_monthly = [ExcelDataExtractor._to_number(ws.cell(98, c).value, 0) for c in month_cols]
            incidents_monthly = [ExcelDataExtractor._to_number(ws.cell(99, c).value, 0) for c in month_cols]

            end_month = period.get("end_month")
            if end_month and end_month in months:
                idx = months.index(end_month)
                threat_effectiveness["xdr_attack"] = ExcelDataExtractor._to_text(external_attacks[idx])
                if idx < len(alerts_monthly):
                    threat_effectiveness["threat_alert"] = ExcelDataExtractor._to_text(alerts_monthly[idx])
                if idx < len(incidents_monthly):
                    threat_effectiveness["mss_ticket"] = ExcelDataExtractor._to_text(incidents_monthly[idx])

            threat_trend = {
                "months": months,
                "external_attacks": external_attacks,
                "malicious_outbound": malicious_outbound,
            }
            threat_effectiveness["threat_trend"] = threat_trend
            threat_effectiveness["P15_line"] = threat_trend

        threat_map = {
            "threat_alert": "D85",
            "mss_ticket": "D86",
            "ticket_response": "D87",
            "policy_check": "D88",
            "policy_optimized": "D89",
            "threat_intel": "D90",
            "threat_asset": "D91",
        }
        for token, addr in threat_map.items():
            ExcelDataExtractor._put_text(threat_effectiveness, token, ws[addr].value)
        ExcelDataExtractor._add_section(output, "threat_effectiveness", threat_effectiveness)

        critical_assurance: Dict[str, Any] = {}
        critical_map = {
            "duty_critical": "D102",
            "incident_critical": "D103",
            "availability_assure": "D104",
        }
        for token, addr in critical_map.items():
            ExcelDataExtractor._put_text(critical_assurance, token, ws[addr].value)

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
