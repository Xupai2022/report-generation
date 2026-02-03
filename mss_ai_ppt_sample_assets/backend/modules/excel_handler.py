"""Excel file upload and processing module.

This module handles:
- File validation (format, size, MIME type)
- Excel data extraction to JSON
- Session-based storage for concurrent requests

Security features:
- Only .xlsx format accepted (no macros)
- Configurable file size limits
- MIME type verification
- Path traversal protection via session validation
"""

from __future__ import annotations

import json
import logging
from pathlib import Path
from typing import Dict, Any, Set, Optional
import openpyxl
from openpyxl.utils.exceptions import InvalidFileException

from mss_ai_ppt_sample_assets.backend.exceptions import (
    FileValidationError,
    DataValidationError
)

logger = logging.getLogger(__name__)


class ExcelValidator:
    """Validates Excel file uploads for security and format compliance."""

    ALLOWED_EXTENSIONS: Set[str] = {'.xlsx'}
    FORBIDDEN_EXTENSIONS: Set[str] = {'.xlsm', '.xls', '.xlsb', '.csv'}
    ALLOWED_MIME_TYPES: Set[str] = {
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'application/zip',
    }

    FORMAT_TIPS = {
        '.xlsm': '此格式可能包含宏代码，请使用Excel另存为 .xlsx 格式',
        '.xls': '旧版Excel格式，请使用Excel 2007+另存为 .xlsx 格式',
        '.xlsb': '二进制Excel格式，请另存为 .xlsx 格式',
        '.csv': 'CSV格式不支持，请使用Excel打开并另存为 .xlsx 格式'
    }

    def __init__(self, max_size_mb: int = 10):
        """Initialize validator.

        Args:
            max_size_mb: Maximum file size in MB (default: 10)
        """
        self.max_size_bytes = max_size_mb * 1024 * 1024

    def validate_extension(self, filename: str) -> None:
        """Validate file extension.

        Args:
            filename: Name of uploaded file

        Raises:
            FileValidationError: If file extension is invalid or forbidden
        """
        file_ext = Path(filename).suffix.lower()

        if file_ext not in self.ALLOWED_EXTENSIONS:
            if file_ext in self.FORBIDDEN_EXTENSIONS:
                detail_msg = self.FORMAT_TIPS.get(file_ext, '请转换为 .xlsx 格式')
                raise FileValidationError(
                    filename=filename,
                    reason=f"禁止上传 {file_ext} 格式文件: {detail_msg}"
                )
            else:
                raise FileValidationError(
                    filename=filename,
                    reason=f"不支持的文件格式 {file_ext}，仅允许 .xlsx"
                )

    def validate_mime_type(self, filename: str, content_type: str) -> None:
        """Validate MIME type.

        Args:
            filename: Name of uploaded file
            content_type: MIME type from upload

        Raises:
            FileValidationError: If MIME type doesn't match expected types
        """
        if content_type not in self.ALLOWED_MIME_TYPES:
            logger.warning(f"MIME mismatch: {filename} - {content_type}")
            raise FileValidationError(
                filename=filename,
                reason=f"文件类型验证失败: 检测到 {content_type}，但期望 .xlsx 格式"
            )

    def validate_size(self, filename: str, size_bytes: int) -> None:
        """Validate file size.

        Args:
            filename: Name of uploaded file
            size_bytes: File size in bytes

        Raises:
            FileValidationError: If file exceeds size limit
        """
        if size_bytes > self.max_size_bytes:
            max_mb = self.max_size_bytes / 1024 / 1024
            actual_mb = round(size_bytes / 1024 / 1024, 2)
            raise FileValidationError(
                filename=filename,
                reason=f"文件大小 {actual_mb}MB 超过限制 {max_mb}MB"
            )


class ExcelDataExtractor:
    """Extracts structured data from Excel files."""

    REQUIRED_SHEET = "数据表"

    @staticmethod
    def extract_data(excel_path: Path) -> Dict[str, Any]:
        """Extract data from Excel file following agreed format.

        Expected Excel format:
        - Sheet name: "数据表"
        - A2: customer_name
        - B2:C2: period (start, end)
        - D2: total alerts
        - E2:H2: alerts by severity
        - A5:B10: alert categories
        - A15:E20: incident details

        Args:
            excel_path: Path to Excel file

        Returns:
            Extracted data as dictionary

        Raises:
            DataValidationError: If required fields are missing or invalid
            FileValidationError: If Excel file is corrupted
        """
        try:
            wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)

            if ExcelDataExtractor.REQUIRED_SHEET not in wb.sheetnames:
                raise DataValidationError(
                    field="工作表",
                    message=f"工作表 '{ExcelDataExtractor.REQUIRED_SHEET}' 不存在。"
                           f"找到的工作表: {', '.join(wb.sheetnames)}"
                )

            ws = wb[ExcelDataExtractor.REQUIRED_SHEET]

            # Extract and validate basic info
            customer_name = ws['A2'].value
            if not customer_name:
                raise DataValidationError(
                    field="customer_name",
                    message="必填字段 'customer_name' (A2) 为空"
                )

            period_start = ws['B2'].value
            if not period_start:
                raise DataValidationError(
                    field="period.start",
                    message="必填字段 'period.start' (B2) 为空"
                )

            period_end = ws['C2'].value
            if not period_end:
                raise DataValidationError(
                    field="period.end",
                    message="必填字段 'period.end' (C2) 为空"
                )

            data = {
                "customer_name": str(customer_name).strip(),
                "period": {
                    "start": str(period_start),
                    "end": str(period_end)
                },
                "alerts": {
                    "total": ws['D2'].value or 0,
                    "by_severity": {
                        "high": ws['E2'].value or 0,
                        "medium": ws['F2'].value or 0,
                        "low": ws['G2'].value or 0,
                        "info": ws['H2'].value or 0
                    }
                },
                "top_alert_categories": [],
                "top_incidents": []
            }

            # Extract alert categories (A5:B10)
            for row in range(5, 11):
                category = ws[f'A{row}'].value
                count = ws[f'B{row}'].value
                if category and count is not None:
                    data["top_alert_categories"].append({
                        "category": str(category).strip(),
                        "count": int(count) if isinstance(count, (int, float)) else 0
                    })

            # Extract incidents (A15:E20)
            for row in range(15, 21):
                incident_id = ws[f'A{row}'].value
                if incident_id:
                    data["top_incidents"].append({
                        "id": str(incident_id).strip(),
                        "type": str(ws[f'B{row}'].value or "").strip(),
                        "severity": str(ws[f'C{row}'].value or "").strip(),
                        "status": str(ws[f'D{row}'].value or "").strip(),
                        "description": str(ws[f'E{row}'].value or "").strip()
                    })

            wb.close()
            logger.info(
                f"✅ Excel parsed: {len(data['top_alert_categories'])} categories, "
                f"{len(data['top_incidents'])} incidents"
            )
            return data

        except InvalidFileException as e:
            raise FileValidationError(
                filename=str(excel_path.name),
                reason=f"无效的Excel文件格式: {str(e)}"
            )
        except (DataValidationError, FileValidationError):
            raise
        except Exception as e:
            raise DataValidationError(
                field="Excel数据",
                message=f"Excel解析失败: {str(e)}"
            )


class ExcelHandler:
    """High-level handler for Excel file uploads."""

    def __init__(
        self,
        max_size_mb: int = 10,
        chunk_size: int = 8192
    ):
        """Initialize handler.

        Args:
            max_size_mb: Maximum file size in MB
            chunk_size: Chunk size for file reading (bytes)
        """
        self.validator = ExcelValidator(max_size_mb=max_size_mb)
        self.extractor = ExcelDataExtractor()
        self.chunk_size = chunk_size

    async def process_upload(
        self,
        file_content: bytes,
        filename: str,
        content_type: str,
        session_dir: Path
    ) -> Dict[str, Any]:
        """Process uploaded Excel file.

        Args:
            file_content: File content as bytes
            filename: Original filename
            content_type: MIME type
            session_dir: Session directory for saving files

        Returns:
            Extracted data dictionary

        Raises:
            FileValidationError: If file validation fails
            DataValidationError: If data extraction fails
        """
        # 1. Validate file
        self.validator.validate_extension(filename)
        self.validator.validate_mime_type(filename, content_type)
        self.validator.validate_size(filename, len(file_content))

        # 2. Save file
        session_dir.mkdir(parents=True, exist_ok=True)
        excel_path = session_dir / "uploaded.xlsx"

        try:
            excel_path.write_bytes(file_content)
            logger.info(f"✅ File saved: {excel_path} ({len(file_content) / 1024:.2f} KB)")

            # 3. Extract data
            input_data = self.extractor.extract_data(excel_path)

            # 4. Save JSON
            json_path = session_dir / "input.json"
            with json_path.open("w", encoding="utf-8") as f:
                json.dump(input_data, f, ensure_ascii=False, indent=2)

            logger.info(f"✅ JSON saved: {json_path}")
            return input_data

        except Exception as e:
            # Clean up on failure
            if excel_path.exists():
                try:
                    excel_path.unlink()
                except Exception:
                    pass
            raise

    def get_file_info(self, data: Dict[str, Any], file_size: int) -> Dict[str, Any]:
        """Generate file information summary.

        Args:
            data: Extracted data
            file_size: File size in bytes

        Returns:
            Summary dictionary for API response
        """
        return {
            "file_size_mb": round(file_size / 1024 / 1024, 2),
            "preview": {
                "customer_name": data.get("customer_name"),
                "period": data.get("period"),
                "alerts": {
                    "total": data.get("alerts", {}).get("total"),
                    "high": data.get("alerts", {}).get("by_severity", {}).get("high")
                },
                "categories_count": len(data.get("top_alert_categories", [])),
                "incidents_count": len(data.get("top_incidents", []))
            }
        }
