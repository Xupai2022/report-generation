from __future__ import annotations

import logging
from pathlib import Path
from typing import Dict, Any, List, Optional

from mss_ai_ppt_sample_assets.backend.models.slidespec import SlideSpecV2
from mss_ai_ppt_sample_assets.backend.modules.template_loader import TemplateRepository

logger = logging.getLogger(__name__)


class PPTGeneratorV2:
    """Fill PPTX template placeholders with V2 slidespec content.

    V2 simplification: slidespec.placeholders directly maps token -> value,
    no need to traverse render_map paths.
    """

    def __init__(self, template_repo: TemplateRepository):
        self.template_repo = template_repo
        try:
            from pptx import Presentation
            from pptx.util import Inches, Pt
            from pptx.enum.chart import XL_CHART_TYPE
            from pptx.chart.data import CategoryChartData
            from pptx.enum.text import PP_ALIGN
            from pptx.dml.color import RGBColor

            self._Presentation = Presentation
            self._Inches = Inches
            self._Pt = Pt
            self._XL_CHART_TYPE = XL_CHART_TYPE
            self._CategoryChartData = CategoryChartData
            self._PP_ALIGN = PP_ALIGN
            self._RGBColor = RGBColor
        except ImportError as exc:
            raise RuntimeError(
                "python-pptx is required for PPT rendering. Please install via requirements.txt."
            ) from exc

    def _replace_tokens_in_shape(self, shape, mapping: Dict[str, str]) -> None:
        """Replace {{TOKEN}} placeholders in shape text."""
        if not shape.has_text_frame:
            return
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                text = run.text
                for token, value in mapping.items():
                    placeholder = f"{{{{{token}}}}}"
                    if placeholder in text:
                        run.text = text.replace(placeholder, str(value) if value else "")
                        text = run.text

    def _render_native_table(
        self,
        slide,
        table_data: Dict[str, Any],
        position: Optional[Dict[str, float]] = None
    ) -> None:
        """Render a native PowerPoint table.

        Args:
            slide: pptx slide object
            table_data: Dict with 'headers' (list) and 'rows' (list of lists)
            position: Dict with 'left', 'top', 'width', 'height' in inches
        """
        if not table_data or 'headers' not in table_data or 'rows' not in table_data:
            logger.warning("Invalid table data format")
            return

        headers = table_data['headers']
        rows = table_data['rows']

        if not headers or not rows:
            logger.warning("Empty table data")
            return

        # Default position if not specified
        if position is None:
            position = {'left': 1.0, 'top': 2.5, 'width': 8.0, 'height': 3.0}

        # Calculate dimensions
        num_rows = len(rows) + 1  # +1 for header
        num_cols = len(headers)

        # Create table
        left = self._Inches(position.get('left', 1.0))
        top = self._Inches(position.get('top', 2.5))
        width = self._Inches(position.get('width', 8.0))
        height = self._Inches(position.get('height', 3.0))

        table = slide.shapes.add_table(num_rows, num_cols, left, top, width, height).table

        # Set headers
        for col_idx, header in enumerate(headers):
            cell = table.rows[0].cells[col_idx]
            cell.text = str(header)
            # Style header
            cell.fill.solid()
            cell.fill.fore_color.rgb = self._RGBColor(79, 129, 189)  # Blue header
            paragraph = cell.text_frame.paragraphs[0]
            paragraph.font.bold = True
            paragraph.font.size = self._Pt(11)
            paragraph.font.color.rgb = self._RGBColor(255, 255, 255)  # White text

        # Set data rows
        for row_idx, row_data in enumerate(rows):
            for col_idx, cell_value in enumerate(row_data):
                if col_idx < num_cols:  # Prevent index out of range
                    cell = table.rows[row_idx + 1].cells[col_idx]
                    cell.text = str(cell_value) if cell_value is not None else ""
                    # Style data cells
                    paragraph = cell.text_frame.paragraphs[0]
                    paragraph.font.size = self._Pt(10)

        logger.info(f"Rendered native table: {num_rows} rows x {num_cols} cols")

    def _render_bar_chart(
        self,
        slide,
        chart_data: Dict[str, Any],
        position: Optional[Dict[str, float]] = None
    ) -> None:
        """Render a bar/column chart.

        Args:
            slide: pptx slide object
            chart_data: Dict with 'categories' (list), 'series' (list of dicts with 'name' and 'values')
            position: Dict with 'left', 'top', 'width', 'height' in inches
        """
        if not chart_data or 'categories' not in chart_data or 'series' not in chart_data:
            logger.warning("Invalid bar chart data format")
            return

        categories = chart_data['categories']
        series_list = chart_data['series']

        if not categories or not series_list:
            logger.warning("Empty bar chart data")
            return

        # Default position if not specified
        if position is None:
            position = {'left': 1.0, 'top': 2.0, 'width': 8.0, 'height': 4.5}

        # Create chart data
        chart_data_obj = self._CategoryChartData()
        chart_data_obj.categories = categories

        for series in series_list:
            series_name = series.get('name', 'Series')
            series_values = series.get('values', [])
            chart_data_obj.add_series(series_name, series_values)

        # Add chart to slide
        left = self._Inches(position.get('left', 1.0))
        top = self._Inches(position.get('top', 2.0))
        width = self._Inches(position.get('width', 8.0))
        height = self._Inches(position.get('height', 4.5))

        chart = slide.shapes.add_chart(
            self._XL_CHART_TYPE.COLUMN_CLUSTERED,
            left, top, width, height,
            chart_data_obj
        ).chart

        # Set chart title if provided
        if 'title' in chart_data and chart_data['title']:
            chart.has_title = True
            chart.chart_title.text_frame.text = chart_data['title']

        logger.info(f"Rendered bar chart with {len(categories)} categories and {len(series_list)} series")

    def _render_pie_chart(
        self,
        slide,
        chart_data: Dict[str, Any],
        position: Optional[Dict[str, float]] = None
    ) -> None:
        """Render a pie chart.

        Args:
            slide: pptx slide object
            chart_data: Dict with 'categories' (list) and 'values' (list)
            position: Dict with 'left', 'top', 'width', 'height' in inches
        """
        if not chart_data or 'categories' not in chart_data or 'values' not in chart_data:
            logger.warning("Invalid pie chart data format")
            return

        categories = chart_data['categories']
        values = chart_data['values']

        if not categories or not values or len(categories) != len(values):
            logger.warning("Invalid or mismatched pie chart data")
            return

        # Default position if not specified
        if position is None:
            position = {'left': 2.0, 'top': 2.0, 'width': 6.0, 'height': 4.5}

        # Create chart data
        chart_data_obj = self._CategoryChartData()
        chart_data_obj.categories = categories
        chart_data_obj.add_series('', values)  # Pie chart typically has one series

        # Add chart to slide
        left = self._Inches(position.get('left', 2.0))
        top = self._Inches(position.get('top', 2.0))
        width = self._Inches(position.get('width', 6.0))
        height = self._Inches(position.get('height', 4.5))

        chart = slide.shapes.add_chart(
            self._XL_CHART_TYPE.PIE,
            left, top, width, height,
            chart_data_obj
        ).chart

        # Set chart title if provided
        if 'title' in chart_data and chart_data['title']:
            chart.has_title = True
            chart.chart_title.text_frame.text = chart_data['title']

        logger.info(f"Rendered pie chart with {len(categories)} categories")

    def _process_chart_placeholder(
        self,
        slide,
        token: str,
        value: Any,
        chart_type: str
    ) -> bool:
        """Process a chart placeholder value.

        Args:
            slide: pptx slide object
            token: placeholder token name
            value: chart data (dict or special format)
            chart_type: 'bar_chart' or 'pie_chart'

        Returns:
            True if chart was rendered, False otherwise
        """
        if not isinstance(value, dict):
            logger.warning(f"Chart placeholder {token} has invalid data type: {type(value)}")
            return False

        try:
            if chart_type == 'bar_chart':
                self._render_bar_chart(slide, value, value.get('position'))
            elif chart_type == 'pie_chart':
                self._render_pie_chart(slide, value, value.get('position'))
            else:
                logger.warning(f"Unknown chart type: {chart_type}")
                return False
            return True
        except Exception as e:
            logger.error(f"Failed to render {chart_type} for {token}: {e}")
            return False

    def _process_table_placeholder(
        self,
        slide,
        token: str,
        value: Any
    ) -> bool:
        """Process a native table placeholder value.

        Args:
            slide: pptx slide object
            token: placeholder token name
            value: table data (dict with headers and rows)

        Returns:
            True if table was rendered, False otherwise
        """
        if not isinstance(value, dict):
            logger.warning(f"Table placeholder {token} has invalid data type: {type(value)}")
            return False

        try:
            self._render_native_table(slide, value, value.get('position'))
            return True
        except Exception as e:
            logger.error(f"Failed to render table for {token}: {e}")
            return False

    def render(self, slidespec: SlideSpecV2, output_path: Path) -> Path:
        """Render V2 slidespec to PPTX.

        Args:
            slidespec: V2 slidespec with placeholders filled
            output_path: Where to save the output PPTX

        Returns:
            Path to the saved PPTX file
        """
        template_path = self.template_repo.get_pptx_path(slidespec.template_id)
        prs = self._Presentation(template_path)

        # Load template descriptor to get placeholder types
        template_desc = self.template_repo.get_descriptor_v2(slidespec.template_id)

        # Build mapping: slide_key -> placeholder_definitions
        placeholder_types = {}
        for slide_def in template_desc.slides:
            placeholder_types[slide_def.slide_key] = {
                ph.token: ph.type for ph in slide_def.placeholders
            }

        # Build mapping: slide_no -> placeholders dict
        slides_by_no = {s.slide_no: s for s in slidespec.slides}

        for slide_no, slide_content in slides_by_no.items():
            if slide_no > len(prs.slides):
                continue

            # pptx slides are 0-indexed
            pptx_slide = prs.slides[slide_no - 1]

            # Get placeholder types for this slide
            slide_types = placeholder_types.get(slide_content.slide_key, {})

            # Separate placeholders by type
            text_placeholders = {}
            chart_bar_placeholders = []
            chart_pie_placeholders = []
            table_placeholders = []

            for token, value in slide_content.placeholders.items():
                ph_type = slide_types.get(token, 'text')

                if ph_type in ('bar_chart',):
                    chart_bar_placeholders.append((token, value))
                elif ph_type in ('pie_chart',):
                    chart_pie_placeholders.append((token, value))
                elif ph_type in ('native_table',):
                    table_placeholders.append((token, value))
                else:
                    # Regular text placeholder
                    if value is None:
                        value = ""
                    elif isinstance(value, list):
                        # Join list items with newlines
                        value = "\n".join(str(v) for v in value)
                    text_placeholders[token] = str(value)

            # Replace text tokens in all shapes
            for shape in pptx_slide.shapes:
                self._replace_tokens_in_shape(shape, text_placeholders)

            # Render charts
            for token, value in chart_bar_placeholders:
                self._process_chart_placeholder(pptx_slide, token, value, 'bar_chart')

            for token, value in chart_pie_placeholders:
                self._process_chart_placeholder(pptx_slide, token, value, 'pie_chart')

            # Render tables
            for token, value in table_placeholders:
                self._process_table_placeholder(pptx_slide, token, value)

        # Save output
        output_path.parent.mkdir(parents=True, exist_ok=True)
        if output_path.exists():
            try:
                output_path.unlink()
            except OSError:
                output_path = output_path.with_name(output_path.stem + "_new" + output_path.suffix)

        prs.save(output_path)
        return output_path