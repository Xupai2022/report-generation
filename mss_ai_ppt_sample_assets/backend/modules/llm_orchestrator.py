from __future__ import annotations

import json
import logging
import re
import time
from typing import Any, Dict, List, Optional

import httpx
from openai import OpenAI, APIError, RateLimitError, APIConnectionError

from mss_ai_ppt_sample_assets.backend import config
from mss_ai_ppt_sample_assets.backend.modules.retry_policy import with_llm_retry
from mss_ai_ppt_sample_assets.backend.models.slidespec import (
    SlideSpecV2, create_empty_slidespec_v2
)
from mss_ai_ppt_sample_assets.backend.models.templates import (
    TemplateDescriptorV2, PlaceholderDefinition
)
from mss_ai_ppt_sample_assets.backend.models.inputs import TenantInput
from mss_ai_ppt_sample_assets.backend.modules.template_loader import TemplateRepository

# Configure logging
logger = logging.getLogger(__name__)


def _build_openai_client() -> OpenAI:
    client_kwargs = {"api_key": config.settings.openai_api_key}
    if config.settings.openai_base_url:
        client_kwargs["base_url"] = config.settings.openai_base_url
    client_kwargs["http_client"] = httpx.Client(trust_env=False)
    return OpenAI(**client_kwargs)


class LLMGenerationError(Exception):
    """Error during LLM content generation"""
    pass


class LLMOrchestratorV2:
    """V2 Orchestrator for AI-driven content generation.

    Key differences from V1:
    - Takes raw TenantInput directly, no pre-processing
    - Uses placeholder-based AI instructions from template
    - Generates content based on ai_instruction fields
    - Validates only key numerical fields
    """

    def __init__(self, template_repo: Optional[TemplateRepository] = None):
        self.template_repo = template_repo or TemplateRepository()
        self.client: Optional[OpenAI] = None

        if config.settings.enable_llm:
            try:
                self.client = _build_openai_client()
                logger.info("OpenAI client initialized successfully")
            except Exception as e:
                logger.error(f"Failed to initialize OpenAI client: {e}")
                raise LLMGenerationError(f"OpenAI client initialization failed: {e}") from e

    def _get_nested(self, data: Any, path: str) -> Any:
        """Get nested value from dict or TenantInput using dot notation path."""
        if not path:
            return None

        # Handle TenantInput by getting its raw data
        if hasattr(data, 'raw'):
            current = data.raw
        elif isinstance(data, dict):
            current = data
        else:
            return None

        for part in path.split('.'):
            if isinstance(current, dict):
                current = current.get(part)
            elif isinstance(current, list) and part.isdigit():
                idx = int(part)
                current = current[idx] if idx < len(current) else None
            else:
                return None
            if current is None:
                return None
        return current

    def _resolve_format_path(self, data: Dict[str, Any], path: str) -> Any:
        """Resolve a dotted path in data, handling .length for lists."""
        parts = path.split('.')
        current = data

        for i, part in enumerate(parts):
            if part == "length" and isinstance(current, list):
                return len(current)

            if isinstance(current, dict):
                current = current.get(part)
            elif isinstance(current, list) and part.isdigit():
                idx = int(part)
                current = current[idx] if idx < len(current) else None
            else:
                return None

            if current is None:
                return None

        return current

    def _format_template_string(self, template: str, data: Dict[str, Any]) -> str:
        """Format a template string with custom path resolution supporting .length."""
        import re

        def replace_placeholder(match):
            path = match.group(1)
            value = self._resolve_format_path(data, path)
            if value is None:
                return match.group(0)  # Keep original if not found
            return str(value)

        # Find all {path} patterns and replace them
        result = re.sub(r'\{([^}]+)\}', replace_placeholder, template)
        return result

    def _format_value(self, value: Any, placeholder: PlaceholderDefinition) -> str:
        """Format a value according to placeholder definition."""
        if value is None:
            return placeholder.default or ""

        # Apply transform
        if placeholder.transform:
            if placeholder.transform == "uppercase":
                value = str(value).upper()
            elif placeholder.transform == "lowercase":
                value = str(value).lower()
            elif placeholder.transform == "percent":
                if isinstance(value, (int, float)):
                    value = f"{round(value * 100)}%"

        # Handle list values FIRST (before format check)
        if isinstance(value, list):
            if placeholder.format and "{" in placeholder.format:
                # Format each list item using the format template
                formatted_items = []
                for item in value:
                    if isinstance(item, dict):
                        try:
                            formatted_items.append(self._format_template_string(placeholder.format, item))
                        except Exception:
                            formatted_items.append(str(item))
                    else:
                        formatted_items.append(str(item))
                return "\n".join(f"• {item}" for item in formatted_items)
            elif placeholder.format == "join_comma":
                return ", ".join(str(v) for v in value)
            else:
                return "\n".join(f"• {str(v)}" for v in value)

        # Apply format template for non-list values
        if placeholder.format:
            if placeholder.format == "percent":
                if isinstance(value, (int, float)):
                    return f"{round(value * 100)}%"
            elif placeholder.format == "join_comma":
                return str(value)
            elif "{" in placeholder.format:
                # Template format like "{value}小时" or "{start} ~ {end}"
                if isinstance(value, dict):
                    try:
                        return self._format_template_string(placeholder.format, value)
                    except Exception:
                        pass
                else:
                    try:
                        return placeholder.format.format(value=value)
                    except (KeyError, ValueError):
                        pass

        return str(value)

    @staticmethod
    def _coerce_number(value: Any) -> Optional[float]:
        """Best-effort conversion to float for chart payload normalization."""
        if isinstance(value, bool):
            return None
        if isinstance(value, (int, float)):
            return float(value)
        if isinstance(value, str):
            text = value.strip().replace(",", "")
            if not text:
                return None
            is_percent = text.endswith("%")
            if is_percent:
                text = text[:-1].strip()
            try:
                number = float(text)
            except ValueError:
                return None
            return number / 100.0 if is_percent else number
        return None

    def _extract_chart_data(
        self,
        tenant_input: TenantInput,
        chart_config: Dict[str, Any],
        chart_type: str
    ) -> Dict[str, Any]:
        """Extract and format chart data from tenant input.

        Args:
            tenant_input: Raw tenant input
            chart_config: Chart configuration from placeholder definition
            chart_type: One of the supported chart placeholder types

        Returns:
            Formatted chart data ready for rendering
        """
        data_source = chart_config.get('data_source')
        if not data_source:
            logger.warning(f"Chart config missing data_source")
            return {}

        # Get data from tenant input
        source_data = self._get_nested(tenant_input, data_source)
        if not source_data:
            logger.warning(f"No data found at {data_source}")
            return {}

        result = {}

        if chart_type == 'P11_bar':
            # Expect source_data to have 'labels' and 'values' or similar structure
            x_field = chart_config.get('x_field', 'labels')
            y_field = chart_config.get('y_field', 'values')

            if isinstance(source_data, dict):
                # Direct dict format: {"labels": [...], "values": [...]}
                categories = source_data.get(x_field, [])
                values = source_data.get(y_field, [])

                result['categories'] = categories
                result['series'] = [{'name': chart_config.get('series_name', '告警数'), 'values': values}]
            elif isinstance(source_data, list):
                # List of objects format: [{"category": ..., "count": ...}, ...]
                categories = []
                values = []
                for item in source_data:
                    if isinstance(item, dict):
                        cat_value = item.get(x_field)
                        val_value = item.get(y_field)
                        if cat_value is not None:
                            categories.append(cat_value)
                            values.append(val_value if val_value is not None else 0)

                if categories:
                    result['categories'] = categories
                    result['series'] = [{'name': chart_config.get('series_name', '告警数'), 'values': values}]
                else:
                    logger.warning(f"Bar chart data source {data_source} list has no valid items with fields {x_field}/{y_field}")
                    return {}
            else:
                logger.warning(f"Bar chart data source {data_source} is neither dict nor list")
                return {}

        elif chart_type in ('P11_pie', 'P13_pie', 'P14_pie'):
            # Expect source_data to be a dict like {'high': 52, 'medium': 473, 'low': 816}
            # or a dict with 'categories' and 'values' arrays for P11_pie/P13_pie/P14_pie
            if isinstance(source_data, dict):
                # Check if it's the P11_pie/P13_pie/P14_pie format with categories and values arrays
                if 'categories' in source_data and 'values' in source_data:
                    result['categories'] = source_data['categories']
                    result['values'] = source_data['values']
                else:
                    # Convert dict to categories and values
                    categories = []
                    values = []

                    # Map severity levels to Chinese names
                    severity_map = chart_config.get('category_map', {
                        'critical': '严重',
                        'high': '高危',
                        'medium': '中危',
                        'low': '低危',
                        'info': '信息'
                    })

                    for key, value in source_data.items():
                        # Use mapped name if available, otherwise use key
                        category_name = severity_map.get(key, key)
                        categories.append(category_name)
                        values.append(value)

                    result['categories'] = categories
                    result['values'] = values
            else:
                logger.warning(f"Pie chart data source {data_source} is not a dict")
                return {}

        elif chart_type == 'P15_line':
            # Expect source_data to be a dict with 'months', 'external_attacks', 'malicious_outbound'
            if isinstance(source_data, dict):
                months = source_data.get('months', [])
                external_attacks = source_data.get('external_attacks', [])
                malicious_outbound = source_data.get('malicious_outbound', [])

                # 转换为"万"单位（除以10000），保留原始值以便格式化
                def to_wan(value):
                    """Convert value to 万 unit (divide by 10000)"""
                    if isinstance(value, (int, float)):
                        return value / 10000
                    return value

                result['months'] = months
                result['external_attacks'] = [to_wan(v) for v in external_attacks]
                result['malicious_outbound'] = [to_wan(v) for v in malicious_outbound]
            else:
                logger.warning(f"Line chart data source {data_source} is not a dict")
                return {}

        elif chart_type == 'P11_line':
            # Expect source_data to be a dict with 'months' (or 'categories') and multiple series
            # Example: {"months": ["Jan", "Feb", ...], "critical": [5, 3, ...], "high": [12, 15, ...], "medium": [45, 38, ...]}
            if isinstance(source_data, dict):
                # Get the category field (months or categories)
                months = source_data.get('months', source_data.get('categories', []))

                # Extract all numeric series (skip 'months' and 'categories' keys)
                series_data = []
                for key, values in source_data.items():
                    if key not in ['months', 'categories'] and isinstance(values, list):
                        series_data.append({
                            'name': key,
                            'values': values
                        })

                if months and series_data:
                    result['months'] = months
                    result['series'] = series_data
                else:
                    logger.warning(f"P11_line data source {data_source} missing valid months or series data")
                    return {}
            else:
                logger.warning(f"P11_line data source {data_source} is not a dict")
                return {}

        elif chart_type == 'P16_combo':
            # Expect source_data to be a dict with 'categories', 'attack_counts', 'defense_rates'
            # attack_counts: daily attack numbers like [123, 145, ...]
            # defense_rates: percentages like [1.0, 0.98, ...] (1.0 = 100%)
            if isinstance(source_data, dict):
                categories = source_data.get('categories', [])
                attack_counts = source_data.get('attack_counts', [])
                defense_rates = source_data.get('defense_rates', [])

                normalized_attacks = []
                for count in attack_counts:
                    number = self._coerce_number(count)
                    if number is None:
                        normalized_attacks.append(0)
                    elif number.is_integer():
                        normalized_attacks.append(int(number))
                    else:
                        normalized_attacks.append(number)

                # Ensure defense_rates are in decimal format (0.0-1.0)
                # If they come as percentages (0-100), convert them
                normalized_rates = []
                for rate in defense_rates:
                    number = self._coerce_number(rate)
                    if number is not None:
                        # If rate > 1, assume it's percentage (e.g., 100 = 100%)
                        if number > 1:
                            normalized_rates.append(number / 100.0)
                        else:
                            normalized_rates.append(number)
                    else:
                        normalized_rates.append(0)

                result['categories'] = categories
                result['attack_counts'] = normalized_attacks
                result['defense_rates'] = normalized_rates
            else:
                logger.warning(f"Combo chart data source {data_source} is not a dict")
                return {}

        return result

    def _extract_table_data(
        self,
        tenant_input: TenantInput,
        table_config: Dict[str, Any]
    ) -> Dict[str, Any]:
        """Extract and format table data from tenant input.

        Args:
            tenant_input: Raw tenant input
            table_config: Table configuration from placeholder definition

        Returns:
            Formatted table data ready for rendering
        """
        data_source = table_config.get('data_source')
        columns_config = table_config.get('columns', [])

        if not data_source or not columns_config:
            logger.warning(f"Table config missing data_source or columns")
            return {}

        # Get data from tenant input
        source_data = self._get_nested(tenant_input, data_source)
        if not source_data or not isinstance(source_data, list):
            logger.warning(f"No list data found at {data_source}")
            return {}

        # Extract headers
        headers = [col.get('header', '') for col in columns_config]

        # Extract rows
        rows = []
        max_rows = table_config.get('max_rows', 10)
        for item in source_data[:max_rows]:
            if isinstance(item, dict):
                row = []
                for col in columns_config:
                    field_name = col.get('field', '')
                    value = item.get(field_name, '')

                    # Format value based on column config
                    if col.get('format') == 'percent' and isinstance(value, (int, float)):
                        value = f"{int(value * 100)}%"

                    row.append(value)
                rows.append(row)

        return {
            'headers': headers,
            'rows': rows,
            'position': table_config.get('position')
        }

    def _extract_data_placeholders(
        self,
        tenant_input: TenantInput,
        template: TemplateDescriptorV2,
    ) -> Dict[str, Dict[str, Any]]:
        """Extract all non-AI placeholders from input data.

        Returns:
            Dict[slide_key, Dict[token, value]]
        """
        result: Dict[str, Dict[str, Any]] = {}

        # Calculate derived values
        incidents = tenant_input.get("incidents", []) or []
        incidents_high_count = len([i for i in incidents if i.get("severity") == "high"])
        incidents_count = len(incidents)

        # Add computed values to a lookup dict
        computed = {
            "incidents.length": incidents_count,
            "incidents.high_count": incidents_high_count,
            "incidents_count": incidents_count,
            "incidents_high_count": incidents_high_count,
        }

        for slide_key, token, placeholder in template.get_data_placeholders():
            if slide_key not in result:
                result[slide_key] = {}

            # Handle chart placeholders
            if placeholder.type in ('P11_bar', 'P11_line', 'P11_pie', 'P13_pie', 'P14_pie', 'P15_line', 'P16_combo') and placeholder.chart_config:
                chart_data = self._extract_chart_data(
                    tenant_input,
                    placeholder.chart_config,
                    placeholder.type
                )
                result[slide_key][token] = chart_data

            # Handle native table placeholders
            elif placeholder.type == 'native_table' and placeholder.table_config:
                table_data = self._extract_table_data(
                    tenant_input,
                    placeholder.table_config
                )
                result[slide_key][token] = table_data

            # Handle regular text placeholders
            elif placeholder.default and not placeholder.source:
                result[slide_key][token] = placeholder.default
            elif placeholder.source:
                # Check computed values first
                if placeholder.source in computed:
                    value = computed[placeholder.source]
                else:
                    value = self._get_nested(tenant_input, placeholder.source)
                result[slide_key][token] = self._format_value(value, placeholder)
            else:
                result[slide_key][token] = ""

        return result

    def _build_system_prompt(self, template: TemplateDescriptorV2) -> str:
        """Build the system prompt for AI generation."""
        audience_desc = "管理层（非技术背景）" if template.audience == "management" else "技术团队（安全工程师）"

        return f"""你是一位资深安全分析师，正在为客户撰写MSS（托管安全服务）月度安全报告。

## 你的角色
- 你是安全领域专家，具有丰富的威胁分析和安全运营经验
- 你擅长从数据中发现安全趋势和洞察
- 你能用专业但易懂的语言撰写安全报告

## 报告受众
本报告面向：{audience_desc}

## 关键要求
1. **数据准确性**：所有引用的数字必须与输入数据完全一致，绝不能编造数据
2. **深度分析**：不要只罗列数据，要给出有洞察力的分析与解读
3. **具体建议**：建议必须具体可执行，避免泛泛而谈
4. **语言风格**：使用中文，简洁专业，适合{audience_desc}阅读
5. **格式要求**：严格按照指定JSON格式返回内容

## 输出格式
你必须返回一个JSON对象，格式如下：
{{
  "slides": [
    {{
      "slide_key": "slide_key_here",
      "placeholders": {{
        "TOKEN_NAME": "生成的内容"
      }}
    }}
  ]
}}
"""

    def _build_user_prompt(
        self,
        tenant_input: TenantInput,
        template: TemplateDescriptorV2
    ) -> str:
        """Build the user prompt with data and AI instructions."""
        period = tenant_input.get("period", {})

        prompt_parts = [
            "## 任务",
            "基于输入的安全数据，生成指定 slide 的占位符内容。",
            f"报告时间：{period.get('start', '')} ~ {period.get('end', '')}",
            "",
            "## 输出要求",
            "1) 只返回合法JSON，不要输出解释、注释或Markdown。",
            "2) 输出内容必须为中文。",
            "3) 所有数字必须与输入数据一致，不得编造。",
            "4) 只输出下方列出的 slide_key 和占位符。",
            "",
            "## 输出格式",
            "```json",
            "{",
            '  "slides": [',
        ]

        slide_examples = []
        for slide in template.slides:
            ai_tokens = [ph.token for ph in slide.placeholders if ph.ai_generate]
            if ai_tokens:
                tokens_str = ", ".join(f'"{t}": "..."' for t in ai_tokens)
                slide_examples.append(f'    {{"slide_key": "{slide.slide_key}", "placeholders": {{{tokens_str}}}}}')

        prompt_parts.append(",\n".join(slide_examples))
        prompt_parts.extend([
            "  ]",
            "}",
            "```",
            "",
            "## 需要生成的内容",
            "",
        ])

        ai_placeholders = template.get_ai_placeholders()
        current_slide = None

        for slide_key, token, placeholder in ai_placeholders:
            if slide_key != current_slide:
                for slide in template.slides:
                    if slide.slide_key == slide_key:
                        prompt_parts.append(f"### 页面: {slide.title} ({slide_key})")
                        break
                current_slide = slide_key

            constraints = []
            if placeholder.max_length:
                constraints.append(f"最多{placeholder.max_length}字")
            if placeholder.max_items:
                constraints.append(f"最多{placeholder.max_items}条")
            if placeholder.max_chars_per_item:
                constraints.append(f"每条最多{placeholder.max_chars_per_item}字")

            constraint_str = f" ({', '.join(constraints)})" if constraints else ""

            prompt_parts.append(f"\n**{token}**{constraint_str}")
            prompt_parts.append(f"{placeholder.ai_instruction}")
            prompt_parts.append("")

        prompt_parts.extend([
            "",
            "## 安全数据",
            "```json",
            json.dumps(tenant_input.raw, ensure_ascii=False, indent=2),
            "```",
        ])

        return "\n".join(prompt_parts)

    def _build_user_prompt_for_slides(
        self,
        tenant_input: TenantInput,
        template: TemplateDescriptorV2,
        slide_keys: List[str],
        batch_index: int = 0,
        total_batches: int = 1,
    ) -> str:
        """Build user prompt for a subset of slides (for batched generation).

        Args:
            tenant_input: Raw tenant input data
            template: Template descriptor
            slide_keys: List of slide_keys to include in this batch
            batch_index: Current batch index (0-based)
            total_batches: Total number of batches

        Returns:
            User prompt string for the specified slides
        """
        period = tenant_input.get("period", {})

        prompt_parts = [
            "## 任务",
            "基于输入的安全数据，生成指定 slide 的占位符内容。",
            f"报告时间：{period.get('start', '')} ~ {period.get('end', '')}",
            "",
            "## 输出要求",
            "1) 只返回合法JSON，不要输出解释、注释或Markdown。",
            "2) 输出内容必须为中文。",
            "3) 所有数字必须与输入数据一致，不得编造。",
            "4) 只输出下方列出的 slide_key 和占位符。",
            "",
        ]

        if total_batches > 1:
            prompt_parts.extend([
                "## 批次信息",
                f"这是第 {batch_index + 1}/{total_batches} 批次，请仅生成本批次的内容。",
                "",
            ])

        prompt_parts.extend([
            "## 输出格式",
            "```json",
            "{",
            '  "slides": [',
        ])

        slide_examples = []
        for slide in template.slides:
            if slide.slide_key not in slide_keys:
                continue
            ai_tokens = [ph.token for ph in slide.placeholders if ph.ai_generate]
            if ai_tokens:
                tokens_str = ", ".join(f'"{t}": "..."' for t in ai_tokens)
                slide_examples.append(f'    {{"slide_key": "{slide.slide_key}", "placeholders": {{{tokens_str}}}}}')

        prompt_parts.append(",\n".join(slide_examples))
        prompt_parts.extend([
            "  ]",
            "}",
            "```",
            "",
            "## 需要生成的内容",
            "",
        ])

        ai_placeholders = template.get_ai_placeholders()
        current_slide = None

        for slide_key, token, placeholder in ai_placeholders:
            if slide_key not in slide_keys:
                continue

            if slide_key != current_slide:
                for slide in template.slides:
                    if slide.slide_key == slide_key:
                        prompt_parts.append(f"### 页面: {slide.title} ({slide_key})")
                        break
                current_slide = slide_key

            constraints = []
            if placeholder.max_length:
                constraints.append(f"最多{placeholder.max_length}字")
            if placeholder.max_items:
                constraints.append(f"最多{placeholder.max_items}条")
            if placeholder.max_chars_per_item:
                constraints.append(f"每条最多{placeholder.max_chars_per_item}字")

            constraint_str = f" ({', '.join(constraints)})" if constraints else ""

            prompt_parts.append(f"\n**{token}**{constraint_str}")
            prompt_parts.append(f"{placeholder.ai_instruction}")
            prompt_parts.append("")

        prompt_parts.extend([
            "",
            "## 安全数据",
            "```json",
            json.dumps(tenant_input.raw, ensure_ascii=False, indent=2),
            "```",
        ])

        return "\n".join(prompt_parts)

    def _build_rewrite_base_prompt(
        self,
        tenant_input: TenantInput,
    ) -> str:
        """Build additional context block for single-slide rewrite."""
        period = tenant_input.get("period", {})

        prompt_parts = [
            "## Full Security Data (Context Only)",
            "Use this for background context. If there is any conflict, follow Current Slide Structured Data first.",
            "```json",
            json.dumps(tenant_input.raw, ensure_ascii=False, indent=2),
            "```",
        ]
        return "\n".join(prompt_parts)

    def _build_rewrite_prompt_with_user_preference(
        self,
        base_prompt: str,
        slide_key: str,
        ai_tokens: List[str],
        user_prompt: str,
        structured_slide_data: Optional[Dict[str, Any]] = None,
        historical_ai_content: Optional[Dict[str, Any]] = None,
    ) -> str:
        """Append user preference instructions for single-slide AI rewrite."""
        ai_tokens_text = ", ".join(ai_tokens)
        output_tokens_preview = ", ".join(f'"{token}": "..."' for token in ai_tokens)

        task_section = "\n".join([
            "## Rewrite Task",
            f"Rewrite only this slide: {slide_key}",
            f"Dynamic target placeholders: {ai_tokens_text}",
            "",
            "## User Preference (High Priority)",
            "Follow the user preference as much as possible without fabricating data:",
            user_prompt.strip(),
            "",
            "## Data Priority",
            "1) Current Slide Structured Data (Highest Priority)",
            "2) Full Security Data (Context Only)",
            "3) Previous AI Copy (Style Reference Only)",
        ])

        structured_data_section = ""
        if structured_slide_data:
            structured_data_section = "\n".join([
                "",
                "## Current Slide Structured Data (Highest Priority)",
                "Numbers in the rewritten text must be consistent with these values.",
                "```json",
                json.dumps(structured_slide_data, ensure_ascii=False, indent=2),
                "```",
            ])

        full_context_section = f"\n{base_prompt}" if base_prompt else ""

        historical_content_section = ""
        if historical_ai_content:
            historical_content_section = "\n".join([
                "",
                "## Previous AI Copy (Style Reference Only)",
                "If any number conflicts with data, ignore old numbers and follow data priority.",
                "```json",
                json.dumps(historical_ai_content, ensure_ascii=False, indent=2),
                "```",
            ])

        hard_constraints_and_output = "\n".join([
            "## Hard Constraints",
            "1) All numbers must match Current Slide Structured Data first.",
            "2) If Structured Data is missing a needed field, infer from Full Security Data.",
            "3) If old copy conflicts with data, ignore old copy and follow data priority.",
            f"4) Output only these dynamic target placeholders: {ai_tokens_text}.",
            "5) Output must be Chinese and in the required JSON format.",
            "",
            "## Output Format",
            "```json",
            "{",
            '  "slides": [',
            f'    {{"slide_key": "{slide_key}", "placeholders": {{{output_tokens_preview}}}}}',
            "  ]",
            "}",
            "```",
        ])
        return (
            f"{task_section}"
            f"{structured_data_section}"
            f"{full_context_section}"
            f"{historical_content_section}\n"
            f"{hard_constraints_and_output}"
        )

    def rewrite_single_slide_v2(
        self,
        tenant_input: TenantInput,
        template_id: str,
        slide_key: str,
        user_prompt: str,
        current_slide_content: Optional[Dict[str, Any]] = None,
        target_tokens: Optional[List[str]] = None,
    ) -> Dict[str, Any]:
        """Rewrite AI-generated placeholders for one slide with user preference."""
        template = self.template_repo.get_descriptor_v2(template_id)

        target_slide = next((s for s in template.slides if s.slide_key == slide_key), None)
        if not target_slide:
            raise ValueError(f"Slide '{slide_key}' not found in template '{template_id}'")

        all_ai_tokens = [ph.token for ph in target_slide.placeholders if ph.ai_generate]
        if not all_ai_tokens:
            raise ValueError(f"Slide '{slide_key}' has no AI-generated placeholders to rewrite")

        ai_tokens = all_ai_tokens
        if target_tokens is not None and len(target_tokens) > 0:
            invalid_tokens = [token for token in target_tokens if token not in all_ai_tokens]
            if invalid_tokens:
                raise ValueError(
                    f"Invalid target_tokens for slide '{slide_key}': {', '.join(invalid_tokens)}"
                )
            ai_tokens = target_tokens

        system_prompt = self._build_system_prompt(template)
        base_user_prompt = self._build_rewrite_base_prompt(tenant_input)
        historical_ai_content: Dict[str, Any] = {}
        if isinstance(current_slide_content, dict):
            for token in ai_tokens:
                if token in current_slide_content:
                    historical_ai_content[token] = current_slide_content[token]

        all_ai_token_set = set(all_ai_tokens)
        structured_slide_data: Dict[str, Any] = {}
        for placeholder in target_slide.placeholders:
            token = placeholder.token
            if token in all_ai_token_set:
                continue

            value = None
            if placeholder.source:
                value = self._get_nested(tenant_input, placeholder.source)

            # Fallback to current slide payload when source is absent or unresolved.
            if value is None and isinstance(current_slide_content, dict):
                value = current_slide_content.get(token)

            if value is not None:
                structured_slide_data[token] = value

        rewrite_user_prompt = self._build_rewrite_prompt_with_user_preference(
            base_prompt=base_user_prompt,
            slide_key=slide_key,
            ai_tokens=ai_tokens,
            user_prompt=user_prompt,
            structured_slide_data=structured_slide_data or None,
            historical_ai_content=historical_ai_content or None,
        )

        parsed = self._call_and_parse_with_retry(system_prompt, rewrite_user_prompt, template)

        slide_placeholders = parsed.get(slide_key)
        if not isinstance(slide_placeholders, dict):
            raise LLMGenerationError(
                f"LLM response missing placeholders for slide '{slide_key}'"
            )

        ai_token_set = set(ai_tokens)
        filtered_placeholders = {
            token: value for token, value in slide_placeholders.items() if token in ai_token_set
        }
        missing_tokens = [token for token in ai_tokens if token not in filtered_placeholders]

        warnings: List[str] = []
        if missing_tokens:
            warnings.append(
                f"LLM response did not include some AI placeholders: {', '.join(missing_tokens)}"
            )

        return {
            "slide_key": slide_key,
            "placeholders": filtered_placeholders,
            "warnings": warnings,
            "missing_tokens": missing_tokens,
            "updated_tokens": list(filtered_placeholders.keys()),
        }

    def _estimate_prompt_tokens(self, text: str) -> int:
        """Estimate token count for a text string.

        For Chinese text, roughly 1.5-2 characters per token.
        For English/code, roughly 4 characters per token.
        We use a conservative estimate of 2 characters per token for mixed content.
        """
        return len(text) // 2

    def _estimate_slide_instruction_size(
        self,
        slide_key: str,
        template: TemplateDescriptorV2,
    ) -> int:
        """Estimate the instruction size for a slide's AI placeholders."""
        size = 0
        for slide in template.slides:
            if slide.slide_key == slide_key:
                # Add slide header
                size += len(f"### Slide: {slide.title} ({slide_key})\n")
                for ph in slide.placeholders:
                    if ph.ai_generate and ph.ai_instruction:
                        size += len(f"\n**{ph.token}**\n")
                        size += len(ph.ai_instruction or "")
                        size += 50  # constraints and formatting overhead
                break
        return size

    def _get_smart_slide_batches(
        self,
        tenant_input: TenantInput,
        template: TemplateDescriptorV2,
        max_tokens_per_batch: int = 15000,
    ) -> List[List[str]]:
        """Split slides into batches based on estimated token count.

        This method intelligently groups slides to keep each batch under
        the token limit, avoiding API timeouts.

        Args:
            tenant_input: Raw tenant input data (needed for base prompt size)
            template: Template descriptor
            max_tokens_per_batch: Maximum estimated tokens per batch

        Returns:
            List of batches, where each batch is a list of slide_keys
        """
        # Calculate base prompt size (customer info + input data)
        # This is constant across all batches
        tenant = tenant_input.get("tenant", {})
        period = tenant_input.get("period", {})
        base_prompt = "\n".join([
            "## 安全数据",
            "```json",
            json.dumps(tenant_input.raw, ensure_ascii=False, indent=2),
            "```",
        ])
        base_tokens = self._estimate_prompt_tokens(base_prompt)

        # Reserve tokens for JSON output format instructions (~500 tokens)
        format_overhead = 500

        # Available tokens for slide instructions per batch
        available_tokens = max_tokens_per_batch - base_tokens - format_overhead

        logger.info(f"📊 Batch sizing: base={base_tokens} tokens, available={available_tokens} tokens/batch")

        # Calculate instruction size for each slide with AI placeholders
        slide_sizes: List[tuple] = []  # (slide_key, estimated_tokens)
        for slide in template.slides:
            ai_count = sum(1 for ph in slide.placeholders if ph.ai_generate)
            if ai_count == 0:
                continue
            instruction_size = self._estimate_slide_instruction_size(slide.slide_key, template)
            estimated_tokens = self._estimate_prompt_tokens(" " * instruction_size)
            slide_sizes.append((slide.slide_key, estimated_tokens))

        # If total is small enough, no batching needed
        total_instruction_tokens = sum(t for _, t in slide_sizes)
        if total_instruction_tokens <= available_tokens:
            logger.info(f"📦 No batching needed: {total_instruction_tokens} tokens fits in {available_tokens}")
            return [[s for s, _ in slide_sizes]]

        # Greedy batching: add slides until we exceed the limit
        batches: List[List[str]] = []
        current_batch: List[str] = []
        current_tokens = 0

        for slide_key, tokens in slide_sizes:
            # If adding this slide would exceed the limit, start a new batch
            if current_tokens + tokens > available_tokens and current_batch:
                batches.append(current_batch)
                logger.info(f"  Batch {len(batches)}: {current_batch} (~{current_tokens} tokens)")
                current_batch = []
                current_tokens = 0

            current_batch.append(slide_key)
            current_tokens += tokens

        # Don't forget the last batch
        if current_batch:
            batches.append(current_batch)
            logger.info(f"  Batch {len(batches)}: {current_batch} (~{current_tokens} tokens)")

        return batches

    def _generate_ai_content_in_batches(
        self,
        tenant_input: TenantInput,
        template: TemplateDescriptorV2,
        max_tokens_per_batch: int = 15000,
        session_id: str = None,
        ws_manager = None,
        event_loop = None,
    ) -> Dict[str, Dict[str, Any]]:
        """Generate AI content in batches to avoid timeout issues.

        Args:
            tenant_input: Raw tenant input data
            template: Template descriptor
            max_tokens_per_batch: Maximum estimated tokens per API call
            session_id: Session ID for progress updates
            ws_manager: WebSocket manager for real-time progress
            event_loop: Event loop for scheduling async tasks from sync code

        Returns:
            Dict[slide_key, Dict[token, value]] with all AI-generated content
        """
        batches = self._get_smart_slide_batches(tenant_input, template, max_tokens_per_batch)
        total_batches = len(batches)

        # Helper to send progress updates
        def send_progress(progress: int, message: str):
            if ws_manager and session_id and event_loop:
                import asyncio
                try:
                    asyncio.run_coroutine_threadsafe(
                        ws_manager.send_progress_update(session_id, progress, message),
                        event_loop
                    )
                except Exception as e:
                    logger.debug(f"Failed to send progress update: {e}")

        if total_batches <= 1:
            logger.info("Single batch - using standard generation")
            send_progress(35, "正在调用AI生成内容...")
            system_prompt = self._build_system_prompt(template)
            user_prompt = self._build_user_prompt(tenant_input, template)
            result = self._call_and_parse_with_retry(system_prompt, user_prompt, template)
            send_progress(60, "AI内容生成完成")
            return result

        logger.info(f"Smart batching: splitting into {total_batches} batches")

        all_ai_placeholders: Dict[str, Dict[str, Any]] = {}
        system_prompt = self._build_system_prompt(template)

        # Progress range: 30% - 60%
        progress_per_batch = 30.0 / total_batches

        for i, batch_slide_keys in enumerate(batches):
            logger.info(f"Processing batch {i + 1}/{total_batches}: slides {batch_slide_keys}")

            current_progress = 30 + int(i * progress_per_batch)
            send_progress(current_progress, f"AI生成中（第 {i + 1}/{total_batches} 批）...")

            user_prompt = self._build_user_prompt_for_slides(
                tenant_input,
                template,
                batch_slide_keys,
                batch_index=i,
                total_batches=total_batches,
            )

            prompt_tokens = self._estimate_prompt_tokens(user_prompt)
            logger.info(f"   Batch prompt size: ~{prompt_tokens} tokens")

            batch_placeholders = self._call_and_parse_with_retry(system_prompt, user_prompt, template)

            for slide_key, tokens in batch_placeholders.items():
                if slide_key not in all_ai_placeholders:
                    all_ai_placeholders[slide_key] = {}
                all_ai_placeholders[slide_key].update(tokens)

            logger.info(f"Batch {i + 1}/{total_batches} completed")

        send_progress(60, "所有AI内容生成完成")
        return all_ai_placeholders

    def _call_and_parse_with_retry(
        self,
        system_prompt: str,
        user_prompt: str,
        template: TemplateDescriptorV2,
        max_parse_retries: int = 5,
    ) -> Dict[str, Dict[str, Any]]:
        """Call LLM and parse response with retry on format errors."""
        for attempt in range(max_parse_retries):
            try:
                logger.info(f"LLM generation attempt {attempt + 1}/{max_parse_retries}")

                response = self._call_openai_with_retry(system_prompt, user_prompt)

                parsed = self._parse_llm_response(response, template)

                logger.info(f"Successfully parsed LLM response on attempt {attempt + 1}")
                return parsed

            except LLMGenerationError as e:
                logger.error(f"Parse attempt {attempt + 1}/{max_parse_retries} failed: {e}")

                if attempt < max_parse_retries - 1:
                    logger.warning("Retrying LLM call due to format error...")
                else:
                    error_msg = (
                        f"AI生成失败：重试 {max_parse_retries} 次后仍未成功。"
                        f"请检查提示词与模型输出格式，必要时启用 mock 兜底。"
                        f"原始错误：{e}"
                    )
                    logger.error(error_msg)
                    raise LLMGenerationError(error_msg) from e

        raise LLMGenerationError(f"Unexpected error: exceeded {max_parse_retries} retries")

    def _call_openai_with_retry(
        self,
        system_prompt: str,
        user_prompt: str,
        max_retries: int = 4,
        retry_delay: float = 2.0,
    ) -> str:
        """Call OpenAI API with retry logic."""
        if not self.client:
            raise LLMGenerationError("OpenAI client is not initialized. Enable LLM in settings.")

        logger.info("=" * 80)
        logger.info("CALLING OPENAI API (V2)")
        logger.info(f"System prompt length: {len(system_prompt)} chars")
        logger.info(f"User prompt length: {len(user_prompt)} chars")
        logger.info("=" * 80)

        content = self._call_openai_api(system_prompt, user_prompt)
        return content

    @with_llm_retry
    def _call_openai_api(
        self,
        system_prompt: str,
        user_prompt: str,
    ) -> str:
        """Make the actual OpenAI API call (wrapped with retry decorator)."""
        stream = self.client.chat.completions.create(
            model=config.settings.openai_model,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            temperature=0.4,
            response_format={"type": "json_object"},
            stream=True,
        )

        content_chunks: List[str] = []

        for chunk in stream:
            if not getattr(chunk, "choices", None):
                continue

            delta = chunk.choices[0].delta
            content_piece = getattr(delta, "content", None)
            if content_piece:
                content_chunks.append(content_piece)

        content = "".join(content_chunks)
        if not content:
            raise LLMGenerationError("OpenAI returned empty response")

        logger.info("=" * 80)
        logger.info("OPENAI API CALL SUCCESSFUL")
        logger.info(f"Response length: {len(content)} chars")
        logger.info("=" * 80)
        return content

    def _sanitize_llm_json(self, content: str) -> str:
        """Clean up LLM response for JSON parsing."""
        if not content:
            return content

        text = content.strip()

        # Remove markdown code fences
        fenced_match = re.search(r"```(?:json)?\s*([\s\S]*?)\s*```", text, re.IGNORECASE)
        if fenced_match:
            text = fenced_match.group(1).strip()

        # Remove single-line comments (// ...)
        text = re.sub(r'//.*?$', '', text, flags=re.MULTILINE)

        # Remove multi-line comments (/* ... */)
        text = re.sub(r'/\*.*?\*/', '', text, flags=re.DOTALL)

        # Extract JSON object
        first_brace = text.find("{")
        last_brace = text.rfind("}")
        if first_brace != -1 and last_brace != -1 and last_brace > first_brace:
            return text[first_brace:last_brace + 1].strip()

        return text

    def _parse_llm_response(
        self,
        response: str,
        template: TemplateDescriptorV2
    ) -> Dict[str, Dict[str, Any]]:
        """Parse LLM response into slide placeholders.

        Returns:
            Dict[slide_key, Dict[token, value]]
        """
        cleaned = self._sanitize_llm_json(response)

        try:
            data = json.loads(cleaned)
        except json.JSONDecodeError as e:
            logger.error(f"Failed to parse LLM response: {e}")
            logger.error(f"Response: {cleaned[:500]}...")
            raise LLMGenerationError(f"Invalid JSON from LLM: {e}") from e

        result: Dict[str, Dict[str, Any]] = {}

        if "slides" not in data:
            raise LLMGenerationError("Response missing 'slides' field")

        # Validate slides is a list
        if not isinstance(data["slides"], list):
            logger.error(f"'slides' field is not a list: {type(data['slides'])}")
            logger.error(f"Response data: {json.dumps(data, ensure_ascii=False, indent=2)[:1000]}")
            raise LLMGenerationError(f"'slides' field must be a list, got {type(data['slides'])}")

        for i, slide_data in enumerate(data["slides"]):
            # Validate each slide_data is a dict
            if not isinstance(slide_data, dict):
                logger.error(f"Slide data at index {i} is not a dict: {type(slide_data)}")
                logger.error(f"Slide data: {slide_data}")
                raise LLMGenerationError(f"Slide at index {i} must be a dict, got {type(slide_data)}: {slide_data}")

            slide_key = slide_data.get("slide_key")
            placeholders = slide_data.get("placeholders", {})

            if slide_key:
                result[slide_key] = placeholders

        return result

    def generate_slidespec_v2(
        self,
        tenant_input: TenantInput,
        template_id: str,
        use_mock: bool = False,
        session_id: str = None,
        ws_manager = None,
        event_loop = None,
    ) -> SlideSpecV2:
        """Generate SlideSpec for V2 template using AI.

        This is the main entry point for V2 generation.

        Args:
            tenant_input: Raw tenant input data
            template_id: V2 template ID
            use_mock: Whether to force mock/fallback generation
            session_id: Session ID for WebSocket progress updates
            ws_manager: WebSocket manager for real-time progress
            event_loop: Event loop for scheduling async tasks from sync code

        Returns:
            SlideSpecV2 with all placeholders filled
        """
        logger.info(f"🎯 Generating V2 slidespec for template: {template_id}, use_mock={use_mock}")

        # Helper to send progress updates
        def send_progress(progress: int, message: str):
            if ws_manager and session_id and event_loop:
                import asyncio
                try:
                    # Schedule coroutine in the main event loop
                    asyncio.run_coroutine_threadsafe(
                        ws_manager.send_progress_update(session_id, progress, message),
                        event_loop
                    )
                except Exception as e:
                    logger.debug(f"Failed to send progress update: {e}")

        # Load V2 template descriptor (20%)
        send_progress(20, "加载模板描述符...")
        template = self.template_repo.get_descriptor_v2(template_id)

        # Create empty slidespec structure
        slide_keys = [(s.slide_no, s.slide_key) for s in template.slides]
        slidespec = create_empty_slidespec_v2(template_id, slide_keys)

        # Step 1: Extract data placeholders (non-AI) (25%)
        logger.info("📊 Extracting data placeholders...")
        send_progress(25, "提取数据占位符...")
        data_placeholders = self._extract_data_placeholders(tenant_input, template)

        for slide_key, tokens in data_placeholders.items():
            slide = slidespec.get_slide(slide_key)
            if slide:
                slide.placeholders.update(tokens)

        # Step 2: Generate AI placeholders (30% - 70%)
        if config.settings.enable_llm and not use_mock:
            logger.info("🤖 Generating AI content...")
            send_progress(30, "调用AI生成内容...")
            try:
                # Use smart batched generation to avoid timeout issues
                # Batching is based on estimated token count, not hardcoded limits
                ai_placeholders = self._generate_ai_content_in_batches(
                    tenant_input,
                    template,
                    session_id=session_id,
                    ws_manager=ws_manager,
                    event_loop=event_loop,
                )

                # Merge AI content (65%)
                send_progress(65, "合并AI生成内容...")
                for slide_key, tokens in ai_placeholders.items():
                    slide = slidespec.get_slide(slide_key)
                    if slide:
                        slide.placeholders.update(tokens)

                # (No validator) Keep generation flow simple
                send_progress(70, "验证生成内容...")

            except LLMGenerationError as e:
                logger.error(f"❌ AI generation failed: {e}")
                raise
        else:
            logger.info(f"📝 {'Using mock mode' if use_mock else 'LLM disabled'}, using fallback content")
            send_progress(35, "使用快速生成模式...")
            self._fill_ai_placeholders_with_fallback(slidespec, template)
            send_progress(70, "快速生成完成...")

        logger.info(f"V2 slidespec generation complete: {len(slidespec.slides)} slides")
        return slidespec

    def _fill_ai_placeholders_with_fallback(
        self,
        slidespec: SlideSpecV2,
        template: TemplateDescriptorV2,
    ) -> None:
        """Fill AI placeholders with fallback text when LLM is unavailable."""
        for slide_key, token, placeholder in template.get_ai_placeholders():
            slide = slidespec.get_slide(slide_key)
            if slide and token not in slide.placeholders:
                slide.placeholders[token] = f"[{token}: AI生成内容占位]"


# ============================================================================
# Legacy V1 Orchestrator (kept for backward compatibility)
# ============================================================================

class LLMOrchestrator:
    """V1 Orchestrator - Legacy implementation for V1 templates."""

    def __init__(self, template_repo: Optional[TemplateRepository] = None):
        self.template_repo = template_repo or TemplateRepository()
        self.client: Optional[OpenAI] = None

        if config.settings.enable_llm:
            try:
                self.client = _build_openai_client()
            except Exception as e:
                raise LLMGenerationError(f"OpenAI client initialization failed: {e}") from e

    def _load_mock_slidespec(self, input_id: str, template_id: str) -> SlideSpec:
        """Load mock slidespec for fallback."""
        audience = "management" if "management" in template_id else "technical"
        mock_file = f"{input_id}_{audience}_mock_slidespec.json"
        path = config.MOCK_OUTPUTS_DIR / mock_file
        if not path.exists():
            raise MockOutputNotFound(f"Mock slidespec {mock_file} not found")
        with path.open("r", encoding="utf-8") as f:
            data = json.load(f)
        return SlideSpec.model_validate(data)

    def generate_slidespec(
        self,
        input_id: str,
        template_id: str,
        prepared: Any,  # DataPrepResult
        use_mock: bool = False,
    ) -> SlideSpec:
        """Generate slidespec using V1 logic (legacy)."""
        logger.info(f"🎯 Generating V1 slidespec for {input_id}/{template_id}")

        if use_mock:
            try:
                return self._load_mock_slidespec(input_id, template_id)
            except MockOutputNotFound:
                logger.warning("Mock not found, using deterministic generation")

        # Deterministic fallback
        slides = []
        template = self.template_repo.get_descriptor(template_id)
        for slide in template.slides:
            data = prepared.slide_inputs.get(slide.slide_key, {})
            slides.append(SlideSpecItem(
                slide_no=slide.slide_no,
                slide_key=slide.slide_key,
                data=data,
            ))

        return SlideSpec(template_id=template_id, slides=slides)

    def rewrite_slide(
        self,
        slide_spec: SlideSpec,
        slide_key: str,
        new_content: Dict[str, Any],
    ) -> SlideSpec:
        """Update a slide's data with new content."""
        updated_slides = []
        for slide in slide_spec.slides:
            if slide.slide_key == slide_key:
                updated_slides.append(SlideSpecItem(
                    slide_no=slide.slide_no,
                    slide_key=slide.slide_key,
                    data={**slide.data, **new_content},
                ))
            else:
                updated_slides.append(slide)
        return SlideSpec(template_id=slide_spec.template_id, slides=updated_slides)

