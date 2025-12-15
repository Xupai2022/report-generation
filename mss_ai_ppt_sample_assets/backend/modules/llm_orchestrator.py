from __future__ import annotations

import json
import logging
import re
import time
from typing import Any, Dict, List, Optional

from openai import OpenAI, APIError, RateLimitError, APIConnectionError

from mss_ai_ppt_sample_assets.backend import config
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
                client_kwargs = {"api_key": config.settings.openai_api_key}
                if config.settings.openai_base_url:
                    client_kwargs["base_url"] = config.settings.openai_base_url
                self.client = OpenAI(**client_kwargs)
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
            chart_type: 'bar_chart' or 'pie_chart'

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

        result = {
            'position': chart_config.get('position')
        }

        if chart_type == 'bar_chart':
            # Expect source_data to have 'labels' and 'values' or similar structure
            x_field = chart_config.get('x_field', 'labels')
            y_field = chart_config.get('y_field', 'values')

            if isinstance(source_data, dict):
                categories = source_data.get(x_field, [])
                values = source_data.get(y_field, [])

                result['categories'] = categories
                result['series'] = [{'name': chart_config.get('series_name', '告警数'), 'values': values}]
            else:
                logger.warning(f"Bar chart data source {data_source} is not a dict")
                return {}

        elif chart_type == 'pie_chart':
            # Expect source_data to be a dict like {'high': 52, 'medium': 473, 'low': 816}
            if isinstance(source_data, dict):
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
            if placeholder.type in ('bar_chart', 'pie_chart') and placeholder.chart_config:
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
2. **深度分析**：不要只是简单罗列数据，要提供有洞察力的分析和解读
3. **具体建议**：整改建议必须具体可执行，避免泛泛而谈
4. **语言风格**：使用中文，简洁专业，适合{audience_desc}阅读
5. **格式要求**：严格按照指定的JSON格式返回内容

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
        tenant = tenant_input.get("tenant", {})
        period = tenant_input.get("period", {})

        prompt_parts = [
            "## 客户信息",
            f"- 客户名称：{tenant.get('name', '')}",
            f"- 行业：{tenant.get('industry', '')}",
            f"- 地区：{tenant.get('region', '')}",
            "",
            "## 报告周期",
            f"- 开始日期：{period.get('start', '')}",
            f"- 结束日期：{period.get('end', '')}",
            "",
            "## 安全数据",
            "```json",
            json.dumps(tenant_input.raw, ensure_ascii=False, indent=2),
            "```",
            "",
            "## 需要生成的内容",
            "",
        ]

        # Add AI placeholder instructions
        ai_placeholders = template.get_ai_placeholders()
        current_slide = None

        for slide_key, token, placeholder in ai_placeholders:
            if slide_key != current_slide:
                # Find slide title
                for slide in template.slides:
                    if slide.slide_key == slide_key:
                        prompt_parts.append(f"### Slide: {slide.title} ({slide_key})")
                        break
                current_slide = slide_key

            # Add placeholder instruction
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

        # Add output format reminder
        prompt_parts.extend([
            "",
            "## 请按以下JSON格式返回：",
            "```json",
            "{",
            '  "slides": [',
        ])

        # Generate expected structure
        slide_examples = []
        for slide in template.slides:
            ai_tokens = [ph.token for ph in slide.placeholders if ph.ai_generate]
            if ai_tokens:
                tokens_str = ", ".join(f'"{t}": "..."' for t in ai_tokens)
                slide_examples.append(f'    {{"slide_key": "{slide.slide_key}", "placeholders": {{{tokens_str}}}}}')

        prompt_parts.append(",\n".join(slide_examples))
        prompt_parts.extend([
            "  ],",
            "}",
            "```",
        ])

        return "\n".join(prompt_parts)

    def _call_openai_with_retry(
        self,
        system_prompt: str,
        user_prompt: str,
        max_retries: int = 3,
        retry_delay: float = 2.0,
    ) -> str:
        """Call OpenAI API with retry logic."""
        if not self.client:
            raise LLMGenerationError("OpenAI client is not initialized. Enable LLM in settings.")

        logger.info("=" * 80)
        logger.info("CALLING OPENAI API (V2)")
        logger.info(f"Model: {config.settings.openai_model}")
        logger.info(f"System prompt length: {len(system_prompt)} chars")
        logger.info(f"User prompt length: {len(user_prompt)} chars")
        logger.info("=" * 80)

        last_error = None
        for attempt in range(max_retries):
            try:
                logger.info(f"🔄 API Call Attempt {attempt + 1}/{max_retries}")

                response = self.client.chat.completions.create(
                    model=config.settings.openai_model,
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": user_prompt}
                    ],
                    temperature=0.7,
                    response_format={"type": "json_object"},
                )

                content = response.choices[0].message.content
                if not content:
                    raise LLMGenerationError("OpenAI returned empty response")

                logger.info("=" * 80)
                logger.info("✅ OPENAI API CALL SUCCESSFUL")
                logger.info(f"Total tokens used: {response.usage.total_tokens}")
                logger.info(f"Response length: {len(content)} chars")
                logger.info("=" * 80)
                return content

            except RateLimitError as e:
                last_error = e
                if attempt < max_retries - 1:
                    wait_time = retry_delay * (2 ** attempt)
                    logger.warning(f"⚠️ Rate limit hit, waiting {wait_time}s...")
                    time.sleep(wait_time)

            except APIConnectionError as e:
                last_error = e
                if attempt < max_retries - 1:
                    logger.warning(f"⚠️ Connection error: {e}, retrying...")
                    time.sleep(retry_delay)

            except APIError as e:
                last_error = e
                logger.error(f"❌ OpenAI API error: {e}")
                break

            except Exception as e:
                last_error = e
                logger.error(f"❌ Unexpected error: {e}")
                break

        raise LLMGenerationError(f"Failed after {max_retries} attempts: {last_error}") from last_error

    def _sanitize_llm_json(self, content: str) -> str:
        """Clean up LLM response for JSON parsing."""
        if not content:
            return content

        text = content.strip()

        # Remove markdown code fences
        fenced_match = re.search(r"```(?:json)?\s*([\s\S]*?)\s*```", text, re.IGNORECASE)
        if fenced_match:
            text = fenced_match.group(1).strip()

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

        for slide_data in data["slides"]:
            slide_key = slide_data.get("slide_key")
            placeholders = slide_data.get("placeholders", {})

            if slide_key:
                result[slide_key] = placeholders

        return result

    def _validate_key_numbers(
        self,
        slidespec: SlideSpecV2,
        tenant_input: TenantInput,
        template: TemplateDescriptorV2
    ) -> List[str]:
        """Validate key numerical fields match input data.

        Returns:
            List of validation error messages (empty if all valid)
        """
        errors = []
        validation_fields = template.get_validation_fields()

        # Computed values
        incidents = tenant_input.get("incidents", []) or []
        computed = {
            "incidents_count": len(incidents),
            "incidents_high_count": len([i for i in incidents if i.get("severity") == "high"]),
        }

        for token, field_path in validation_fields.items():
            # Get expected value
            if field_path in computed:
                expected = computed[field_path]
            else:
                expected = self._get_nested(tenant_input, field_path)

            if expected is None:
                continue

            # Find actual value in slidespec
            for slide in slidespec.slides:
                if token in slide.placeholders:
                    actual = slide.placeholders[token]

                    # Extract number from string if needed
                    if isinstance(actual, str):
                        numbers = re.findall(r'[\d.]+', actual)
                        if numbers:
                            try:
                                actual = float(numbers[0]) if '.' in numbers[0] else int(numbers[0])
                            except ValueError:
                                pass

                    # Compare
                    if isinstance(expected, (int, float)) and isinstance(actual, (int, float)):
                        if abs(expected - actual) > 0.01:
                            errors.append(f"{token}: expected {expected}, got {actual}")

        return errors

    def generate_slidespec_v2(
        self,
        tenant_input: TenantInput,
        template_id: str,
        use_mock: bool = False,
    ) -> SlideSpecV2:
        """Generate SlideSpec for V2 template using AI.

        This is the main entry point for V2 generation.

        Args:
            tenant_input: Raw tenant input data
            template_id: V2 template ID
            use_mock: Whether to force mock/fallback generation

        Returns:
            SlideSpecV2 with all placeholders filled
        """
        logger.info(f"🎯 Generating V2 slidespec for template: {template_id}, use_mock={use_mock}")

        # Load V2 template descriptor
        template = self.template_repo.get_descriptor_v2(template_id)

        # Create empty slidespec structure
        slide_keys = [(s.slide_no, s.slide_key) for s in template.slides]
        slidespec = create_empty_slidespec_v2(template_id, slide_keys)

        # Step 1: Extract data placeholders (non-AI)
        logger.info("📊 Extracting data placeholders...")
        data_placeholders = self._extract_data_placeholders(tenant_input, template)

        for slide_key, tokens in data_placeholders.items():
            slide = slidespec.get_slide(slide_key)
            if slide:
                slide.placeholders.update(tokens)

        # Step 2: Generate AI placeholders
        if config.settings.enable_llm and not use_mock:
            logger.info("🤖 Generating AI content...")
            try:
                system_prompt = self._build_system_prompt(template)
                user_prompt = self._build_user_prompt(tenant_input, template)

                response = self._call_openai_with_retry(system_prompt, user_prompt)
                ai_placeholders = self._parse_llm_response(response, template)

                # Merge AI content
                for slide_key, tokens in ai_placeholders.items():
                    slide = slidespec.get_slide(slide_key)
                    if slide:
                        slide.placeholders.update(tokens)

                # Validate key numbers
                errors = self._validate_key_numbers(slidespec, tenant_input, template)
                if errors:
                    logger.warning(f"⚠️ Validation warnings: {errors}")

            except LLMGenerationError as e:
                logger.error(f"❌ AI generation failed: {e}")
                logger.warning("⚠️ Falling back to placeholder text")
                self._fill_ai_placeholders_with_fallback(slidespec, template)
        else:
            logger.info(f"📝 {'Using mock mode' if use_mock else 'LLM disabled'}, using fallback content")
            self._fill_ai_placeholders_with_fallback(slidespec, template)

        logger.info(f"✅ V2 slidespec generation complete: {len(slidespec.slides)} slides")
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
                if placeholder.type == "bullet_list":
                    slide.placeholders[token] = "• [AI生成内容占位]\n• [请启用LLM以生成实际内容]"
                else:
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
                client_kwargs = {"api_key": config.settings.openai_api_key}
                if config.settings.openai_base_url:
                    client_kwargs["base_url"] = config.settings.openai_base_url
                self.client = OpenAI(**client_kwargs)
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
