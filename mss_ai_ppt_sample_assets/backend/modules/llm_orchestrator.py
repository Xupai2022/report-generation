from __future__ import annotations

import json
import logging
import re
from typing import Any, Dict, Optional
from openai import OpenAI, APIError, RateLimitError, APIConnectionError
import time

from mss_ai_ppt_sample_assets.backend import config
from mss_ai_ppt_sample_assets.backend.models.slidespec import SlideSpec, SlideSpecItem
from mss_ai_ppt_sample_assets.backend.modules.template_loader import TemplateRepository
from mss_ai_ppt_sample_assets.backend.modules.data_prep import DataPrepResult

# Configure logging
logger = logging.getLogger(__name__)


class MockOutputNotFound(Exception):
    pass


class LLMGenerationError(Exception):
    """Error during LLM content generation"""
    pass


class LLMOrchestrator:
    """Orchestrates content generation using OpenAI API."""

    def __init__(self, template_repo: Optional[TemplateRepository] = None):
        self.template_repo = template_repo or TemplateRepository()
        self.client: Optional[OpenAI] = None

        # Initialize OpenAI client if LLM is enabled
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

    def _load_mock_slidespec(self, input_id: str, template_id: str) -> SlideSpec:
        """Load mock slidespec for fallback (deprecated)."""
        audience = "management" if "management" in template_id else "technical"
        mock_file = f"{input_id}_{audience}_mock_slidespec.json"
        path = config.MOCK_OUTPUTS_DIR / mock_file
        if not path.exists():
            raise MockOutputNotFound(f"Mock slidespec {mock_file} not found")
        with path.open("r", encoding="utf-8") as f:
            data = json.load(f)
        return SlideSpec.model_validate(data)

    def _build_generation_prompt(
        self,
        template_id: str,
        prepared: DataPrepResult,
    ) -> str:
        """Build the prompt for OpenAI to generate slidespec content."""
        template = self.template_repo.get_descriptor(template_id)

        prompt_parts = [
            "You are an expert security analyst generating content for a managed security service (MSS) report.",
            f"\n## Report Template: {template.name}",
            f"Template ID: {template_id}",
            f"Audience: {'Management' if 'management' in template_id else 'Technical'}",
            f"\n## Input Data:",
            json.dumps(prepared.facts, ensure_ascii=False, indent=2),
            f"\n## Slide Definitions:",
        ]

        # Add slide schemas and requirements
        for slide in template.slides:
            prompt_parts.append(f"\n### Slide {slide.slide_no}: {slide.slide_key}")
            prompt_parts.append(f"Description: {slide.description}")
            prompt_parts.append(f"Schema: {json.dumps(slide.schema, ensure_ascii=False, indent=2)}")

            # Include prepared slide inputs if available
            if slide.slide_key in prepared.slide_inputs:
                prompt_parts.append(f"Prepared Input: {json.dumps(prepared.slide_inputs[slide.slide_key], ensure_ascii=False, indent=2)}")

        prompt_parts.extend([
            "\n## Task:",
            "Generate a complete slidespec JSON that follows this structure:",
            "{",
            f'  "template_id": "{template_id}",',
            '  "slides": [',
            '    {',
            '      "slide_no": 1,',
            '      "slide_key": "cover",',
            '      "data": { /* slide content matching schema */ }',
            '    },',
            '    ...',
            '  ]',
            '}',
            "\n## Requirements:",
            "1. Generate content for ALL slides defined in the template",
            "2. Each slide's 'data' must match its schema exactly",
            "3. Use the input data provided to fill in facts and metrics",
            "4. For management audience: focus on high-level insights, KPIs, and business impact",
            "5. For technical audience: include detailed technical findings and recommendations",
            "6. All text should be in Chinese (zh-CN) unless field names require English",
            "7. Return ONLY valid JSON, no additional explanation",
            "\nGenerate the complete slidespec JSON now:"
        ])

        return "\n".join(prompt_parts)

    def _call_openai_with_retry(
        self,
        prompt: str,
        max_retries: int = 3,
        retry_delay: float = 2.0,
    ) -> str:
        """Call OpenAI API with retry logic for transient errors."""
        if not self.client:
            raise LLMGenerationError("OpenAI client is not initialized. Enable LLM in settings.")

        logger.info("=" * 80)
        logger.info("CALLING OPENAI API")
        logger.info(f"Model: {config.settings.openai_model}")
        logger.info(f"Base URL: {config.settings.openai_base_url}")
        logger.info(f"Prompt length: {len(prompt)} chars")
        logger.info("=" * 80)

        last_error = None
        for attempt in range(max_retries):
            try:
                logger.info(f"🔄 API Call Attempt {attempt + 1}/{max_retries}")

                response = self.client.chat.completions.create(
                    model=config.settings.openai_model,
                    messages=[
                        {
                            "role": "system",
                            "content": "You are an expert security analyst. Always respond with valid JSON only."
                        },
                        {
                            "role": "user",
                            "content": prompt
                        }
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
                logger.info(f"Response preview: {content[:200]}...")
                logger.info("=" * 80)
                return content

            except RateLimitError as e:
                last_error = e
                if attempt < max_retries - 1:
                    wait_time = retry_delay * (2 ** attempt)  # Exponential backoff
                    logger.warning(f"⚠️  Rate limit hit, waiting {wait_time}s before retry...")
                    time.sleep(wait_time)
                else:
                    logger.error(f"❌ Rate limit exceeded after {max_retries} attempts")

            except APIConnectionError as e:
                last_error = e
                if attempt < max_retries - 1:
                    logger.warning(f"⚠️  Connection error: {e}, retrying in {retry_delay}s...")
                    time.sleep(retry_delay)
                else:
                    logger.error(f"❌ Connection failed after {max_retries} attempts: {e}")

            except APIError as e:
                last_error = e
                logger.error(f"❌ OpenAI API error: {e}")
                logger.error(f"Error details: {e.__dict__}")
                # Don't retry on API errors (usually client errors)
                break

            except Exception as e:
                last_error = e
                logger.error(f"❌ Unexpected error calling OpenAI: {e}")
                logger.exception("Full traceback:")
                break

        raise LLMGenerationError(f"Failed to generate content after {max_retries} attempts: {last_error}") from last_error

    def _parse_llm_response(self, response_content: str, template_id: str) -> SlideSpec:
        """Parse and validate LLM response into SlideSpec."""
        logger.info("📝 Parsing LLM response...")
        logger.debug(f"Raw response: {response_content[:500]}...")

        try:
            cleaned = self._sanitize_llm_json(response_content)
            if cleaned != response_content:
                logger.info("Sanitized LLM response before JSON parse")
            data = json.loads(cleaned)
            logger.info(f"✓ JSON parsed successfully")

            # Validate structure
            if "slides" not in data:
                raise ValueError("Response missing 'slides' field")

            logger.info(f"✓ Found 'slides' field with {len(data.get('slides', []))} slides")

            # Ensure template_id is set
            if "template_id" not in data:
                data["template_id"] = template_id
                logger.info(f"✓ Added template_id: {template_id}")

            # Parse into SlideSpec model
            slidespec = SlideSpec.model_validate(data)
            logger.info(f"✅ Successfully parsed slidespec with {len(slidespec.slides)} slides")

            # Log slide summary
            for idx, slide in enumerate(slidespec.slides, 1):
                logger.info(f"  Slide {idx}: {slide.slide_key} (slide_no={slide.slide_no})")

            return slidespec

        except json.JSONDecodeError as e:
            logger.error(f"❌ Invalid JSON from LLM: {e}")
            logger.error(f"Response content: {response_content[:1000]}")
            raise LLMGenerationError(f"LLM returned invalid JSON: {e}") from e
        except Exception as e:
            logger.error(f"❌ Failed to parse LLM response: {e}")
            logger.exception("Full traceback:")
            raise LLMGenerationError(f"Failed to parse LLM response: {e}") from e

    def _sanitize_llm_json(self, content: str) -> str:
        """Best-effort cleanup for models that wrap JSON in code fences or extra text."""
        if not content:
            return content

        text = content.strip()

        fenced_match = re.match(
            r"^```(?:json)?\s*([\s\S]*?)\s*```$",
            text,
            flags=re.IGNORECASE,
        )
        if fenced_match:
            text = fenced_match.group(1).strip()
        else:
            fenced_search = re.search(
                r"```(?:json)?\s*([\s\S]*?)\s*```",
                text,
                flags=re.IGNORECASE,
            )
            if fenced_search:
                text = fenced_search.group(1).strip()

        first_brace = text.find("{")
        last_brace = text.rfind("}")
        if first_brace != -1 and last_brace != -1 and last_brace > first_brace:
            return text[first_brace:last_brace + 1].strip()

        return text

    def generate_slidespec(
        self,
        input_id: str,
        template_id: str,
        prepared: DataPrepResult,
        use_mock: bool = False,
    ) -> SlideSpec:
        """Generate slidespec using OpenAI API or fallback to deterministic/mock."""

        logger.info(f"🎯 Generating slidespec for {input_id}/{template_id}")
        logger.info(f"   Use mock: {use_mock}, LLM enabled: {config.settings.enable_llm}")

        # Use LLM if enabled and not forced to use mock
        if config.settings.enable_llm and not use_mock:
            try:
                logger.info(f"🤖 Attempting OpenAI generation for {input_id}/{template_id}")
                prompt = self._build_generation_prompt(template_id, prepared)
                response = self._call_openai_with_retry(prompt)
                slidespec = self._parse_llm_response(response, template_id)
                logger.info(f"✅ OpenAI generation completed successfully!")
                return slidespec

            except LLMGenerationError as e:
                logger.error(f"❌ LLM generation failed: {e}")
                # Fall through to deterministic generation
                logger.warning("⚠️  Falling back to deterministic generation")
            except Exception as e:
                logger.error(f"❌ Unexpected error in LLM generation: {e}")
                logger.exception("Full traceback:")
                logger.warning("⚠️  Falling back to deterministic generation")

        # Fallback 1: Try to load mock slidespec
        if use_mock:
            try:
                logger.info(f"📦 Attempting to load mock slidespec...")
                return self._load_mock_slidespec(input_id, template_id)
            except MockOutputNotFound:
                logger.warning(f"⚠️  Mock slidespec not found, using deterministic generation")

        # Fallback 2: Deterministic content generation
        logger.info(f"🔧 Using deterministic content generation for {input_id}/{template_id}")
        slides = []
        template = self.template_repo.get_descriptor(template_id)
        for slide in template.slides:
            data = prepared.slide_inputs.get(slide.slide_key, {})
            slides.append(
                SlideSpecItem(
                    slide_no=slide.slide_no,
                    slide_key=slide.slide_key,
                    data=data,
                )
            )
        logger.info(f"✅ Deterministic generation completed with {len(slides)} slides")
        return SlideSpec(template_id=template_id, slides=slides)

    def rewrite_slide(
        self,
        slide_spec: SlideSpec,
        slide_key: str,
        new_content: Dict[str, Any],
    ) -> SlideSpec:
        """Update a slide's data with new content (with optional LLM enhancement)."""

        # For now, simple merge of new content
        # TODO: Could enhance with LLM to refine the content based on context
        updated_slides = []
        for slide in slide_spec.slides:
            if slide.slide_key == slide_key:
                updated_slides.append(
                    SlideSpecItem(
                        slide_no=slide.slide_no,
                        slide_key=slide.slide_key,
                        data={**slide.data, **new_content},
                    )
                )
            else:
                updated_slides.append(slide)

        return SlideSpec(template_id=slide_spec.template_id, slides=updated_slides)
