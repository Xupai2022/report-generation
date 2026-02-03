"""Retry policies for external API calls.

This module provides retry decorators for handling transient failures
in external API calls, particularly OpenAI LLM requests.
"""

from tenacity import (
    retry,
    stop_after_attempt,
    wait_exponential,
    retry_if_exception_type,
    before_sleep_log
)
from openai import RateLimitError, APIConnectionError, APITimeoutError
import logging

logger = logging.getLogger(__name__)


# LLM retry policy configuration
llm_retry_policy = retry(
    # Stop after 3 attempts
    stop=stop_after_attempt(3),

    # Exponential backoff: 2s, 4s, 8s, ... up to 60s
    wait=wait_exponential(multiplier=1, min=2, max=60),

    # Only retry on specific exceptions
    retry=retry_if_exception_type((
        RateLimitError,      # Rate limit hit, wait and retry
        APIConnectionError,  # Network/connection issues
        APITimeoutError      # Request timed out
    )),

    # Log before sleeping
    before_sleep=before_sleep_log(logger, logging.WARNING)
)


def with_llm_retry(func):
    """Decorator to add retry logic to LLM API calls.

    Automatically retries on rate limits, connection errors, and timeouts.
    Uses exponential backoff with a maximum of 3 attempts.

    Example:
        @with_llm_retry
        def call_openai(prompt: str):
            response = client.chat.completions.create(...)
            return response

    Args:
        func: Function to wrap with retry logic

    Returns:
        Wrapped function with retry logic
    """
    return llm_retry_policy(func)
