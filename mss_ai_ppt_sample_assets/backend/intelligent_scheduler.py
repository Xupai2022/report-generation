"""Intelligent request scheduler with dynamic concurrency control.

Adjusts concurrent LLM requests based on:
- Request size (number of slides)
- System load
- Historical performance
- API rate limits
"""

from __future__ import annotations

import asyncio
import time
from typing import Dict, Optional
from dataclasses import dataclass
from collections import deque
import logging

logger = logging.getLogger(__name__)


@dataclass
class RequestMetrics:
    """Metrics for a completed request."""
    template_id: str
    num_slides: int
    duration: float
    timestamp: float
    success: bool


class IntelligentScheduler:
    """Dynamically adjusts concurrency based on request size and system load.

    Features:
    - Small requests (1-3 slides): Use more concurrent slots
    - Large requests (8+ slides): Use fewer slots to avoid timeouts
    - Adapts based on success rate and response times
    """

    def __init__(
        self,
        min_concurrency: int = 2,
        max_concurrency: int = 5,
        default_concurrency: int = 3
    ):
        """Initialize scheduler.

        Args:
            min_concurrency: Minimum concurrent requests
            max_concurrency: Maximum concurrent requests
            default_concurrency: Starting concurrency level
        """
        self.min_concurrency = min_concurrency
        self.max_concurrency = max_concurrency
        self.current_concurrency = default_concurrency

        # Metrics tracking
        self.metrics_history: deque[RequestMetrics] = deque(maxlen=100)
        self.active_requests: Dict[str, float] = {}  # session_id -> start_time

        # Semaphore for concurrency control
        self.semaphore = asyncio.Semaphore(default_concurrency)
        self._semaphore_lock = asyncio.Lock()

    def estimate_request_weight(self, template_id: str, num_slides: Optional[int] = None) -> int:
        """Estimate request weight based on template and size.

        Args:
            template_id: Template identifier
            num_slides: Number of slides (if known)

        Returns:
            Weight value (1-3, where 3 = heaviest)
        """
        # Use historical data if available
        if num_slides is None:
            num_slides = self._get_avg_slides_for_template(template_id)

        # Weight calculation
        if num_slides <= 3:
            return 1  # Light request
        elif num_slides <= 6:
            return 2  # Medium request
        else:
            return 3  # Heavy request

    def _get_avg_slides_for_template(self, template_id: str) -> int:
        """Get average slide count for a template from history."""
        relevant_metrics = [
            m.num_slides for m in self.metrics_history
            if m.template_id == template_id
        ]
        if relevant_metrics:
            return int(sum(relevant_metrics) / len(relevant_metrics))
        return 5  # Default assumption

    async def adjust_concurrency(self):
        """Dynamically adjust concurrency based on current metrics.

        Called periodically or after each request completion.
        """
        if len(self.metrics_history) < 10:
            return  # Not enough data yet

        # Calculate recent success rate
        recent = list(self.metrics_history)[-20:]
        success_rate = sum(1 for m in recent if m.success) / len(recent)

        # Calculate average response time
        avg_duration = sum(m.duration for m in recent) / len(recent)

        # Decision logic
        new_concurrency = self.current_concurrency

        if success_rate < 0.8:
            # Too many failures, reduce concurrency
            new_concurrency = max(self.min_concurrency, self.current_concurrency - 1)
            logger.info(f"⬇️ Reducing concurrency due to low success rate: {success_rate:.1%}")

        elif success_rate > 0.95 and avg_duration < 180:
            # High success and fast responses, increase concurrency
            new_concurrency = min(self.max_concurrency, self.current_concurrency + 1)
            logger.info(f"⬆️ Increasing concurrency: success={success_rate:.1%}, avg_time={avg_duration:.0f}s")

        elif avg_duration > 300:
            # Slow responses, reduce concurrency
            new_concurrency = max(self.min_concurrency, self.current_concurrency - 1)
            logger.info(f"⬇️ Reducing concurrency due to slow responses: {avg_duration:.0f}s")

        if new_concurrency != self.current_concurrency:
            await self._update_semaphore(new_concurrency)

    async def _update_semaphore(self, new_value: int):
        """Update semaphore to new concurrency value.

        Args:
            new_value: New concurrency limit
        """
        async with self._semaphore_lock:
            old_value = self.current_concurrency
            self.current_concurrency = new_value

            # Create new semaphore with updated value
            self.semaphore = asyncio.Semaphore(new_value)

            logger.info(f"🔄 Concurrency updated: {old_value} -> {new_value}")

    async def acquire(self, session_id: str, weight: int = 1) -> None:
        """Acquire a slot for request execution.

        Args:
            session_id: Session identifier
            weight: Request weight (1-3)
        """
        # For now, simple approach: heavy requests wait for more available slots
        if weight >= 3:
            # Heavy request: wait for at least 2 free slots
            while self.semaphore._value < 2:
                await asyncio.sleep(0.1)

        await self.semaphore.acquire()
        self.active_requests[session_id] = time.time()

        logger.info(
            f"🟢 Acquired slot: session={session_id}, weight={weight}, "
            f"active={len(self.active_requests)}/{self.current_concurrency}"
        )

    def release(self, session_id: str):
        """Release a slot after request completion.

        Args:
            session_id: Session identifier
        """
        self.semaphore.release()

        if session_id in self.active_requests:
            del self.active_requests[session_id]

        logger.info(
            f"🔴 Released slot: session={session_id}, "
            f"active={len(self.active_requests)}/{self.current_concurrency}"
        )

    def record_completion(
        self,
        session_id: str,
        template_id: str,
        num_slides: int,
        success: bool
    ):
        """Record request completion metrics.

        Args:
            session_id: Session identifier
            template_id: Template used
            num_slides: Number of slides generated
            success: Whether request succeeded
        """
        start_time = self.active_requests.get(session_id)
        if start_time is None:
            logger.warning(f"No start time found for session {session_id}")
            return

        duration = time.time() - start_time

        metrics = RequestMetrics(
            template_id=template_id,
            num_slides=num_slides,
            duration=duration,
            timestamp=time.time(),
            success=success
        )

        self.metrics_history.append(metrics)

        logger.info(
            f"📊 Recorded metrics: session={session_id}, "
            f"duration={duration:.1f}s, success={success}"
        )

    def get_stats(self) -> dict:
        """Get current scheduler statistics.

        Returns:
            Dictionary with scheduler stats
        """
        if not self.metrics_history:
            return {
                "current_concurrency": self.current_concurrency,
                "active_requests": len(self.active_requests),
                "total_completed": 0
            }

        recent = list(self.metrics_history)[-20:]
        success_rate = sum(1 for m in recent if m.success) / len(recent) if recent else 0
        avg_duration = sum(m.duration for m in recent) / len(recent) if recent else 0

        return {
            "current_concurrency": self.current_concurrency,
            "min_concurrency": self.min_concurrency,
            "max_concurrency": self.max_concurrency,
            "active_requests": len(self.active_requests),
            "total_completed": len(self.metrics_history),
            "recent_success_rate": round(success_rate, 3),
            "avg_response_time": round(avg_duration, 1)
        }


# Example usage in app.py
"""
from intelligent_scheduler import IntelligentScheduler

# Initialize scheduler
scheduler = IntelligentScheduler(min_concurrency=2, max_concurrency=5)

@app.post("/generate")
async def generate(req: GenerateRequest):
    session_id = req.session_id or generate_session_id()

    # Get request weight
    num_slides = get_num_slides(req.template_id)
    weight = scheduler.estimate_request_weight(req.template_id, num_slides)

    # Acquire slot with weight consideration
    await scheduler.acquire(session_id, weight)

    try:
        # Generate report
        result = await generate_report(...)

        # Record success
        scheduler.record_completion(session_id, req.template_id, num_slides, True)

        return result
    except Exception as e:
        # Record failure
        scheduler.record_completion(session_id, req.template_id, num_slides, False)
        raise
    finally:
        # Release slot
        scheduler.release(session_id)

        # Adjust concurrency if needed
        await scheduler.adjust_concurrency()

@app.get("/scheduler/stats")
def get_scheduler_stats():
    return scheduler.get_stats()
"""
