"""Test script for OpenAI concurrency limiter.

This script simulates multiple concurrent requests to verify that the
semaphore correctly limits concurrent LLM calls.

Usage:
    python test_concurrency_limiter.py
"""

import asyncio
import aiohttp
import time
from typing import List, Dict
import json


BASE_URL = "http://localhost:8000"
MAX_CONCURRENT_REQUESTS = 3  # Should match the value in app.py


async def check_health() -> Dict:
    """Check server health and concurrency status."""
    async with aiohttp.ClientSession() as session:
        async with session.get(f"{BASE_URL}/health") as resp:
            return await resp.json()


async def generate_report(session: aiohttp.ClientSession, request_id: int, use_mock: bool = True) -> Dict:
    """Send a generate request and track timing."""
    start_time = time.time()

    payload = {
        "input_id": "tenant_acme_2025-11",
        "template_id": "mss_executive_v2",
        "use_mock": use_mock
    }

    print(f"[Request {request_id}] 🚀 Starting at {start_time:.2f}s")

    # Set timeout to 10 minutes (600 seconds)
    timeout = aiohttp.ClientTimeout(total=600)

    try:
        async with session.post(f"{BASE_URL}/generate", json=payload, timeout=timeout) as resp:
            end_time = time.time()
            duration = end_time - start_time

            if resp.status == 200:
                result = await resp.json()
                print(f"[Request {request_id}] ✅ Completed in {duration:.2f}s - Job: {result.get('job_id')}")
                return {
                    "request_id": request_id,
                    "status": "success",
                    "duration": duration,
                    "job_id": result.get("job_id")
                }
            else:
                error_text = await resp.text()
                print(f"[Request {request_id}] ❌ Failed ({resp.status}) in {duration:.2f}s - {error_text[:100]}")
                return {
                    "request_id": request_id,
                    "status": "error",
                    "duration": duration,
                    "error": error_text
                }
    except Exception as e:
        end_time = time.time()
        duration = end_time - start_time
        print(f"[Request {request_id}] ❌ Exception in {duration:.2f}s - {str(e)}")
        return {
            "request_id": request_id,
            "status": "exception",
            "duration": duration,
            "error": str(e)
        }


async def test_concurrent_requests(num_requests: int = 10, use_mock: bool = True):
    """Test concurrent request limiting.

    Args:
        num_requests: Number of concurrent requests to send
        use_mock: Whether to use mock mode (faster, no real LLM calls)
    """
    print(f"\n{'='*80}")
    print(f"🧪 Testing Concurrency Limiter")
    print(f"{'='*80}")
    print(f"Number of requests: {num_requests}")
    print(f"Max concurrent LLM requests: {MAX_CONCURRENT_REQUESTS}")
    print(f"Mode: {'MOCK' if use_mock else 'REAL LLM'}")
    print(f"{'='*80}\n")

    # Check initial health
    print("📊 Initial server health:")
    health = await check_health()
    print(json.dumps(health, indent=2))
    print()

    # Create session and send all requests concurrently
    async with aiohttp.ClientSession() as session:
        start_time = time.time()

        # Launch all requests at once
        tasks = [
            generate_report(session, i, use_mock=use_mock)
            for i in range(1, num_requests + 1)
        ]

        results = await asyncio.gather(*tasks)

        end_time = time.time()
        total_duration = end_time - start_time

    # Check final health
    print("\n📊 Final server health:")
    health = await check_health()
    print(json.dumps(health, indent=2))
    print()

    # Analyze results
    print(f"\n{'='*80}")
    print(f"📈 Results Summary")
    print(f"{'='*80}")
    print(f"Total time: {total_duration:.2f}s")
    print(f"Average time per request: {sum(r['duration'] for r in results) / len(results):.2f}s")
    print(f"Successful requests: {sum(1 for r in results if r['status'] == 'success')}/{num_requests}")
    print(f"Failed requests: {sum(1 for r in results if r['status'] != 'success')}/{num_requests}")

    # Expected behavior analysis
    print(f"\n{'='*80}")
    print(f"🔍 Concurrency Analysis")
    print(f"{'='*80}")

    if use_mock:
        print("⚠️  Mock mode: Requests bypass semaphore (fast execution)")
    else:
        print(f"✅ Real LLM mode: Max {MAX_CONCURRENT_REQUESTS} requests should run concurrently")
        print(f"   Expected batches: {(num_requests + MAX_CONCURRENT_REQUESTS - 1) // MAX_CONCURRENT_REQUESTS}")

        # Check if requests were properly queued
        sorted_results = sorted(results, key=lambda x: x['duration'])
        fastest = sorted_results[0]['duration']
        slowest = sorted_results[-1]['duration']

        print(f"   Fastest request: {fastest:.2f}s")
        print(f"   Slowest request: {slowest:.2f}s")

        if num_requests > MAX_CONCURRENT_REQUESTS:
            if slowest > fastest * 1.5:
                print("   ✅ Queueing detected: Later requests took longer (as expected)")
            else:
                print("   ⚠️  Unexpected: All requests finished in similar time")

    print(f"{'='*80}\n")


async def test_semaphore_blocking():
    """Test that semaphore correctly blocks requests when saturated."""
    print(f"\n{'='*80}")
    print(f"🧪 Testing Semaphore Blocking (Real LLM Mode)")
    print(f"{'='*80}")
    print(f"This test sends {MAX_CONCURRENT_REQUESTS + 2} requests to verify blocking\n")

    # Use real LLM mode to trigger semaphore
    await test_concurrent_requests(
        num_requests=MAX_CONCURRENT_REQUESTS + 2,
        use_mock=False  # Real LLM mode to trigger semaphore
    )


async def main():
    """Run all concurrency tests."""
    print("\n" + "🎯 " * 20)
    print("CONCURRENCY LIMITER TEST SUITE")
    print("🎯 " * 20 + "\n")

    # Test 1: Mock mode (baseline - no semaphore)
    print("\n📋 Test 1: Mock Mode (No Semaphore)")
    print("-" * 80)
    await test_concurrent_requests(num_requests=5, use_mock=True)

    # Test 2: Real LLM mode with moderate load
    print("\n📋 Test 2: Real LLM Mode - Moderate Load")
    print("-" * 80)
    await test_concurrent_requests(num_requests=5, use_mock=False)

    # Test 3: Real LLM mode with high load (exceeds limit)
    print("\n📋 Test 3: Real LLM Mode - High Load (Exceeds Limit)")
    print("-" * 80)
    await test_concurrent_requests(num_requests=10, use_mock=False)

    print("\n" + "✨ " * 20)
    print("ALL TESTS COMPLETED")
    print("✨ " * 20 + "\n")


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\n\n⚠️  Tests interrupted by user")
    except Exception as e:
        print(f"\n\n❌ Test suite failed: {e}")
        import traceback
        traceback.print_exc()
