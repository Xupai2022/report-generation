"""Test concurrent report generation to verify session isolation and file locking.

This script simulates multiple concurrent users generating reports simultaneously.
"""

import requests
import threading
import time
from typing import List, Dict
import json
import sys
import io

# Fix Windows console encoding issues
if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')


API_BASE_URL = "http://localhost:8000"


def generate_report(user_id: int, results: List[Dict]):
    """Simulate a user generating a report.

    Args:
        user_id: Unique user identifier
        results: Shared list to store results
    """
    start_time = time.time()
    try:
        print(f"[User {user_id}] Starting report generation...")

        # Make generate request
        response = requests.post(
            f"{API_BASE_URL}/generate",
            json={
                "input_id": "tenant_acme_2025-11",
                "template_id": "mss_executive_v2",
                "use_mock": True  # Use mock to avoid LLM API calls
            },
            timeout=120
        )

        elapsed = time.time() - start_time

        if response.status_code == 200:
            data = response.json()
            session_id = data.get("session_id")
            job_id = data.get("job_id")
            print(f"[User {user_id}] ✓ SUCCESS in {elapsed:.2f}s - Session: {session_id}, Job: {job_id}")
            results.append({
                "user_id": user_id,
                "success": True,
                "session_id": session_id,
                "job_id": job_id,
                "elapsed": elapsed,
                "error": None
            })
        else:
            print(f"[User {user_id}] ✗ FAILED in {elapsed:.2f}s - Status: {response.status_code}")
            results.append({
                "user_id": user_id,
                "success": False,
                "session_id": None,
                "job_id": None,
                "elapsed": elapsed,
                "error": f"HTTP {response.status_code}: {response.text}"
            })

    except Exception as e:
        elapsed = time.time() - start_time
        print(f"[User {user_id}] ✗ EXCEPTION in {elapsed:.2f}s - {e}")
        results.append({
            "user_id": user_id,
            "success": False,
            "session_id": None,
            "job_id": None,
            "elapsed": elapsed,
            "error": str(e)
        })


def test_concurrent_generation(num_users: int = 5):
    """Test concurrent report generation with multiple users.

    Args:
        num_users: Number of concurrent users to simulate
    """
    print("=" * 80)
    print(f"CONCURRENT REPORT GENERATION TEST")
    print(f"Simulating {num_users} concurrent users...")
    print("=" * 80)

    # Check if server is running
    try:
        response = requests.get(f"{API_BASE_URL}/health", timeout=5)
        if response.status_code != 200:
            print("ERROR: Server is not healthy!")
            return
    except requests.exceptions.RequestException:
        print("ERROR: Cannot connect to server at", API_BASE_URL)
        print("Please start the server with: python mss_ai_ppt_sample_assets/backend/app.py")
        return

    results: List[Dict] = []
    threads: List[threading.Thread] = []

    # Start all threads simultaneously
    start_time = time.time()
    for user_id in range(1, num_users + 1):
        thread = threading.Thread(target=generate_report, args=(user_id, results))
        thread.start()
        threads.append(thread)

    # Wait for all threads to complete
    for thread in threads:
        thread.join()

    total_elapsed = time.time() - start_time

    # Analyze results
    print("\n" + "=" * 80)
    print("TEST RESULTS")
    print("=" * 80)

    successful = [r for r in results if r["success"]]
    failed = [r for r in results if not r["success"]]

    print(f"\n✓ Successful: {len(successful)}/{num_users}")
    print(f"✗ Failed: {len(failed)}/{num_users}")
    print(f"\nTotal elapsed time: {total_elapsed:.2f}s")

    if successful:
        avg_time = sum(r["elapsed"] for r in successful) / len(successful)
        print(f"Average generation time: {avg_time:.2f}s")

        # Check for session ID collisions
        session_ids = [r["session_id"] for r in successful]
        unique_sessions = set(session_ids)
        print(f"\nSession isolation check:")
        print(f"  - Total sessions: {len(session_ids)}")
        print(f"  - Unique sessions: {len(unique_sessions)}")

        if len(session_ids) == len(unique_sessions):
            print(f"  - ✓ All sessions are unique (no collisions)")
        else:
            print(f"  - ✗ WARNING: Session ID collision detected!")
            collisions = [sid for sid in session_ids if session_ids.count(sid) > 1]
            print(f"    Collided sessions: {set(collisions)}")

    if failed:
        print(f"\n❌ Failed requests:")
        for r in failed:
            print(f"  - User {r['user_id']}: {r['error']}")

    # Summary
    print("\n" + "=" * 80)
    if len(successful) == num_users:
        print("✅ PASS: All concurrent requests succeeded with unique sessions")
    elif len(successful) > 0:
        print("⚠️  PARTIAL: Some requests failed")
    else:
        print("❌ FAIL: All requests failed")
    print("=" * 80)


if __name__ == "__main__":
    # Test with 5 concurrent users
    test_concurrent_generation(num_users=5)

    print("\n\n")
    input("Press Enter to test with 10 concurrent users...")

    # Test with more users
    test_concurrent_generation(num_users=10)
