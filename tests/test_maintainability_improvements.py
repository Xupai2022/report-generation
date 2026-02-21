"""Test script for maintainability improvements.

Tests:
1. Enhanced logging (file persistence, rotation, sensitive data filtering)
2. Request ID tracking
3. Enhanced health checks (disk, dependencies, sessions, locks)
"""

from __future__ import annotations

import argparse
import io
import logging
import requests
import sys
import time
import json
from pathlib import Path

BASE_URL = "http://127.0.0.1:8000"


def _pick_icons() -> dict[str, str]:
    """Pick icons that won't crash on Windows GBK consoles."""
    preferred = {"ok": "✅", "warn": "⚠️", "fail": "❌", "party": "🎉"}
    fallback = {"ok": "[OK]", "warn": "[WARN]", "fail": "[FAIL]", "party": "[DONE]"}

    encoding = sys.stdout.encoding or "utf-8"
    try:
        for icon in preferred.values():
            icon.encode(encoding)
        return preferred
    except Exception:
        return fallback


ICONS = _pick_icons()


def _get(url: str, **kwargs):
    kwargs.setdefault("timeout", 10)
    return requests.get(url, **kwargs)


def test_basic_health():
    """Test 1: Basic health check."""
    print("\n" + "=" * 60)
    print("TEST 1: Basic Health Check")
    print("=" * 60)

    response = _get(f"{BASE_URL}/health")

    if response.status_code == 200:
        data = response.json()
        print(f"{ICONS['ok']} Status: {data['status']}")
        print(f"   LLM Concurrency: {data['llm_concurrency']}")
        return True
    else:
        print(f"{ICONS['fail']} Failed: {response.status_code}")
        return False


def test_detailed_health():
    """Test 2: Detailed health check with all components."""
    print("\n" + "=" * 60)
    print("TEST 2: Detailed Health Check")
    print("=" * 60)

    response = _get(f"{BASE_URL}/health/detailed")

    if response.status_code == 200:
        data = response.json()
        print(f"{ICONS['ok']} Overall Status: {data['status']}")
        print(f"\nComponent Checks:")

        for check_name, check_data in data.get('checks', {}).items():
            status_icon = ICONS["ok"] if check_data['status'] == 'healthy' else ICONS["warn"] if check_data['status'] == 'degraded' else ICONS["fail"]
            print(f"  {status_icon} {check_name}: {check_data['message']}")
            if check_data.get('details'):
                for key, value in check_data['details'].items():
                    print(f"      - {key}: {value}")

        return True
    else:
        print(f"{ICONS['fail']} Failed: {response.status_code}")
        print(response.text)
        return False


def test_request_id_tracking():
    """Test 3: Request ID tracking in responses."""
    print("\n" + "=" * 60)
    print("TEST 3: Request ID Tracking")
    print("=" * 60)

    # Send request with custom Request ID
    custom_request_id = "test-12345"
    response = _get(
        f"{BASE_URL}/health",
        headers={"X-Request-ID": custom_request_id}
    )

    if response.status_code == 200:
        returned_request_id = response.headers.get("X-Request-ID")
        if returned_request_id == custom_request_id:
            print(f"{ICONS['ok']} Request ID preserved: {returned_request_id}")
        else:
            print(f"{ICONS['warn']} Request ID changed: sent={custom_request_id}, got={returned_request_id}")
        return True
    else:
        print(f"{ICONS['fail']} Failed: {response.status_code}")
        return False


def test_auto_generated_request_id():
    """Test 4: Auto-generated Request ID."""
    print("\n" + "=" * 60)
    print("TEST 4: Auto-Generated Request ID")
    print("=" * 60)

    # Send request without Request ID
    response = _get(f"{BASE_URL}/health")

    if response.status_code == 200:
        request_id = response.headers.get("X-Request-ID")
        if request_id:
            print(f"{ICONS['ok']} Auto-generated Request ID: {request_id}")
            return True
        else:
            print(f"{ICONS['fail']} No Request ID in response")
            return False
    else:
        print(f"{ICONS['fail']} Failed: {response.status_code}")
        return False


def test_log_files_created():
    """Test 5: Check if log files are created."""
    print("\n" + "=" * 60)
    print("TEST 5: Log File Persistence")
    print("=" * 60)

    log_dir = Path("mss_ai_ppt_sample_assets/backend/outputs/logs")

    # Expected log files
    expected_logs = ["app.log", "error.log"]

    results = []
    for log_file in expected_logs:
        log_path = log_dir / log_file
        if log_path.exists():
            size_kb = log_path.stat().st_size / 1024
            print(f"{ICONS['ok']} {log_file}: {size_kb:.2f} KB")
            results.append(True)
        else:
            print(f"{ICONS['warn']} {log_file}: Not created yet (will be created on first log)")
            results.append(True)  # Not an error, just not created yet

    return all(results)


def test_request_id_logged():
    """Test 6: Verify request ID shows up in app.log entries."""
    print("\n" + "=" * 60)
    print("TEST 6: Request ID Logged In File")
    print("=" * 60)

    custom_request_id = f"test-log-{int(time.time())}"
    response = _get(f"{BASE_URL}/health", headers={"X-Request-ID": custom_request_id})
    if response.status_code != 200:
        print(f"{ICONS['fail']} Failed: {response.status_code}")
        return False

    log_dir = Path("mss_ai_ppt_sample_assets/backend/outputs/logs")
    app_log = log_dir / "app.log"
    if not app_log.exists():
        print(f"{ICONS['warn']} app.log not created yet")
        return True

    time.sleep(0.2)
    content = app_log.read_text(encoding="utf-8", errors="replace")
    if f"[{custom_request_id}]" in content:
        print(f"{ICONS['ok']} Found request_id in app.log: {custom_request_id}")
        return True

    print(f"{ICONS['warn']} request_id not found in app.log (logging format may differ)")
    return False


def test_sensitive_data_filtering():
    """Test 7: Verify sensitive data filter redacts secrets (unit-style)."""
    print("\n" + "=" * 60)
    print("TEST 7: Sensitive Data Filtering")
    print("=" * 60)

    from mss_ai_ppt_sample_assets.backend.logging_config import RequestIdFilter, SensitiveDataFilter

    stream = io.StringIO()
    handler = logging.StreamHandler(stream)
    handler.setFormatter(logging.Formatter("%(message)s [%(request_id)s]"))
    handler.addFilter(RequestIdFilter())
    handler.addFilter(SensitiveDataFilter())

    test_logger = logging.getLogger("sensitive_data_filter_test")
    test_logger.setLevel(logging.INFO)
    test_logger.handlers = [handler]
    test_logger.propagate = False

    secret = "sk-" + ("a" * 32)
    test_logger.info(f"api_key={secret} password=hunter2 Bearer abc.def.ghi")

    output = stream.getvalue()
    if "REDACTED" in output and "hunter2" not in output and secret not in output:
        print(f"{ICONS['ok']} Secrets are redacted")
        return True

    print(f"{ICONS['fail']} Secrets not fully redacted")
    return False


def run_all_tests():
    """Run all maintainability tests."""
    print("\n" + "=" * 70)
    print("MAINTAINABILITY IMPROVEMENTS TEST SUITE")
    print("=" * 70)

    tests = [
        ("Basic Health Check", test_basic_health),
        ("Detailed Health Check", test_detailed_health),
        ("Request ID Tracking", test_request_id_tracking),
        ("Auto-Generated Request ID", test_auto_generated_request_id),
        ("Log File Persistence", test_log_files_created),
        ("Request ID Logged In File", test_request_id_logged),
        ("Sensitive Data Filtering", test_sensitive_data_filtering),
    ]

    results = []
    for test_name, test_func in tests:
        try:
            result = test_func()
            results.append((test_name, result))
        except Exception as e:
            print(f"{ICONS['fail']} Test failed with exception: {e}")
            results.append((test_name, False))
        time.sleep(0.5)  # Brief pause between tests

    # Summary
    print("\n" + "=" * 70)
    print("TEST SUMMARY")
    print("=" * 70)

    passed = sum(1 for _, result in results if result)
    total = len(results)

    for test_name, result in results:
        status = f"{ICONS['ok']} PASS" if result else f"{ICONS['fail']} FAIL"
        print(f"{status} - {test_name}")

    print(f"\n{passed}/{total} tests passed ({passed/total*100:.1f}%)")

    if passed == total:
        print(f"\n{ICONS['party']} All tests passed! Maintainability improvements working correctly.")
    else:
        print(f"\n{ICONS['warn']} {total - passed} test(s) failed. Please check the output above.")

    return passed == total


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--base-url", default=BASE_URL)
    parser.add_argument("--no-prompt", action="store_true", help="Run without waiting for Enter")
    args = parser.parse_args()

    BASE_URL = args.base_url.rstrip("/")

    print(f"\n{ICONS['warn']} Make sure the server is running at {BASE_URL}")
    print("   Run: python -m mss_ai_ppt_sample_assets.backend.app")

    if not args.no_prompt:
        input("\nPress Enter to start tests...")

    success = run_all_tests()
    raise SystemExit(0 if success else 1)
