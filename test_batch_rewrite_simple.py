# -*- coding: utf-8 -*-
"""Simple test for batch rewrite API"""
import requests
import json
import sys

# Force UTF-8 output
if sys.platform == 'win32':
    import codecs
    sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'strict')

BASE_URL = "http://127.0.0.1:8000"

def test_batch_rewrite():
    """Test batch rewrite functionality"""
    print("=" * 60)
    print("Test: Batch Rewrite API")
    print("=" * 60)

    # Step 1: Generate a report
    print("\n1. Generating test report...")
    gen_response = requests.post(
        f"{BASE_URL}/generate",
        json={
            "input_id": "tenant_contoso_2025-01",
            "template_id": "mss_executive_v2",
            "use_mock": True
        }
    )

    if gen_response.status_code != 200:
        print(f"[FAIL] Generation failed: {gen_response.text}")
        return False

    gen_result = gen_response.json()
    job_id = gen_result["job_id"]
    print(f"[OK] Report generated: {job_id}")

    # Step 2: Rewrite multiple slides
    print(f"\n2. Rewriting 3 slides in batch...")
    rewrite_response = requests.post(
        f"{BASE_URL}/rewrite",
        json={
            "job_id": job_id,
            "slides": [
                {
                    "slide_key": "cover",
                    "new_content": {
                        "TITLE": "[BATCH TEST] Modified Title 1",
                        "SUBTITLE": "Batch test subtitle"
                    }
                },
                {
                    "slide_key": "summary",
                    "new_content": {
                        "HEADLINE": "[BATCH TEST] Modified Headline 2"
                    }
                },
                {
                    "slide_key": "threats",
                    "new_content": {
                        "HEADLINE": "[BATCH TEST] Modified Headline 3"
                    }
                }
            ]
        }
    )

    if rewrite_response.status_code != 200:
        print(f"[FAIL] Batch rewrite failed:")
        print(f"  Status: {rewrite_response.status_code}")
        print(f"  Response: {rewrite_response.text}")
        return False

    rewrite_result = rewrite_response.json()
    print(f"[OK] Batch rewrite successful!")
    print(f"  Updated slides: {rewrite_result.get('updated_slides')}")
    print(f"  Updated count: {rewrite_result.get('updated_count')}")
    print(f"  Warnings: {rewrite_result.get('warnings')}")

    # Step 3: Test single slide mode (legacy)
    print(f"\n3. Testing single slide mode (legacy)...")
    single_response = requests.post(
        f"{BASE_URL}/rewrite",
        json={
            "job_id": job_id,
            "slide_key": "cover",
            "new_content": {
                "TITLE": "[SINGLE TEST] Modified via legacy mode"
            }
        }
    )

    if single_response.status_code != 200:
        print(f"[FAIL] Single rewrite failed: {single_response.text}")
        return False

    single_result = single_response.json()
    print(f"[OK] Single rewrite successful!")
    print(f"  Updated slides: {single_result.get('updated_slides')}")
    print(f"  Updated count: {single_result.get('updated_count')}")

    return True


if __name__ == "__main__":
    print("\n[TEST] Testing Batch Rewrite API\n")

    try:
        # Check server health
        health = requests.get(f"{BASE_URL}/health")
        if health.status_code != 200:
            print("[FAIL] Server health check failed")
            sys.exit(1)
        print("[OK] Server is running\n")

        # Run test
        success = test_batch_rewrite()

        print("\n" + "=" * 60)
        if success:
            print("[SUCCESS] All tests passed!")
        else:
            print("[FAILED] Some tests failed")
        print("=" * 60)

    except requests.exceptions.ConnectionError:
        print("[FAIL] Cannot connect to server on port 8000")
        sys.exit(1)
    except Exception as e:
        print(f"[ERROR] Unexpected error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
