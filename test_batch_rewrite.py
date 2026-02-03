"""Test script for batch rewrite functionality."""
import requests

BASE_URL = "http://127.0.0.1:8000"

def test_single_slide_rewrite():
    """Test single slide rewrite (legacy mode)."""
    print("=" * 60)
    print("Test 1: Single Slide Rewrite (Legacy Mode)")
    print("=" * 60)

    # First, generate a report
    print("\n1. Generating a test report...")
    gen_response = requests.post(
        f"{BASE_URL}/generate",
        json={
            "input_id": "tenant_contoso_2025-01",
            "template_id": "mss_executive_v2",
            "use_mock": True
        }
    )

    if gen_response.status_code != 200:
        print(f"❌ Generation failed: {gen_response.text}")
        return

    gen_result = gen_response.json()
    job_id = gen_result["job_id"]
    print(f"✅ Report generated: {job_id}")

    # Rewrite a single slide
    print("\n2. Rewriting single slide 'cover'...")
    rewrite_response = requests.post(
        f"{BASE_URL}/rewrite",
        json={
            "job_id": job_id,
            "slide_key": "cover",
            "new_content": {
                "TITLE": "【测试】修改后的标题",
                "SUBTITLE": "单幻灯片改写测试"
            }
        }
    )

    if rewrite_response.status_code != 200:
        print(f"❌ Rewrite failed: {rewrite_response.text}")
        return

    rewrite_result = rewrite_response.json()
    print(f"✅ Rewrite successful!")
    print(f"   Updated slides: {rewrite_result.get('updated_slides')}")
    print(f"   Updated count: {rewrite_result.get('updated_count')}")
    print(f"   Warnings: {rewrite_result.get('warnings')}")


def test_batch_rewrite():
    """Test batch rewrite (new mode)."""
    print("\n" + "=" * 60)
    print("Test 2: Batch Rewrite (New Mode)")
    print("=" * 60)

    # First, generate a report
    print("\n1. Generating a test report...")
    gen_response = requests.post(
        f"{BASE_URL}/generate",
        json={
            "input_id": "tenant_contoso_2025-01",
            "template_id": "mss_executive_v2",
            "use_mock": True
        }
    )

    if gen_response.status_code != 200:
        print(f"❌ Generation failed: {gen_response.text}")
        return

    gen_result = gen_response.json()
    job_id = gen_result["job_id"]
    print(f"✅ Report generated: {job_id}")

    # Rewrite multiple slides
    print("\n2. Rewriting 3 slides in one request...")
    rewrite_response = requests.post(
        f"{BASE_URL}/rewrite",
        json={
            "job_id": job_id,
            "slides": [
                {
                    "slide_key": "cover",
                    "new_content": {
                        "TITLE": "【批量测试】封面标题已修改",
                        "SUBTITLE": "多幻灯片批量改写测试"
                    }
                },
                {
                    "slide_key": "summary",
                    "new_content": {
                        "HEADLINE": "【批量测试】概述标题已修改",
                        "CONTENT": "这是批量修改的概述内容，一次请求修改多个幻灯片。"
                    }
                },
                {
                    "slide_key": "threats",
                    "new_content": {
                        "HEADLINE": "【批量测试】威胁分析标题已修改",
                        "KEY_FINDINGS": "批量修改的威胁发现内容。"
                    }
                }
            ]
        }
    )

    if rewrite_response.status_code != 200:
        print(f"❌ Rewrite failed: {rewrite_response.text}")
        return

    rewrite_result = rewrite_response.json()
    print(f"✅ Batch rewrite successful!")
    print(f"   Updated slides: {rewrite_result.get('updated_slides')}")
    print(f"   Updated count: {rewrite_result.get('updated_count')}")
    print(f"   Warnings: {rewrite_result.get('warnings')}")


def test_invalid_requests():
    """Test error handling."""
    print("\n" + "=" * 60)
    print("Test 3: Error Handling")
    print("=" * 60)

    # Generate a report first
    gen_response = requests.post(
        f"{BASE_URL}/generate",
        json={
            "input_id": "tenant_contoso_2025-01",
            "template_id": "mss_executive_v2",
            "use_mock": True
        }
    )
    job_id = gen_response.json()["job_id"]

    # Test 3a: No mode provided
    print("\n3a. Testing request with no mode...")
    response = requests.post(
        f"{BASE_URL}/rewrite",
        json={"job_id": job_id}
    )
    print(f"   Status: {response.status_code} (expected: 422)")
    if response.status_code == 422:
        print("   ✅ Correctly rejected request with no mode")

    # Test 3b: Both modes provided
    print("\n3b. Testing request with both modes...")
    response = requests.post(
        f"{BASE_URL}/rewrite",
        json={
            "job_id": job_id,
            "slide_key": "cover",
            "new_content": {"TITLE": "test"},
            "slides": [{"slide_key": "summary", "new_content": {}}]
        }
    )
    print(f"   Status: {response.status_code} (expected: 422)")
    if response.status_code == 422:
        print("   ✅ Correctly rejected request with both modes")

    # Test 3c: Non-existent slide
    print("\n3c. Testing rewrite of non-existent slide...")
    response = requests.post(
        f"{BASE_URL}/rewrite",
        json={
            "job_id": job_id,
            "slides": [
                {"slide_key": "cover", "new_content": {"TITLE": "Valid"}},
                {"slide_key": "nonexistent_slide", "new_content": {"TITLE": "Invalid"}}
            ]
        }
    )
    if response.status_code == 200:
        result = response.json()
        print(f"   ✅ Partial success with warnings:")
        print(f"      Updated: {result.get('updated_slides')}")
        print(f"      Not found: {result.get('not_found_slides')}")
        print(f"      Warnings: {result.get('warnings')}")


if __name__ == "__main__":
    print("\n[TEST] Testing Batch Rewrite Functionality\n")

    # Check if server is running
    try:
        health = requests.get(f"{BASE_URL}/health")
        if health.status_code == 200:
            print("✅ Server is running\n")
        else:
            print("❌ Server health check failed")
            exit(1)
    except requests.exceptions.ConnectionError:
        print("❌ Cannot connect to server. Is it running on port 8000?")
        exit(1)

    # Run tests
    test_single_slide_rewrite()
    test_batch_rewrite()
    test_invalid_requests()

    print("\n" + "=" * 60)
    print("✅ All tests completed!")
    print("=" * 60)
