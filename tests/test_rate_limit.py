"""Test rate limiting and frontend debounce functionality."""

import requests
import time

BASE_URL = "http://127.0.0.1:8000"

def test_rate_limiting():
    """Test backend IP-based rate limiting."""
    print("=" * 60)
    print("测试后端 IP 速率限制")
    print("=" * 60)

    # Test data
    payload = {
        "input_id": "tenant_acme_2025-11",
        "template_id": "mss_executive_v2",
        "use_mock": True,
        "session_id": f"test_{int(time.time())}"
    }

    print("\n1. 第一次请求（应该成功）...")
    try:
        response1 = requests.post(f"{BASE_URL}/api/v1/reports", json=payload)
        print(f"   状态码: {response1.status_code}")
        if response1.status_code == 202:
            print(f"   ✅ 成功: {response1.json()}")
        else:
            print(f"   结果: {response1.text}")
    except Exception as e:
        print(f"   ❌ 错误: {e}")

    print("\n2. 立即第二次请求（应该被限流，返回 429）...")
    try:
        response2 = requests.post(f"{BASE_URL}/api/v1/reports", json=payload)
        print(f"   状态码: {response2.status_code}")
        if response2.status_code == 429:
            print(f"   ✅ 成功被限流: {response2.json()}")
        else:
            print(f"   ⚠️  未被限流: {response2.text}")
    except Exception as e:
        print(f"   错误: {e}")

    print("\n3. 等待 11 秒后第三次请求（应该成功）...")
    print("   等待中", end="", flush=True)
    for i in range(11):
        time.sleep(1)
        print(".", end="", flush=True)
    print(" 完成")

    try:
        response3 = requests.post(f"{BASE_URL}/api/v1/reports", json=payload)
        print(f"   状态码: {response3.status_code}")
        if response3.status_code == 202:
            print(f"   ✅ 成功: {response3.json()}")
        else:
            print(f"   结果: {response3.text}")
    except Exception as e:
        print(f"   ❌ 错误: {e}")


def test_api_accessibility():
    """Test if API endpoints are accessible."""
    print("\n" + "=" * 60)
    print("测试 API 端点可访问性")
    print("=" * 60)

    endpoints = [
        "/api/v1/templates",
        "/api/v1/inputs",
        "/api/v1/system/logs?limit=5"
    ]

    for endpoint in endpoints:
        try:
            response = requests.get(f"{BASE_URL}{endpoint}")
            print(f"\n{endpoint}")
            print(f"  状态码: {response.status_code}")
            if response.status_code == 200:
                data = response.json()
                if 'data' in data:
                    print(f"  ✅ 数据项数: {len(data['data']) if isinstance(data['data'], list) else '1'}")
                else:
                    print(f"  ✅ 响应: {str(data)[:100]}...")
            else:
                print(f"  ❌ 错误: {response.text[:100]}")
        except Exception as e:
            print(f"  ❌ 异常: {e}")


def main():
    print("\n")
    print("╔" + "=" * 58 + "╗")
    print("║" + " " * 10 + "前端防抖 + 后端速率限制测试" + " " * 14 + "║")
    print("╚" + "=" * 58 + "╝")

    # First test API accessibility
    test_api_accessibility()

    # Then test rate limiting
    test_rate_limiting()

    print("\n" + "=" * 60)
    print("测试完成")
    print("=" * 60)
    print("\n前端防抖功能说明:")
    print("  - 点击'生成报告'按钮后，按钮将被禁用 30 秒")
    print("  - 按钮显示倒计时: '生成报告 (29秒后可重试)'")
    print("  - 如果生成失败，倒计时会被清除，允许立即重试")
    print("\n后端速率限制说明:")
    print("  - 同一 IP 10 秒内只能提交一次生成请求")
    print("  - 超出限制返回 HTTP 429 错误")
    print("  - 错误消息包含剩余等待时间")
    print("\n请打开浏览器访问 http://127.0.0.1:8000/ui/ 测试前端防抖功能")
    print()


if __name__ == "__main__":
    main()
