"""简单测试后端IP限制"""
import requests
import time
import sys

BASE_URL = "http://127.0.0.1:8000"

def test():
    payload = {
        "input_id": "tenant_acme_2025-11",
        "template_id": "mss_executive_v2",
        "use_mock": True
    }

    print("=" * 60)
    print("测试后端IP速率限制 (10秒/次)")
    print("=" * 60)

    # 第一次请求
    print("\n[1] 第一次请求...")
    r1 = requests.post(f"{BASE_URL}/api/v1/reports", json=payload)
    print(f"    状态码: {r1.status_code}")
    if r1.status_code == 202:
        print(f"    结果: 成功接受请求")
    else:
        print(f"    错误: {r1.text}")

    # 立即第二次请求（应该被限流）
    print("\n[2] 立即第二次请求（应该被限流）...")
    r2 = requests.post(f"{BASE_URL}/api/v1/reports", json=payload)
    print(f"    状态码: {r2.status_code}")
    if r2.status_code == 429:
        print(f"    结果: 成功被限流！")
        print(f"    消息: {r2.json()['detail']}")
    else:
        print(f"    警告: 未被限流 (预期429，实际{r2.status_code})")

    # 等待11秒后第三次请求
    print("\n[3] 等待11秒后重试...")
    for i in range(11):
        sys.stdout.write(f"\r    等待中... {11-i}秒 ")
        sys.stdout.flush()
        time.sleep(1)
    print("\n")

    r3 = requests.post(f"{BASE_URL}/api/v1/reports", json=payload)
    print(f"    状态码: {r3.status_code}")
    if r3.status_code == 202:
        print(f"    结果: 冷却后成功请求")
    else:
        print(f"    错误: {r3.text}")

    print("\n" + "=" * 60)
    print("测试完成！")
    print("=" * 60)

if __name__ == "__main__":
    test()
