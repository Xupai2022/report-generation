#!/usr/bin/env python3
"""性能基准测试脚本 - 测量实际接口响应时间"""

import time
import requests
import statistics
from typing import List, Dict

BASE_URL = "http://127.0.0.1:8000"

def measure_endpoint(url: str, method: str = "GET", json_data: dict = None, iterations: int = 10) -> Dict:
    """测量接口响应时间"""
    times = []

    for i in range(iterations):
        start = time.time()
        try:
            if method == "GET":
                resp = requests.get(url, timeout=120)
            elif method == "POST":
                resp = requests.post(url, json=json_data, timeout=120)
            elapsed = time.time() - start

            if resp.status_code < 400:
                times.append(elapsed * 1000)  # 转换为毫秒
            else:
                print(f"  [WARN] 请求 {i+1} 失败: {resp.status_code}")
        except Exception as e:
            print(f"  [ERROR] 请求 {i+1} 异常: {e}")

    if not times:
        return {"error": "所有请求失败"}

    return {
        "count": len(times),
        "min": min(times),
        "max": max(times),
        "mean": statistics.mean(times),
        "median": statistics.median(times),
        "p95": sorted(times)[int(len(times) * 0.95)] if len(times) > 1 else times[0],
        "p99": sorted(times)[int(len(times) * 0.99)] if len(times) > 1 else times[0],
    }

def main():
    import sys
    import io
    # Fix Windows console encoding
    if sys.platform == 'win32':
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

    print("=" * 60)
    print("性能基准测试 - 实际测量结果")
    print("=" * 60)

    tests = [
        {
            "name": "根路径 GET /",
            "url": f"{BASE_URL}/",
            "method": "GET",
            "iterations": 20
        },
        {
            "name": "模板列表 GET /api/v1/templates",
            "url": f"{BASE_URL}/api/v1/templates",
            "method": "GET",
            "iterations": 20
        },
        {
            "name": "健康检查 GET /api/v1/system/health",
            "url": f"{BASE_URL}/api/v1/system/health",
            "method": "GET",
            "iterations": 20
        },
    ]

    # 测试作业创建（仅创建，不等待完成）
    create_test = {
        "name": "创建报告 POST /api/v1/reports (仅创建作业)",
        "url": f"{BASE_URL}/api/v1/reports",
        "method": "POST",
        "json_data": {
            "input_id": "tenant_acme_2025-12",  # 使用实际存在的输入
            "template_id": "mss_technical_v2",
            "use_mock": True  # 使用 mock 避免真实 LLM 调用
        },
        "iterations": 5  # 减少次数避免限流
    }

    for test in tests:
        print(f"\n[TEST] {test['name']}")
        print(f"   迭代次数: {test['iterations']}")
        result = measure_endpoint(test['url'], test['method'], iterations=test['iterations'])

        if "error" in result:
            print(f"   [ERROR] {result['error']}")
            continue

        print(f"   样本数: {result['count']}")
        print(f"   最小值: {result['min']:.2f}ms")
        print(f"   平均值: {result['mean']:.2f}ms")
        print(f"   中位数: {result['median']:.2f}ms")
        print(f"   P95:   {result['p95']:.2f}ms")
        print(f"   P99:   {result['p99']:.2f}ms")
        print(f"   最大值: {result['max']:.2f}ms")

    # 测试作业创建（会受限流影响）
    print(f"\n[TEST] {create_test['name']}")
    print(f"   [WARNING] 注意: 会触发 30秒 IP 限流，测试间隔会很长")
    print(f"   迭代次数: {create_test['iterations']}")

    times = []
    for i in range(create_test['iterations']):
        print(f"   正在测试 {i+1}/{create_test['iterations']}...", end=" ")
        start = time.time()
        try:
            resp = requests.post(create_test['url'], json=create_test['json_data'], timeout=120)
            elapsed = time.time() - start

            if resp.status_code == 202:
                times.append(elapsed * 1000)
                print(f"OK {elapsed*1000:.2f}ms (作业已创建)")
            elif resp.status_code == 429:
                print(f"[RATE_LIMIT] 限流中，等待 30 秒...")
                time.sleep(30)
                # 重试
                start = time.time()
                resp = requests.post(create_test['url'], json=create_test['json_data'], timeout=120)
                elapsed = time.time() - start
                if resp.status_code == 202:
                    times.append(elapsed * 1000)
                    print(f"   重试成功: {elapsed*1000:.2f}ms")
            else:
                print(f"[ERROR] 状态码 {resp.status_code}")
        except Exception as e:
            print(f"[ERROR] 异常: {e}")

        # 等待避免限流（除了最后一次）
        if i < create_test['iterations'] - 1:
            time.sleep(31)

    if times:
        result = {
            "min": min(times),
            "max": max(times),
            "mean": statistics.mean(times),
            "median": statistics.median(times),
        }
        print(f"\n   样本数: {len(times)}")
        print(f"   最小值: {result['min']:.2f}ms")
        print(f"   平均值: {result['mean']:.2f}ms")
        print(f"   中位数: {result['median']:.2f}ms")
        print(f"   最大值: {result['max']:.2f}ms")

    print("\n" + "=" * 60)
    print("测试完成")
    print("=" * 60)

if __name__ == "__main__":
    main()
