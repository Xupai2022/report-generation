#!/usr/bin/env python3
"""压力测试脚本 - 测试并发性能和资源使用"""

import time
import requests
import threading
import statistics
import psutil
import os
from typing import List, Dict
from concurrent.futures import ThreadPoolExecutor, as_completed

BASE_URL = "http://127.0.0.1:8000"

class ResourceMonitor:
    """监控系统资源使用"""
    def __init__(self):
        self.process = psutil.Process(os.getpid())
        self.samples = []

    def sample(self):
        """采集当前资源使用"""
        cpu_percent = self.process.cpu_percent(interval=0.1)
        mem_info = self.process.memory_info()
        return {
            "cpu": cpu_percent,
            "memory_mb": mem_info.rss / 1024 / 1024,
            "timestamp": time.time()
        }

    def start_monitoring(self, interval=0.5):
        """开始后台监控"""
        self.monitoring = True
        def monitor_loop():
            while self.monitoring:
                self.samples.append(self.sample())
                time.sleep(interval)

        self.thread = threading.Thread(target=monitor_loop, daemon=True)
        self.thread.start()

    def stop_monitoring(self):
        """停止监控"""
        self.monitoring = False
        if hasattr(self, 'thread'):
            self.thread.join(timeout=2)

    def get_stats(self):
        """获取统计信息"""
        if not self.samples:
            return {}

        cpus = [s["cpu"] for s in self.samples]
        mems = [s["memory_mb"] for s in self.samples]

        return {
            "cpu_mean": statistics.mean(cpus),
            "cpu_max": max(cpus),
            "memory_mean_mb": statistics.mean(mems),
            "memory_max_mb": max(mems),
            "sample_count": len(self.samples)
        }

def concurrent_requests(url: str, num_requests: int, num_threads: int) -> Dict:
    """并发请求测试"""
    times = []
    errors = 0

    def make_request():
        start = time.time()
        try:
            resp = requests.get(url, timeout=30)
            elapsed = time.time() - start
            if resp.status_code < 400:
                return elapsed * 1000  # ms
            else:
                return None
        except Exception:
            return None

    with ThreadPoolExecutor(max_workers=num_threads) as executor:
        futures = [executor.submit(make_request) for _ in range(num_requests)]
        for future in as_completed(futures):
            result = future.result()
            if result is not None:
                times.append(result)
            else:
                errors += 1

    if not times:
        return {"error": "所有请求失败"}

    return {
        "total_requests": num_requests,
        "successful": len(times),
        "errors": errors,
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
    if sys.platform == 'win32':
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

    print("=" * 70)
    print("压力测试 - 并发性能与资源监控")
    print("=" * 70)

    # 测试场景
    scenarios = [
        {
            "name": "轻负载 (10 并发, 50 请求)",
            "url": f"{BASE_URL}/api/v1/templates",
            "requests": 50,
            "threads": 10
        },
        {
            "name": "中负载 (20 并发, 100 请求)",
            "url": f"{BASE_URL}/api/v1/templates",
            "requests": 100,
            "threads": 20
        },
        {
            "name": "高负载 (50 并发, 200 请求)",
            "url": f"{BASE_URL}/api/v1/templates",
            "requests": 200,
            "threads": 50
        },
    ]

    for scenario in scenarios:
        print(f"\n[场景] {scenario['name']}")
        print(f"   URL: {scenario['url']}")
        print(f"   请求数: {scenario['requests']}, 并发数: {scenario['threads']}")

        # 启动资源监控
        monitor = ResourceMonitor()
        monitor.start_monitoring(interval=0.2)

        # 执行并发测试
        start_time = time.time()
        result = concurrent_requests(scenario['url'], scenario['requests'], scenario['threads'])
        total_time = time.time() - start_time

        # 停止监控
        monitor.stop_monitoring()
        resource_stats = monitor.get_stats()

        if "error" in result:
            print(f"   [ERROR] {result['error']}")
            continue

        # 输出性能指标
        print(f"\n   性能指标:")
        print(f"   - 成功请求: {result['successful']}/{result['total_requests']}")
        print(f"   - 失败请求: {result['errors']}")
        print(f"   - 总耗时: {total_time:.2f}s")
        print(f"   - 吞吐量: {result['successful']/total_time:.2f} req/s")
        print(f"   - 响应时间 (最小/平均/中位/P95/P99/最大):")
        print(f"     {result['min']:.2f} / {result['mean']:.2f} / {result['median']:.2f} / {result['p95']:.2f} / {result['p99']:.2f} / {result['max']:.2f} ms")

        # 输出资源使用
        if resource_stats:
            print(f"\n   资源使用:")
            print(f"   - CPU: 平均 {resource_stats['cpu_mean']:.1f}%, 峰值 {resource_stats['cpu_max']:.1f}%")
            print(f"   - 内存: 平均 {resource_stats['memory_mean_mb']:.1f}MB, 峰值 {resource_stats['memory_max_mb']:.1f}MB")

    print("\n" + "=" * 70)
    print("压力测试完成")
    print("=" * 70)

if __name__ == "__main__":
    main()
