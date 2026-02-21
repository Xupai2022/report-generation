#!/usr/bin/env python3
"""测试渐进式加载预览图"""

import sys
import io
import requests
import time

if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

BASE_URL = "http://localhost:8000"


def test_progressive_preview():
    """测试渐进式预览加载"""

    print("=" * 70)
    print("渐进式预览加载测试")
    print("=" * 70)

    # 1. 先生成一个报告
    print("\n[步骤 1/4] 生成报告...")
    create_resp = requests.post(
        f"{BASE_URL}/api/v1/reports",
        json={
            "input_id": "tenant_acme_2025-11",
            "template_id": "mss_technical_v2",
            "use_mock": True
        }
    )

    if create_resp.status_code != 200:
        print(f"[ERROR] 创建报告失败: {create_resp.status_code}")
        print(create_resp.text)
        return

    job_id = create_resp.json()["data"]["job_id"]
    print(f"✓ 报告已创建: {job_id}")

    # 等待报告生成完成
    print("\n等待报告生成...")
    for i in range(60):
        time.sleep(1)
        status_resp = requests.get(f"{BASE_URL}/api/v1/jobs/{job_id}")
        if status_resp.status_code == 200:
            status = status_resp.json()["data"]["status"]
            if status == "completed":
                print(f"✓ 报告生成完成")
                break
            elif status == "failed":
                print(f"[ERROR] 报告生成失败")
                return
        print(f"  等待中... {i+1}s", end='\r')

    # 2. 测试完整预览（不分页）
    print("\n\n[步骤 2/4] 测试完整预览（不分页）")
    start = time.time()
    full_resp = requests.get(
        f"{BASE_URL}/api/v1/reports/{job_id}/preview",
        params={"regenerate_if_missing": True}
    )
    full_time = time.time() - start

    if full_resp.status_code != 200:
        print(f"[ERROR] 预览失败: {full_resp.status_code}")
        return

    full_data = full_resp.json()["data"]
    total_slides = len(full_data["images"])

    print(f"✓ 完整预览已生成")
    print(f"  总页数: {total_slides}")
    print(f"  耗时: {full_time:.2f}s")

    # 3. 测试渐进式加载 - 仅加载第1页
    print("\n[步骤 3/4] 渐进式加载 - 仅加载第1页")
    start = time.time()
    page1_resp = requests.get(
        f"{BASE_URL}/api/v1/reports/{job_id}/preview",
        params={
            "page": 1,
            "page_size": 1,
            "force_regenerate": False  # 使用缓存
        }
    )
    page1_time = time.time() - start

    if page1_resp.status_code != 200:
        print(f"[ERROR] 分页预览失败: {page1_resp.status_code}")
        return

    page1_data = page1_resp.json()["data"]

    print(f"✓ 第1页预览已加载")
    print(f"  当前页: {page1_data['current_page']}")
    print(f"  每页大小: {page1_data['page_size']}")
    print(f"  总页数: {page1_data['total_pages']}")
    print(f"  总幻灯片数: {page1_data['total_slides']}")
    print(f"  还有更多: {page1_data['has_more']}")
    print(f"  返回图片: {len(page1_data['images'])}")
    print(f"  耗时: {page1_time:.2f}s")

    # 4. 测试渐进式加载 - 加载第1-3页
    print("\n[步骤 4/4] 渐进式加载 - 加载前3页")
    start = time.time()
    page3_resp = requests.get(
        f"{BASE_URL}/api/v1/reports/{job_id}/preview",
        params={
            "page": 1,
            "page_size": 3,
            "force_regenerate": False
        }
    )
    page3_time = time.time() - start

    page3_data = page3_resp.json()["data"]

    print(f"✓ 前3页预览已加载")
    print(f"  返回图片: {len(page3_data['images'])}")
    print(f"  耗时: {page3_time:.2f}s")

    # 性能对比
    print("\n" + "=" * 70)
    print("性能对比")
    print("=" * 70)
    print(f"完整预览耗时:     {full_time:.2f}s  (所有 {total_slides} 页)")
    print(f"首页预览耗时:     {page1_time:.2f}s  (仅第 1 页)")
    print(f"前3页预览耗时:    {page3_time:.2f}s  (前 3 页)")

    speedup = (full_time / page1_time) if page1_time > 0 else 0
    print(f"\n首页加载速度提升: {speedup:.1f}x 倍")

    print("\n使用建议:")
    print("  1. 首次访问:  GET /reports/{id}/preview?page=1&page_size=1")
    print("  2. 滚动加载:  GET /reports/{id}/preview?page=2&page_size=3")
    print("  3. 全部加载:  GET /reports/{id}/preview (无分页参数)")

    print("\n示例URL:")
    print(f"  {BASE_URL}/api/v1/reports/{job_id}/preview?page=1&page_size=1")


if __name__ == "__main__":
    try:
        test_progressive_preview()
    except KeyboardInterrupt:
        print("\n\n测试已中断")
    except Exception as e:
        print(f"\n[ERROR] 测试失败: {e}")
        import traceback
        traceback.print_exc()
