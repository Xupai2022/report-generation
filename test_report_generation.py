#!/usr/bin/env python3
"""测试报告生成全流程性能"""

import time
import requests
import json

BASE_URL = "http://127.0.0.1:8000"

def test_report_generation():
    import sys
    import io
    if sys.platform == 'win32':
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

    print("=" * 70)
    print("报告生成全流程性能测试")
    print("=" * 70)

    # 1. 测试创建作业 (mock模式，不调用LLM)
    print("\n[测试1] 创建报告作业 (Mock模式)")
    start = time.time()
    resp = requests.post(
        f"{BASE_URL}/api/v1/reports",
        json={
            "input_id": "tenant_demo_2025-01",
            "template_id": "mss_technical_v2",
            "use_mock": True
        },
        timeout=120
    )
    create_time = time.time() - start
    print(f"   状态码: {resp.status_code}")
    print(f"   创建耗时: {create_time*1000:.2f}ms")

    if resp.status_code == 202:
        data = resp.json()
        job_id = data["job_id"]
        print(f"   作业ID: {job_id}")

        # 2. 轮询作业状态
        print(f"\n[测试2] 轮询作业状态")
        poll_count = 0
        poll_start = time.time()

        while True:
            time.sleep(1)
            poll_count += 1
            status_resp = requests.get(f"{BASE_URL}/api/v1/jobs/{job_id}")

            if status_resp.status_code == 200:
                status_data = status_resp.json()
                state = status_data["state"]
                print(f"   轮询 #{poll_count}: 状态={state}")

                if state in ["completed", "failed"]:
                    total_time = time.time() - poll_start
                    print(f"   总耗时: {total_time:.2f}s")
                    print(f"   最终状态: {state}")

                    if state == "completed":
                        # 3. 测试下载报告
                        print(f"\n[测试3] 下载生成的报告")
                        download_start = time.time()
                        report_resp = requests.get(
                            f"{BASE_URL}/api/v1/sessions/{status_data['session_id']}/download",
                            params={"template_id": "mss_technical_v2"}
                        )
                        download_time = time.time() - download_start

                        print(f"   下载状态: {report_resp.status_code}")
                        if report_resp.status_code == 200:
                            file_size = len(report_resp.content) / 1024 / 1024
                            print(f"   文件大小: {file_size:.2f}MB")
                            print(f"   下载耗时: {download_time*1000:.2f}ms")

                        # 4. 测试生成预览
                        print(f"\n[测试4] 生成预览图片")
                        preview_start = time.time()
                        preview_resp = requests.get(
                            f"{BASE_URL}/api/v1/reports/preview",
                            params={"job_id": job_id},
                            timeout=120
                        )
                        preview_time = time.time() - preview_start

                        print(f"   预览状态: {preview_resp.status_code}")
                        if preview_resp.status_code == 200:
                            preview_data = preview_resp.json()
                            print(f"   预览页数: {len(preview_data.get('slides', []))}")
                            print(f"   生成耗时: {preview_time:.2f}s")

                    break

            if poll_count > 60:  # 最多等待60秒
                print(f"   [超时] 等待超过60秒")
                break

    print("\n" + "=" * 70)
    print("测试完成")
    print("=" * 70)

if __name__ == "__main__":
    test_report_generation()
