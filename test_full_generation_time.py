#!/usr/bin/env python3
"""完整报告生成耗时测试 - 从提交到完成"""

import time
import requests
import sys
import io

if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

BASE_URL = "http://127.0.0.1:8000"

def test_full_generation(input_id: str, template_id: str, use_mock: bool = True):
    """测试完整报告生成流程"""

    print(f"\n{'='*60}")
    print(f"完整报告生成耗时测试")
    print(f"{'='*60}")
    print(f"输入: {input_id}")
    print(f"模板: {template_id}")
    print(f"模式: {'Mock LLM' if use_mock else 'Real LLM'}")
    print(f"{'='*60}\n")

    # Step 1: 创建作业
    print("[1/4] 创建作业...")
    start_total = time.time()
    start_create = time.time()

    try:
        resp = requests.post(
            f"{BASE_URL}/api/v1/reports",
            json={
                "input_id": input_id,
                "template_id": template_id,
                "use_mock": use_mock
            },
            timeout=120
        )
        create_time = time.time() - start_create

        if resp.status_code != 202:
            print(f"[ERROR] 创建失败: {resp.status_code}")
            print(resp.text)
            return

        result = resp.json()
        job_id = result['data']['job_id']
        print(f"✓ 作业已创建: {job_id}")
        print(f"  耗时: {create_time*1000:.2f}ms\n")

    except Exception as e:
        print(f"[ERROR] 创建异常: {e}")
        return

    # Step 2: 轮询状态直到完成
    print("[2/4] 等待生成完成...")
    poll_start = time.time()
    poll_count = 0
    last_progress = -1

    while True:
        try:
            resp = requests.get(f"{BASE_URL}/api/v1/jobs/{job_id}/status", timeout=30)
            poll_count += 1

            if resp.status_code != 200:
                print(f"[ERROR] 状态查询失败: {resp.status_code}")
                break

            status_data = resp.json()['data']
            status = status_data['status']
            progress = status_data.get('progress', 0)
            message = status_data.get('message', '')

            # 显示进度变化
            if progress != last_progress:
                print(f"  进度: {progress}% - {message}")
                last_progress = progress

            if status == 'completed':
                generation_time = time.time() - poll_start
                print(f"✓ 生成完成")
                print(f"  总轮询次数: {poll_count}")
                print(f"  生成耗时: {generation_time:.2f}s\n")
                break
            elif status == 'failed':
                print(f"[ERROR] 生成失败: {status_data.get('last_error', 'Unknown error')}")
                return

            # 等待后再次轮询
            time.sleep(2)

        except Exception as e:
            print(f"[ERROR] 轮询异常: {e}")
            break

    # Step 3: 下载报告
    print("[3/4] 下载报告...")
    download_start = time.time()

    try:
        resp = requests.get(f"{BASE_URL}/api/v1/reports/{job_id}/download", timeout=60)
        download_time = time.time() - download_start

        if resp.status_code == 200:
            file_size = len(resp.content) / 1024  # KB
            print(f"✓ 报告已下载")
            print(f"  文件大小: {file_size:.2f}KB")
            print(f"  下载耗时: {download_time*1000:.2f}ms\n")
        else:
            print(f"[ERROR] 下载失败: {resp.status_code}\n")
    except Exception as e:
        print(f"[ERROR] 下载异常: {e}\n")

    # Step 4: 生成预览
    print("[4/4] 生成预览...")
    preview_start = time.time()

    try:
        resp = requests.get(f"{BASE_URL}/api/v1/reports/{job_id}/preview", timeout=120)
        preview_time = time.time() - preview_start

        if resp.status_code == 200:
            preview_data = resp.json()['data']
            image_count = len(preview_data.get('images', []))
            print(f"✓ 预览已生成")
            print(f"  图片数量: {image_count}")
            print(f"  预览耗时: {preview_time:.2f}s\n")
        else:
            print(f"[ERROR] 预览失败: {resp.status_code}\n")
    except Exception as e:
        print(f"[ERROR] 预览异常: {e}\n")

    # 总结
    total_time = time.time() - start_total
    print(f"{'='*60}")
    print(f"完整流程总耗时: {total_time:.2f}s")
    print(f"{'='*60}")
    print(f"\n时间分解:")
    print(f"  1. 创建作业:     {create_time*1000:>8.2f}ms")
    print(f"  2. 生成报告:     {generation_time:>8.2f}s")
    print(f"  3. 下载报告:     {download_time*1000:>8.2f}ms")
    print(f"  4. 生成预览:     {preview_time:>8.2f}s")
    print(f"  总计:           {total_time:>8.2f}s")
    print(f"\n性能指标:")
    print(f"  生成速度占比:   {generation_time/total_time*100:.1f}%")
    print(f"  预览速度占比:   {preview_time/total_time*100:.1f}%")
    print(f"  接口开销占比:   {(create_time+download_time)/total_time*100:.1f}%")


def main():
    print("\n正在测试完整报告生成流程...")
    print("使用 Mock LLM 模式（避免真实 API 调用）\n")

    # 测试 1: 技术模板
    test_full_generation(
        input_id="tenant_acme_2025-12",
        template_id="mss_technical_v2",
        use_mock=True
    )

    print("\n\n等待 30 秒避免限流...\n")
    time.sleep(31)

    # 测试 2: 管理层模板
    test_full_generation(
        input_id="tenant_acme_2025-12",
        template_id="mss_executive_v2",
        use_mock=True
    )


if __name__ == "__main__":
    main()
