#!/usr/bin/env python3
"""对比串行和并行预览生成性能"""

import sys
import io
import time
import shutil
from pathlib import Path

if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# 添加项目路径
sys.path.insert(0, str(Path(__file__).parent.parent))

from mss_ai_ppt_sample_assets.backend.modules.preview_generator import PPTPreviewGenerator
from mss_ai_ppt_sample_assets.backend import config


def find_test_pptx():
    """查找一个测试用的PPTX文件"""
    sessions_dir = Path("mss_ai_ppt_sample_assets/backend/outputs/sessions")
    pptx_files = sorted(
        sessions_dir.glob("*/report_*.pptx"),
        key=lambda p: p.stat().st_mtime,
        reverse=True
    )

    if not pptx_files:
        print("[ERROR] 未找到已生成的 PPTX 文件")
        print("请先生成一个报告，然后再运行测试")
        sys.exit(1)

    return pptx_files[0]


def test_serial_vs_parallel():
    """对比串行和并行性能"""
    pptx_path = find_test_pptx()
    print(f"测试文件: {pptx_path}")
    print(f"文件大小: {pptx_path.stat().st_size / 1024:.2f}KB\n")

    generator = PPTPreviewGenerator(cleanup_days=0)

    print("=" * 80)
    print("串行 vs 并行预览生成性能对比测试")
    print("=" * 80)

    # 测试1: 串行模式
    print("\n[测试 1/2] 串行模式")
    print("-" * 80)
    config.settings.preview_enable_parallel = False
    job_id_serial = f"test_serial_{int(time.time())}"

    start = time.perf_counter()
    try:
        images1, timings1 = generator.to_images_with_timings(pptx_path, job_id_serial)
        serial_time = time.perf_counter() - start
        print(f"✓ 串行生成完成")
        print(f"  耗时: {serial_time:.3f}s")
        print(f"  页数: {len(images1)}")
        print(f"  PPTX→PDF: {timings1.get('pptx_to_pdf_ms', 0):.0f}ms")
        print(f"  PDF→PNG: {timings1.get('pdf_to_images_ms', 0):.0f}ms")
        print(f"  处理速度: {len(images1)/serial_time:.2f} 页/秒")
    except Exception as e:
        print(f"[ERROR] 串行测试失败: {e}")
        return

    # 清理串行测试文件
    try:
        shutil.rmtree(config.PREVIEWS_DIR / job_id_serial)
    except:
        pass

    # 测试2: 并行模式
    print(f"\n[测试 2/2] 并行模式 (workers={config.settings.preview_parallel_workers})")
    print("-" * 80)
    config.settings.preview_enable_parallel = True
    job_id_parallel = f"test_parallel_{int(time.time())}"

    start = time.perf_counter()
    try:
        images2, timings2 = generator.to_images_with_timings(pptx_path, job_id_parallel)
        parallel_time = time.perf_counter() - start
        print(f"✓ 并行生成完成")
        print(f"  耗时: {parallel_time:.3f}s")
        print(f"  页数: {len(images2)}")
        print(f"  PPTX→PDF: {timings2.get('pptx_to_pdf_ms', 0):.0f}ms")
        print(f"  PDF→PNG: {timings2.get('pdf_to_images_ms', 0):.0f}ms")
        print(f"  处理速度: {len(images2)/parallel_time:.2f} 页/秒")
    except Exception as e:
        print(f"[ERROR] 并行测试失败: {e}")
        return

    # 清理并行测试文件
    try:
        shutil.rmtree(config.PREVIEWS_DIR / job_id_parallel)
    except:
        pass

    # 计算加速比
    speedup = serial_time / parallel_time
    pdf_speedup = timings1.get('pdf_to_images_ms', 1) / timings2.get('pdf_to_images_ms', 1)

    # 结果汇总
    print(f"\n{'='*80}")
    print("性能对比结果")
    print("=" * 80)
    print(f"\n总耗时:")
    print(f"  串行模式:     {serial_time:>8.3f}s")
    print(f"  并行模式:     {parallel_time:>8.3f}s")
    print(f"  加速比:       {speedup:>8.2f}x")
    print(f"  提升:         {(1 - parallel_time/serial_time)*100:>7.1f}%")

    print(f"\nPDF→PNG转换:")
    print(f"  串行模式:     {timings1.get('pdf_to_images_ms', 0):>8.0f}ms")
    print(f"  并行模式:     {timings2.get('pdf_to_images_ms', 0):>8.0f}ms")
    print(f"  加速比:       {pdf_speedup:>8.2f}x")

    print(f"\n处理速度:")
    print(f"  串行模式:     {len(images1)/serial_time:>8.2f} 页/秒")
    print(f"  并行模式:     {len(images2)/parallel_time:>8.2f} 页/秒")

    # 评估结果
    print(f"\n{'='*80}")
    print("性能评估")
    print("=" * 80)

    if speedup >= 1.8:
        print("✅ 优秀! 加速比 >1.8x,并行化效果显著")
    elif speedup >= 1.5:
        print("🟡 良好! 加速比 >1.5x,并行化有明显效果")
    elif speedup >= 1.3:
        print("🟡 一般! 加速比 >1.3x,并行化有轻微效果")
    else:
        print("❌ 失败! 加速比 <1.3x,建议回退串行模式")

    print(f"\n建议:")
    if speedup >= 1.5:
        print("  - 保持并行模式启用")
        print("  - 当前配置效果良好")
    elif speedup >= 1.3:
        print("  - 可考虑配合降低分辨率 (Matrix 2.0→1.5)")
        print("  - 或尝试调整 preview_parallel_workers")
    else:
        print("  - 建议回退串行模式: preview_enable_parallel=false")
        print("  - 或尝试进程池方案 (ProcessPoolExecutor)")

    print(f"\n{'='*80}\n")

    return speedup


if __name__ == "__main__":
    try:
        speedup = test_serial_vs_parallel()
        sys.exit(0 if speedup and speedup >= 1.3 else 1)
    except Exception as e:
        print(f"\n[ERROR] 测试异常: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
