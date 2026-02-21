#!/usr/bin/env python3
"""深入分析GIL和IO瓶颈 - CPU时间 vs 墙上时间"""

import sys
import io
import time
import shutil
from pathlib import Path
import threading

if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

sys.path.insert(0, str(Path(__file__).parent.parent))

from mss_ai_ppt_sample_assets.backend.modules.preview_generator import PPTPreviewGenerator
from mss_ai_ppt_sample_assets.backend import config


def find_test_pptx():
    """查找测试PPT文件"""
    sessions_dir = Path("mss_ai_ppt_sample_assets/backend/outputs/sessions")
    pptx_files = sorted(
        sessions_dir.glob("*/report_*.pptx"),
        key=lambda p: p.stat().st_mtime,
        reverse=True
    )
    if not pptx_files:
        print("[ERROR] 未找到测试PPTX文件")
        sys.exit(1)
    return pptx_files[0]


def test_cpu_time_analysis():
    """分析CPU时间占比 - 诊断GIL是否真正释放"""
    import resource  # Unix only

    pptx_path = find_test_pptx()
    generator = PPTPreviewGenerator(cleanup_days=0)

    print("=" * 80)
    print("CPU时间分析 - 诊断GIL释放效果")
    print("=" * 80)

    # 测试串行模式
    print("\n[1/2] 串行模式 CPU分析")
    config.settings.preview_enable_parallel = False
    job_id_serial = f"test_cpu_serial_{int(time.time())}"

    wall_start = time.perf_counter()
    cpu_start = time.process_time()  # CPU time

    images, _ = generator.to_images_with_timings(pptx_path, job_id_serial)

    wall_time = time.perf_counter() - wall_start
    cpu_time = time.process_time() - cpu_start

    cpu_ratio_serial = cpu_time / wall_time * 100

    print(f"  墙上时间: {wall_time:.3f}s")
    print(f"  CPU时间:  {cpu_time:.3f}s")
    print(f"  CPU占比:  {cpu_ratio_serial:.1f}%")
    print(f"  IO等待:   {wall_time - cpu_time:.3f}s ({(1 - cpu_time/wall_time)*100:.1f}%)")

    shutil.rmtree(config.PREVIEWS_DIR / job_id_serial)

    # 测试并行模式
    print(f"\n[2/2] 并行模式 CPU分析 (workers={config.settings.preview_parallel_workers})")
    config.settings.preview_enable_parallel = True
    job_id_parallel = f"test_cpu_parallel_{int(time.time())}"

    wall_start = time.perf_counter()
    cpu_start = time.process_time()

    images, _ = generator.to_images_with_timings(pptx_path, job_id_parallel)

    wall_time = time.perf_counter() - wall_start
    cpu_time = time.process_time() - cpu_start

    cpu_ratio_parallel = cpu_time / wall_time * 100

    print(f"  墙上时间: {wall_time:.3f}s")
    print(f"  CPU时间:  {cpu_time:.3f}s")
    print(f"  CPU占比:  {cpu_ratio_parallel:.1f}%")
    print(f"  IO等待:   {wall_time - cpu_time:.3f}s ({(1 - cpu_time/wall_time)*100:.1f}%)")

    shutil.rmtree(config.PREVIEWS_DIR / job_id_parallel)

    # 分析结论
    print(f"\n{'='*80}")
    print("GIL释放诊断")
    print("=" * 80)

    if cpu_ratio_parallel > cpu_ratio_serial * 1.5:
        print("✅ GIL已释放: 并行模式CPU占比显著提升")
        print(f"   串行CPU占比: {cpu_ratio_serial:.1f}%")
        print(f"   并行CPU占比: {cpu_ratio_parallel:.1f}% (提升 {cpu_ratio_parallel/cpu_ratio_serial:.2f}x)")
    else:
        print("⚠️ GIL可能未完全释放: CPU占比提升不明显")
        print(f"   串行CPU占比: {cpu_ratio_serial:.1f}%")
        print(f"   并行CPU占比: {cpu_ratio_parallel:.1f}%")

    # IO瓶颈分析
    print(f"\n{'='*80}")
    print("IO瓶颈分析")
    print("=" * 80)

    if cpu_ratio_serial < 50:
        print(f"⚠️ 发现IO瓶颈: CPU占比仅{cpu_ratio_serial:.1f}%")
        print("   建议:")
        print("   - 检查磁盘IO性能 (HDD vs SSD)")
        print("   - 考虑使用内存文件系统")
        print("   - 减少PNG写入大小")
    else:
        print(f"✅ CPU密集型任务: CPU占比{cpu_ratio_serial:.1f}%")
        print("   瓶颈在计算,不在IO")


def test_thread_scaling():
    """测试不同线程数的加速效果"""
    pptx_path = find_test_pptx()
    generator = PPTPreviewGenerator(cleanup_days=0)

    print("\n" + "=" * 80)
    print("线程数扩展性测试")
    print("=" * 80)

    # 先测串行基准
    print("\n[基准] 串行模式")
    config.settings.preview_enable_parallel = False
    job_id = f"test_baseline_{int(time.time())}"

    start = time.perf_counter()
    images, _ = generator.to_images_with_timings(pptx_path, job_id)
    baseline_time = time.perf_counter() - start
    print(f"  耗时: {baseline_time:.3f}s")

    shutil.rmtree(config.PREVIEWS_DIR / job_id)

    # 测试不同worker数
    config.settings.preview_enable_parallel = True
    worker_counts = [2, 4, 8]

    results = []
    for workers in worker_counts:
        config.settings.preview_parallel_workers = workers
        job_id = f"test_workers{workers}_{int(time.time())}"

        print(f"\n[测试] {workers} workers")
        start = time.perf_counter()
        images, _ = generator.to_images_with_timings(pptx_path, job_id)
        elapsed = time.perf_counter() - start
        speedup = baseline_time / elapsed

        print(f"  耗时: {elapsed:.3f}s")
        print(f"  加速比: {speedup:.2f}x")
        print(f"  效率: {speedup/workers*100:.1f}%")

        results.append((workers, elapsed, speedup))
        shutil.rmtree(config.PREVIEWS_DIR / job_id)

    # 分析扩展性
    print(f"\n{'='*80}")
    print("扩展性分析")
    print("=" * 80)
    print(f"\n{'Workers':<10} {'时间(s)':<12} {'加速比':<10} {'并行效率'}")
    print("-" * 80)
    print(f"{'1 (串行)':<10} {baseline_time:<12.3f} {'1.00x':<10} {'100.0%'}")

    for workers, elapsed, speedup in results:
        efficiency = speedup / workers * 100
        print(f"{workers:<10} {elapsed:<12.3f} {speedup:<10.2f}x {efficiency:.1f}%")

    # 诊断
    print(f"\n诊断结果:")
    if results[-1][2] < 1.5:  # 8 workers加速比<1.5
        print("❌ 扩展性差: 增加线程数无明显提升")
        print("   可能原因:")
        print("   1. GIL限制 (尽管PyMuPDF声称释放GIL)")
        print("   2. IO瓶颈 (磁盘写入速度限制)")
        print("   3. 内存带宽瓶颈")
        print("\n   建议:")
        print("   - 尝试进程池 (ProcessPoolExecutor)")
        print("   - 降低图片质量减少IO")
        print("   - 使用RAM disk")
    else:
        print("✅ 扩展性良好: 增加线程数有效提升性能")


if __name__ == "__main__":
    try:
        # Windows不支持resource模块,跳过CPU时间分析
        if sys.platform != 'win32':
            test_cpu_time_analysis()
        else:
            print("注意: Windows不支持process_time分析,跳过CPU时间测试\n")

        test_thread_scaling()

    except Exception as e:
        print(f"\n[ERROR] 测试失败: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
