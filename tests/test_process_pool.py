#!/usr/bin/env python3
"""测试进程池方案 - 规避GIL限制"""

import sys
import io
import time
import shutil
from pathlib import Path

if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

sys.path.insert(0, str(Path(__file__).parent.parent))

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


def render_page_process(args):
    """进程池渲染函数 - 必须在顶层定义供pickle序列化"""
    import fitz
    pdf_path_str, page_num, output_dir_str, matrix_scale = args

    try:
        doc = fitz.open(pdf_path_str)
        page = doc.load_page(page_num)
        mat = fitz.Matrix(matrix_scale, matrix_scale)
        pix = page.get_pixmap(matrix=mat)

        target = Path(output_dir_str) / f"slide{page_num+1}.png"
        pix.save(str(target))

        doc.close()
        return target
    except Exception as e:
        return f"ERROR: {e}"


def test_process_pool():
    """测试进程池方案"""
    from concurrent.futures import ProcessPoolExecutor
    import fitz

    pptx_path = find_test_pptx()
    print(f"测试文件: {pptx_path}\n")

    # 先生成PDF
    from mss_ai_ppt_sample_assets.backend.modules.preview_generator import PPTPreviewGenerator
    generator = PPTPreviewGenerator(cleanup_days=0)

    job_id = f"test_process_{int(time.time())}"
    output_dir = config.PREVIEWS_DIR / job_id
    output_dir.mkdir(parents=True, exist_ok=True)

    pdf_path = generator._pptx_to_pdf(pptx_path, output_dir)
    doc = fitz.open(str(pdf_path))
    page_count = len(doc)
    doc.close()

    print("=" * 80)
    print("进程池 vs 线程池性能对比")
    print("=" * 80)

    # 测试1: 串行基准
    print("\n[1/3] 串行基准")
    config.settings.preview_enable_parallel = False

    start = time.perf_counter()
    images, _ = generator._pdf_to_images_with_timings(pdf_path, output_dir)
    serial_time = time.perf_counter() - start

    print(f"  耗时: {serial_time:.3f}s")
    print(f"  速度: {page_count/serial_time:.2f} 页/秒")

    # 清理
    for img in images:
        img.unlink()

    # 测试2: 线程池
    print(f"\n[2/3] 线程池 (workers={config.settings.preview_parallel_workers})")
    config.settings.preview_enable_parallel = True

    start = time.perf_counter()
    images, _ = generator._pdf_to_images_with_timings(pdf_path, output_dir)
    thread_time = time.perf_counter() - start

    print(f"  耗时: {thread_time:.3f}s")
    print(f"  加速比: {serial_time/thread_time:.2f}x")
    print(f"  速度: {page_count/thread_time:.2f} 页/秒")

    # 清理
    for img in images:
        img.unlink()

    # 测试3: 进程池
    print(f"\n[3/3] 进程池 (workers={config.settings.preview_parallel_workers})")

    start = time.perf_counter()

    # 准备参数
    args_list = [
        (str(pdf_path), i, str(output_dir), config.settings.preview_matrix_scale)
        for i in range(page_count)
    ]

    with ProcessPoolExecutor(max_workers=config.settings.preview_parallel_workers) as executor:
        results = list(executor.map(render_page_process, args_list))

    process_time = time.perf_counter() - start

    print(f"  耗时: {process_time:.3f}s")
    print(f"  加速比: {serial_time/process_time:.2f}x")
    print(f"  速度: {page_count/process_time:.2f} 页/秒")

    # 分析结果
    print(f"\n{'='*80}")
    print("性能对比结果")
    print("=" * 80)
    print(f"\n{'模式':<12} {'耗时(s)':<10} {'加速比':<10} {'速度(页/秒)'}")
    print("-" * 80)
    print(f"{'串行':<12} {serial_time:<10.3f} {'1.00x':<10} {page_count/serial_time:.2f}")
    print(f"{'线程池':<12} {thread_time:<10.3f} {serial_time/thread_time:<10.2f}x {page_count/thread_time:.2f}")
    print(f"{'进程池':<12} {process_time:<10.3f} {serial_time/process_time:<10.2f}x {page_count/process_time:.2f}")

    # 诊断
    print(f"\n{'='*80}")
    print("GIL影响诊断")
    print("=" * 80)

    thread_speedup = serial_time / thread_time
    process_speedup = serial_time / process_time

    if process_speedup > thread_speedup * 1.5:
        print("✅ 确认GIL限制: 进程池显著优于线程池")
        print(f"   线程池加速: {thread_speedup:.2f}x (受GIL限制)")
        print(f"   进程池加速: {process_speedup:.2f}x (无GIL限制)")
        print("\n   建议: 使用进程池替代线程池")
    elif process_speedup < 1.5:
        print("⚠️ 非GIL问题: 进程池也无法提速")
        print(f"   线程池加速: {thread_speedup:.2f}x")
        print(f"   进程池加速: {process_speedup:.2f}x")
        print("\n   诊断: 瓶颈在IO或其他因素,不是GIL")
        print("   可能原因:")
        print("   - 磁盘IO瓶颈 (PNG写入速度限制)")
        print("   - 单个PDF渲染本身就很快,并行开销大于收益")
        print("   - 内存带宽限制")
        print("\n   建议:")
        print("   - 降低图片质量/分辨率 (Matrix 2.0→1.5)")
        print("   - 检查磁盘性能 (HDD vs SSD)")
        print("   - 接受串行方案,优化其他环节")
    else:
        print("🟡 进程池有改善但不显著")
        print(f"   线程池加速: {thread_speedup:.2f}x")
        print(f"   进程池加速: {process_speedup:.2f}x")

    # 清理
    shutil.rmtree(output_dir)


if __name__ == "__main__":
    try:
        test_process_pool()
    except Exception as e:
        print(f"\n[ERROR] 测试失败: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
