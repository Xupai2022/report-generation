#!/usr/bin/env python3
"""预览生成耗时分解测试 - 分别测量 PPTX→PDF 和 PDF→PNG 的时间"""

import sys
import io
import time
from pathlib import Path

if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# 添加项目路径
sys.path.insert(0, str(Path(__file__).parent))

from mss_ai_ppt_sample_assets.backend.modules.preview_generator import PPTPreviewGenerator
from mss_ai_ppt_sample_assets.backend import config

def test_preview_breakdown():
    """测试预览生成各步骤的耗时"""

    # 找一个已生成的 PPTX 文件
    sessions_dir = Path("mss_ai_ppt_sample_assets/backend/outputs/sessions")

    # 查找最近的 PPTX 文件
    pptx_files = sorted(sessions_dir.glob("*/report_*.pptx"), key=lambda p: p.stat().st_mtime, reverse=True)

    if not pptx_files:
        print("[ERROR] 未找到已生成的 PPTX 文件")
        print("请先运行 test_full_generation_time.py 生成报告")
        return

    pptx_path = pptx_files[0]
    print(f"测试文件: {pptx_path}")
    print(f"文件大小: {pptx_path.stat().st_size / 1024:.2f}KB\n")

    # 创建预览生成器
    generator = PPTPreviewGenerator(cleanup_days=0)  # 禁用清理避免干扰

    # 测试用的 job_id
    job_id = "test_breakdown_" + str(int(time.time()))
    job_dir = config.PREVIEWS_DIR / job_id
    job_dir.mkdir(parents=True, exist_ok=True)

    print("=" * 60)
    print("预览生成耗时分解测试")
    print("=" * 60)

    # 步骤 1: PPTX → PDF (LibreOffice)
    print("\n[步骤 1/2] PPTX → PDF (LibreOffice)")
    start_pdf = time.time()

    try:
        pdf_path = generator._pptx_to_pdf(pptx_path, job_dir)
        pdf_time = time.time() - start_pdf

        pdf_size = pdf_path.stat().st_size / 1024  # KB
        print(f"✓ PDF 已生成: {pdf_path.name}")
        print(f"  耗时: {pdf_time:.3f}s")
        print(f"  文件大小: {pdf_size:.2f}KB")

    except Exception as e:
        print(f"[ERROR] PDF 生成失败: {e}")
        return

    # 步骤 2: PDF → PNG (PyMuPDF)
    print(f"\n[步骤 2/2] PDF → PNG (PyMuPDF)")
    start_png = time.time()

    try:
        images = generator._pdf_to_images(pdf_path, job_dir)
        png_time = time.time() - start_png

        total_png_size = sum(img.stat().st_size for img in images) / 1024  # KB
        avg_png_size = total_png_size / len(images)

        print(f"✓ PNG 图片已生成: {len(images)} 张")
        print(f"  耗时: {png_time:.3f}s")
        print(f"  平均每张: {png_time/len(images):.3f}s")
        print(f"  总大小: {total_png_size:.2f}KB")
        print(f"  平均大小: {avg_png_size:.2f}KB/张")

    except Exception as e:
        print(f"[ERROR] PNG 生成失败: {e}")
        return

    # 总结
    total_time = pdf_time + png_time
    print(f"\n{'='*60}")
    print(f"总耗时: {total_time:.3f}s")
    print(f"{'='*60}")
    print(f"\n耗时分解:")
    print(f"  PPTX → PDF:    {pdf_time:>8.3f}s  ({pdf_time/total_time*100:>5.1f}%)")
    print(f"  PDF → PNG:     {png_time:>8.3f}s  ({png_time/total_time*100:>5.1f}%)")
    print(f"\n性能指标:")
    print(f"  页数: {len(images)}")
    print(f"  PDF 生成速度: {len(images)/pdf_time:.2f} 页/秒")
    print(f"  PNG 生成速度: {len(images)/png_time:.2f} 页/秒")
    print(f"  整体速度: {len(images)/total_time:.2f} 页/秒")

    # 清理测试文件
    import shutil
    try:
        shutil.rmtree(job_dir)
        print(f"\n✓ 清理测试文件: {job_dir}")
    except:
        pass


def test_multiple_runs():
    """多次运行取平均值"""

    sessions_dir = Path("mss_ai_ppt_sample_assets/backend/outputs/sessions")
    pptx_files = sorted(sessions_dir.glob("*/report_*.pptx"), key=lambda p: p.stat().st_mtime, reverse=True)

    if not pptx_files:
        print("[ERROR] 未找到已生成的 PPTX 文件")
        return

    pptx_path = pptx_files[0]
    generator = PPTPreviewGenerator(cleanup_days=0)

    runs = 3
    pdf_times = []
    png_times = []

    print(f"\n{'='*60}")
    print(f"多次运行测试（{runs} 次）")
    print(f"{'='*60}\n")

    for i in range(runs):
        job_id = f"test_run_{i}_" + str(int(time.time()))
        job_dir = config.PREVIEWS_DIR / job_id
        job_dir.mkdir(parents=True, exist_ok=True)

        print(f"运行 {i+1}/{runs}...")

        # PPTX → PDF
        start = time.time()
        pdf_path = generator._pptx_to_pdf(pptx_path, job_dir)
        pdf_time = time.time() - start
        pdf_times.append(pdf_time)

        # PDF → PNG
        start = time.time()
        images = generator._pdf_to_images(pdf_path, job_dir)
        png_time = time.time() - start
        png_times.append(png_time)

        print(f"  PDF: {pdf_time:.3f}s, PNG: {png_time:.3f}s, 总计: {pdf_time+png_time:.3f}s")

        # 清理
        import shutil
        shutil.rmtree(job_dir)

        # 避免太快
        if i < runs - 1:
            time.sleep(1)

    import statistics
    print(f"\n{'='*60}")
    print(f"统计结果（{runs} 次运行）")
    print(f"{'='*60}")
    print(f"\nPPTX → PDF:")
    print(f"  最小值: {min(pdf_times):.3f}s")
    print(f"  最大值: {max(pdf_times):.3f}s")
    print(f"  平均值: {statistics.mean(pdf_times):.3f}s")
    print(f"  中位数: {statistics.median(pdf_times):.3f}s")

    print(f"\nPDF → PNG:")
    print(f"  最小值: {min(png_times):.3f}s")
    print(f"  最大值: {max(png_times):.3f}s")
    print(f"  平均值: {statistics.mean(png_times):.3f}s")
    print(f"  中位数: {statistics.median(png_times):.3f}s")

    total_times = [pdf + png for pdf, png in zip(pdf_times, png_times)]
    print(f"\n总耗时:")
    print(f"  最小值: {min(total_times):.3f}s")
    print(f"  最大值: {max(total_times):.3f}s")
    print(f"  平均值: {statistics.mean(total_times):.3f}s")
    print(f"  中位数: {statistics.median(total_times):.3f}s")


if __name__ == "__main__":
    # 单次详细测试
    test_preview_breakdown()

    # 多次运行取平均
    print("\n\n")
    test_multiple_runs()
