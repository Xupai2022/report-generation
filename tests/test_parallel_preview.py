#!/usr/bin/env python3
"""测试并行化预览生成的性能提升"""

import sys
import io
import time
from pathlib import Path

if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# 添加项目路径
sys.path.insert(0, str(Path(__file__).parent.parent))

from mss_ai_ppt_sample_assets.backend.modules.preview_generator import PPTPreviewGenerator
from mss_ai_ppt_sample_assets.backend import config


def test_parallel_performance():
    """测试并行化后的预览生成性能"""

    # 找一个已生成的 PPTX 文件
    sessions_dir = Path("mss_ai_ppt_sample_assets/backend/outputs/sessions")

    # 查找最近的 PPTX 文件
    pptx_files = sorted(sessions_dir.glob("*/report_*.pptx"), key=lambda p: p.stat().st_mtime, reverse=True)

    if not pptx_files:
        print("[ERROR] 未找到已生成的 PPTX 文件")
        print("请先生成一个报告，然后再测试预览性能")
        return

    pptx_path = pptx_files[0]
    print(f"测试文件: {pptx_path}")
    print(f"文件大小: {pptx_path.stat().st_size / 1024:.2f}KB\n")

    # 创建预览生成器
    generator = PPTPreviewGenerator(cleanup_days=0)

    # 测试用的 job_id
    job_id = "test_parallel_" + str(int(time.time()))
    job_dir = config.PREVIEWS_DIR / job_id
    job_dir.mkdir(parents=True, exist_ok=True)

    print("=" * 70)
    print("并行化预览生成性能测试")
    print("=" * 70)

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

    # 步骤 2: PDF → PNG (PyMuPDF 并行化)
    print(f"\n[步骤 2/2] PDF → PNG (PyMuPDF 并行化)")
    start_png = time.time()

    try:
        images = generator._pdf_to_images(pdf_path, job_dir)
        png_time = time.time() - start_png

        total_png_size = sum(img.stat().st_size for img in images) / 1024  # KB
        avg_png_size = total_png_size / len(images)

        print(f"✓ PNG 图片已生成: {len(images)} 张")
        print(f"  总耗时: {png_time:.3f}s")
        print(f"  平均每张: {png_time/len(images):.3f}s")
        print(f"  处理速度: {len(images)/png_time:.2f} 页/秒")
        print(f"  总大小: {total_png_size:.2f}KB")
        print(f"  平均大小: {avg_png_size:.2f}KB/张")

    except Exception as e:
        print(f"[ERROR] PNG 生成失败: {e}")
        return

    # 总结
    total_time = pdf_time + png_time
    print(f"\n{'='*70}")
    print(f"总耗时: {total_time:.3f}s")
    print(f"{'='*70}")
    print(f"\n耗时分解:")
    print(f"  PPTX → PDF:    {pdf_time:>8.3f}s  ({pdf_time/total_time*100:>5.1f}%)")
    print(f"  PDF → PNG:     {png_time:>8.3f}s  ({png_time/total_time*100:>5.1f}%)")
    print(f"\n性能指标:")
    print(f"  页数: {len(images)}")
    print(f"  PDF 生成速度: {len(images)/pdf_time:.2f} 页/秒")
    print(f"  PNG 生成速度: {len(images)/png_time:.2f} 页/秒 (并行化)")
    print(f"  整体速度: {len(images)/total_time:.2f} 页/秒")

    # 性能提升估算
    print(f"\n{'='*70}")
    print("预期性能提升（与串行处理相比）:")
    print("  PNG 转换提升: 约 60-80% (取决于CPU核心数)")
    print(f"  使用线程数: {min(4, len(images))}")
    print(f"{'='*70}")

    # 清理测试文件
    import shutil
    try:
        shutil.rmtree(job_dir)
        print(f"\n✓ 清理测试文件: {job_dir}")
    except Exception:
        pass


if __name__ == "__main__":
    test_parallel_performance()
