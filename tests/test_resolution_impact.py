#!/usr/bin/env python3
"""测试分辨率对性能的影响 - IO瓶颈验证"""

import sys
import io
import time
import shutil
from pathlib import Path

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


def test_resolution_impact():
    """测试不同分辨率对性能和文件大小的影响"""
    pptx_path = find_test_pptx()
    print(f"测试文件: {pptx_path}\n")

    generator = PPTPreviewGenerator(cleanup_days=0)

    print("=" * 80)
    print("分辨率对性能影响测试")
    print("=" * 80)

    # 测试不同Matrix值
    matrix_configs = [
        (1.0, "72 DPI (低质量)"),
        (1.5, "108 DPI (中质量)"),
        (2.0, "144 DPI (高质量,当前)"),
        (2.5, "180 DPI (超高质量)"),
    ]

    results = []

    for matrix_scale, description in matrix_configs:
        print(f"\n[测试] Matrix={matrix_scale} - {description}")
        print("-" * 80)

        # 配置
        config.settings.preview_matrix_scale = matrix_scale
        config.settings.preview_enable_parallel = False  # 使用串行避免干扰

        job_id = f"test_matrix{matrix_scale}_{int(time.time())}"
        output_dir = config.PREVIEWS_DIR / job_id

        try:
            # 生成PDF (只需一次)
            if not results:  # 首次生成PDF
                pdf_path = generator._pptx_to_pdf(pptx_path, output_dir)
            else:
                output_dir.mkdir(parents=True, exist_ok=True)
                # 复用之前的PDF
                prev_dir = config.PREVIEWS_DIR / results[0]['job_id']
                pdf_files = list(prev_dir.glob("*.pdf"))
                if pdf_files:
                    import shutil as sh
                    pdf_path = output_dir / pdf_files[0].name
                    sh.copy2(pdf_files[0], pdf_path)
                else:
                    pdf_path = generator._pptx_to_pdf(pptx_path, output_dir)

            # 测试PDF→PNG转换
            start = time.perf_counter()
            images, timings = generator._pdf_to_images_with_timings(pdf_path, output_dir)
            elapsed = time.perf_counter() - start

            # 计算文件大小
            total_size_kb = sum(img.stat().st_size for img in images) / 1024
            avg_size_kb = total_size_kb / len(images)

            print(f"  PDF→PNG耗时: {elapsed:.3f}s")
            print(f"  处理速度: {len(images)/elapsed:.2f} 页/秒")
            print(f"  总大小: {total_size_kb:.1f}KB")
            print(f"  平均大小: {avg_size_kb:.1f}KB/张")

            results.append({
                'matrix_scale': matrix_scale,
                'description': description,
                'time': elapsed,
                'total_size_kb': total_size_kb,
                'avg_size_kb': avg_size_kb,
                'page_count': len(images),
                'job_id': job_id
            })

        except Exception as e:
            print(f"  [ERROR] 测试失败: {e}")
            continue

    # 结果汇总
    print(f"\n{'='*80}")
    print("分辨率性能对比")
    print("=" * 80)
    print(f"\n{'Matrix':<8} {'DPI':<6} {'耗时(s)':<10} {'速度(页/s)':<12} {'平均大小(KB)':<15} {'总大小(KB)'}")
    print("-" * 80)

    baseline = results[2] if len(results) > 2 else results[0]  # 2.0作为基准

    for r in results:
        speedup_vs_baseline = baseline['time'] / r['time']
        size_vs_baseline = r['total_size_kb'] / baseline['total_size_kb']

        dpi = int(r['matrix_scale'] * 72)
        speed = r['page_count'] / r['time']

        marker = " ← 当前" if r['matrix_scale'] == 2.0 else ""

        print(
            f"{r['matrix_scale']:<8} {dpi:<6} {r['time']:<10.3f} "
            f"{speed:<12.2f} {r['avg_size_kb']:<15.1f} {r['total_size_kb']:<.1f}{marker}"
        )

    # IO瓶颈分析
    print(f"\n{'='*80}")
    print("IO瓶颈分析")
    print("=" * 80)

    if len(results) >= 2:
        # 对比最低分辨率和最高分辨率
        lowest = results[0]
        highest = results[-1]

        time_ratio = highest['time'] / lowest['time']
        size_ratio = highest['total_size_kb'] / lowest['total_size_kb']

        print(f"\n最低分辨率 (Matrix={lowest['matrix_scale']})")
        print(f"  耗时: {lowest['time']:.3f}s")
        print(f"  总大小: {lowest['total_size_kb']:.1f}KB")

        print(f"\n最高分辨率 (Matrix={highest['matrix_scale']})")
        print(f"  耗时: {highest['time']:.3f}s")
        print(f"  总大小: {highest['total_size_kb']:.1f}KB")

        print(f"\n对比:")
        print(f"  文件大小增加: {size_ratio:.2f}x")
        print(f"  耗时增加: {time_ratio:.2f}x")

        if abs(time_ratio - size_ratio) / size_ratio < 0.3:
            print(f"\n✅ 确认IO瓶颈: 耗时与文件大小成正比")
            print("   分辨率越高→文件越大→IO耗时越长")
            print("\n   建议: 降低分辨率是最有效的优化手段")
        else:
            print(f"\n🟡 IO影响不明显: 耗时与文件大小不成正比")
            print("   可能还有其他瓶颈因素")

    # 推荐配置
    print(f"\n{'='*80}")
    print("推荐配置")
    print("=" * 80)

    if len(results) >= 2:
        # Matrix 1.5作为推荐
        matrix_15 = next((r for r in results if r['matrix_scale'] == 1.5), None)
        matrix_20 = next((r for r in results if r['matrix_scale'] == 2.0), None)

        if matrix_15 and matrix_20:
            speedup = matrix_20['time'] / matrix_15['time']
            size_reduction = (1 - matrix_15['total_size_kb'] / matrix_20['total_size_kb']) * 100

            print(f"\n推荐: Matrix=1.5 (108 DPI)")
            print(f"  相比当前配置 (Matrix=2.0):")
            print(f"    - 性能提升: {speedup:.2f}x ({(speedup-1)*100:.1f}%)")
            print(f"    - 文件减小: {size_reduction:.1f}%")
            print(f"    - 质量损失: 可接受 (144 DPI → 108 DPI)")
            print(f"\n  .env配置:")
            print(f"    PREVIEW_MATRIX_SCALE=1.5")

    # 清理测试文件
    print(f"\n清理测试文件...")
    for r in results:
        try:
            shutil.rmtree(config.PREVIEWS_DIR / r['job_id'])
        except:
            pass
    print("✓ 清理完成")


if __name__ == "__main__":
    try:
        test_resolution_impact()
    except Exception as e:
        print(f"\n[ERROR] 测试失败: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
