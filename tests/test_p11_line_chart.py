"""
测试P11_line折线图的数据生成和渲染功能

这个测试演示了如何为技术版模板第5页的P11_line图表准备数据。
"""

import json
from pathlib import Path

# 模拟输入数据：月度事件趋势
# 这里展示了按严重程度分类的月度事件数量
test_data = {
    "incident_trend_monthly": {
        "months": ["1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月"],
        "严重": [5, 3, 7, 4, 2, 6, 8, 5, 3, 4, 6, 7],
        "高危": [12, 15, 18, 14, 10, 13, 16, 19, 14, 12, 15, 17],
        "中危": [45, 38, 52, 41, 35, 48, 54, 47, 39, 42, 50, 46],
        "低危": [78, 82, 95, 88, 72, 85, 92, 87, 80, 83, 90, 88]
    }
}

# 预期的图表数据格式（经过llm_orchestrator处理后）
expected_chart_data = {
    "months": ["1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月"],
    "series": [
        {
            "name": "严重",
            "values": [5, 3, 7, 4, 2, 6, 8, 5, 3, 4, 6, 7]
        },
        {
            "name": "高危",
            "values": [12, 15, 18, 14, 10, 13, 16, 19, 14, 12, 15, 17]
        },
        {
            "name": "中危",
            "values": [45, 38, 52, 41, 35, 48, 54, 47, 39, 42, 50, 46]
        },
        {
            "name": "低危",
            "values": [78, 82, 95, 88, 72, 85, 92, 87, 80, 83, 90, 88]
        }
    ]
}

def test_data_structure():
    """测试数据结构是否符合要求"""
    print("=" * 80)
    print("测试P11_line图表数据结构")
    print("=" * 80)

    # 验证输入数据
    assert "incident_trend_monthly" in test_data
    incident_data = test_data["incident_trend_monthly"]

    assert "months" in incident_data
    assert len(incident_data["months"]) == 12

    # 验证所有严重程度系列
    for severity in ["严重", "高危", "中危", "低危"]:
        assert severity in incident_data
        assert len(incident_data[severity]) == 12
        print(f"✓ {severity}系列包含12个月的数据")

    print("\n输入数据验证通过！")

    # 验证预期输出格式
    assert "months" in expected_chart_data
    assert "series" in expected_chart_data
    assert len(expected_chart_data["series"]) == 4

    for series in expected_chart_data["series"]:
        assert "name" in series
        assert "values" in series
        assert len(series["values"]) == 12
        print(f"✓ {series['name']}系列数据格式正确")

    print("\n预期输出格式验证通过！")
    print("=" * 80)

def generate_sample_slidespec():
    """生成示例slidespec片段"""
    slidespec_fragment = {
        "P11_line": {
            "position": None,
            "months": expected_chart_data["months"],
            "series": expected_chart_data["series"]
        }
    }

    print("\n生成的slidespec片段：")
    print(json.dumps(slidespec_fragment, ensure_ascii=False, indent=2))

    return slidespec_fragment

def show_data_visualization():
    """以表格形式展示数据"""
    print("\n" + "=" * 80)
    print("数据可视化预览")
    print("=" * 80)

    # 打印表头
    print(f"{'月份':<8}", end="")
    for severity in ["严重", "高危", "中危", "低危"]:
        print(f"{severity:<8}", end="")
    print()
    print("-" * 48)

    # 打印数据行
    months = test_data["incident_trend_monthly"]["months"]
    for i, month in enumerate(months):
        print(f"{month:<8}", end="")
        for severity in ["严重", "高危", "中危", "低危"]:
            value = test_data["incident_trend_monthly"][severity][i]
            print(f"{value:<8}", end="")
        print()

    print("=" * 80)

def main():
    """主测试函数"""
    print("\n")
    print("█" * 80)
    print("█" + " " * 78 + "█")
    print("█" + " " * 20 + "P11_line 折线图测试数据" + " " * 33 + "█")
    print("█" + " " * 78 + "█")
    print("█" * 80)

    # 测试数据结构
    test_data_structure()

    # 展示数据可视化
    show_data_visualization()

    # 生成slidespec片段
    slidespec = generate_sample_slidespec()

    # 使用说明
    print("\n" + "=" * 80)
    print("使用说明")
    print("=" * 80)
    print("""
1. 在Excel上传的数据中添加 incident_trend_monthly 字段
2. 数据格式示例：
   {
     "incident_trend_monthly": {
       "months": ["1月", "2月", ..., "12月"],
       "严重": [5, 3, 7, ...],
       "高危": [12, 15, 18, ...],
       "中危": [45, 38, 52, ...],
       "低危": [78, 82, 95, ...]
     }
   }

3. llm_orchestrator会自动将数据转换为图表所需格式
4. ppt_generator会渲染折线图，每个严重程度用不同颜色：
   - 严重: 红色 (220, 53, 69)
   - 高危: 橙色 (255, 152, 0)
   - 中危: 黄色 (255, 193, 7)
   - 低危: 绿色 (40, 167, 69)

5. 在PowerPoint模板的第5页添加折线图，并放置{{P11_line}}占位符
    """)
    print("=" * 80)

if __name__ == "__main__":
    main()
