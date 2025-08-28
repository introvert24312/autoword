#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
解决封面行距变化问题的测试脚本
"""

import sys
import os
from pathlib import Path

# 添加项目根目录到Python路径
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from autoword.vnext.simple_pipeline import SimplePipeline
from autoword.vnext.core import VNextConfig

def solve_cover_line_spacing_issue():
    """解决封面行距问题"""
    
    print("🔧 解决封面行距变化问题")
    print("=" * 50)
    
    # 配置
    config = VNextConfig()
    pipeline = SimplePipeline(config)
    
    # 测试文档路径
    test_doc = input("请输入你的测试文档路径（或直接回车使用 test_document.docx）: ").strip()
    if not test_doc:
        test_doc = "test_document.docx"
    
    if not os.path.exists(test_doc):
        print(f"❌ 测试文档不存在: {test_doc}")
        return
    
    print(f"📄 使用文档: {test_doc}")
    
    # 步骤1: 先插入分页符分隔封面和正文
    print("\n步骤1: 插入分页符分隔封面和正文")
    try:
        result1 = pipeline.process_document(test_doc, "插入分页符在封面后")
        if result1.success:
            print(f"✅ 分页符插入成功! 输出: {result1.output_path}")
            test_doc = result1.output_path  # 使用处理后的文档继续
        else:
            print(f"❌ 分页符插入失败: {result1.error}")
            return
    except Exception as e:
        print(f"❌ 分页符插入异常: {e}")
        return
    
    # 步骤2: 修改正文行距
    print("\n步骤2: 修改正文行距（现在封面应该不会受影响）")
    try:
        result2 = pipeline.process_document(test_doc, "1级标题设置为楷体小四号2倍行距")
        if result2.success:
            print(f"✅ 行距修改成功! 输出: {result2.output_path}")
            print("\n🎉 问题解决!")
            print(f"最终文档: {result2.output_path}")
            print("\n现在封面和正文被分页符分隔，样式修改只会影响正文部分。")
        else:
            print(f"❌ 行距修改失败: {result2.error}")
    except Exception as e:
        print(f"❌ 行距修改异常: {e}")

def test_combined_operation():
    """测试组合操作"""
    print("\n" + "=" * 50)
    print("🔧 测试组合操作（一次性解决）")
    
    config = VNextConfig()
    pipeline = SimplePipeline(config)
    
    test_doc = input("请输入测试文档路径: ").strip()
    if not test_doc or not os.path.exists(test_doc):
        print("❌ 文档不存在")
        return
    
    # 组合意图：同时插入分页符和修改样式
    combined_intent = "插入分页符在封面后，然后1级标题设置为楷体小四号2倍行距"
    
    try:
        result = pipeline.process_document(test_doc, combined_intent)
        if result.success:
            print(f"✅ 组合操作成功! 输出: {result.output_path}")
            print(f"执行了 {len(result.plan.get('ops', []))} 个操作:")
            for i, op in enumerate(result.plan.get('ops', [])):
                print(f"  {i+1}. {op.get('operation_type')}")
        else:
            print(f"❌ 组合操作失败: {result.error}")
    except Exception as e:
        print(f"❌ 组合操作异常: {e}")

if __name__ == "__main__":
    print("AutoWord 封面行距问题解决方案")
    print("=" * 50)
    
    choice = input("选择测试方式:\n1. 分步解决（推荐）\n2. 组合操作\n请输入 1 或 2: ").strip()
    
    if choice == "1":
        solve_cover_line_spacing_issue()
    elif choice == "2":
        test_combined_operation()
    else:
        print("无效选择，使用分步解决方案")
        solve_cover_line_spacing_issue()