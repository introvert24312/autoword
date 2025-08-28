#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
测试无页码封面保护功能
专门针对封面和目录没有页码的情况
"""

import sys
import os
from pathlib import Path

# 添加项目根目录到Python路径
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from autoword.vnext.simple_pipeline import SimplePipeline
from autoword.vnext.core import VNextConfig

def test_no_page_number_cover_protection():
    """测试无页码封面保护"""
    
    print("🛡️ 测试无页码封面保护功能")
    print("=" * 60)
    print("适用场景：封面和目录没有页码，只有正文有页码")
    print("=" * 60)
    
    # 配置
    config = VNextConfig()
    pipeline = SimplePipeline(config)
    
    # 测试文档路径
    test_doc = input("请输入你的测试文档路径: ").strip()
    if not test_doc or not os.path.exists(test_doc):
        print("❌ 文档不存在")
        return
    
    print(f"📄 使用文档: {test_doc}")
    
    # 先分析文档结构
    print("\n🔍 分析文档结构...")
    try:
        structure = pipeline._extract_structure(test_doc)
        
        print(f"📊 文档信息:")
        print(f"  总段落数: {len(structure['paragraphs'])}")
        print(f"  标题数: {len(structure['headings'])}")
        
        # 显示前15个段落的详细信息
        print(f"\n📝 前15个段落分析:")
        for i, para in enumerate(structure['paragraphs'][:15]):
            page_num = para.get('page_number', 0)
            is_cover = para.get('is_cover', False)
            outline = para.get('outline_level', 10)
            style = para.get('style_name', '未知')
            content = para.get('preview_text', '')[:40]
            
            status = "🛡️封面" if is_cover else "📄正文"
            page_info = f"无页码" if page_num <= 1 else f"第{page_num}页"
            
            print(f"  {i+1:2d}. {status} | {page_info} | 大纲:{outline} | {style}")
            print(f"      内容: {content}...")
            
    except Exception as e:
        print(f"❌ 文档分析失败: {e}")
        return
    
    # 测试样式修改
    print(f"\n🎯 测试样式修改...")
    test_intent = "1级标题设置为楷体小四号2倍行距"
    
    print(f"意图: {test_intent}")
    print("预期结果: 只修改正文中的1级标题，封面和目录保持不变")
    
    try:
        result = pipeline.process_document(test_doc, test_intent)
        
        if result.success:
            print(f"\n✅ 处理成功!")
            print(f"📁 输出文件: {result.output_path}")
            print(f"🔧 执行了 {len(result.plan.get('ops', []))} 个操作")
            
            print(f"\n🎉 请检查输出文档，确认:")
            print("1. 封面内容样式未改变（无页码区域）")
            print("2. 目录内容样式未改变（无页码区域）") 
            print("3. 正文中的1级标题样式已更新（有页码区域）")
            
        else:
            print(f"\n❌ 处理失败: {result.error}")
            
    except Exception as e:
        print(f"\n❌ 异常: {e}")
        import traceback
        traceback.print_exc()

def test_combined_operations():
    """测试组合操作"""
    print("\n" + "=" * 60)
    print("🔧 测试组合操作（插入分页符 + 样式修改）")
    
    config = VNextConfig()
    pipeline = SimplePipeline(config)
    
    test_doc = input("请输入测试文档路径: ").strip()
    if not test_doc or not os.path.exists(test_doc):
        print("❌ 文档不存在")
        return
    
    # 组合操作：插入分页符并修改样式
    combined_intent = "插入分页符在封面后，1级标题设置为楷体小四号2倍行距，正文设置为宋体小四号1.5倍行距"
    
    print(f"🎯 组合意图: {combined_intent}")
    print("\n预期结果:")
    print("1. 在封面后插入分页符")
    print("2. 封面区域样式保持不变")
    print("3. 正文区域样式按要求修改")
    
    try:
        result = pipeline.process_document(test_doc, combined_intent)
        
        if result.success:
            print(f"\n✅ 组合操作成功!")
            print(f"📁 输出文件: {result.output_path}")
            print(f"🔧 执行的操作:")
            
            for i, op in enumerate(result.plan.get('ops', [])):
                op_type = op.get('operation_type')
                print(f"  {i+1}. {op_type}")
                
                if op_type == "insert_page_break":
                    print(f"     位置: {op.get('position', '未指定')}")
                elif op_type == "set_style_rule":
                    style_name = op.get('target_style_name', '未指定')
                    print(f"     目标样式: {style_name}")
                    
        else:
            print(f"\n❌ 组合操作失败: {result.error}")
            
    except Exception as e:
        print(f"\n❌ 组合操作异常: {e}")

if __name__ == "__main__":
    print("AutoWord 无页码封面保护测试")
    print("=" * 60)
    
    choice = input("选择测试:\n1. 无页码封面保护测试\n2. 组合操作测试\n请输入 1 或 2: ").strip()
    
    if choice == "1":
        test_no_page_number_cover_protection()
    elif choice == "2":
        test_combined_operations()
    else:
        print("无效选择，执行封面保护测试")
        test_no_page_number_cover_protection()