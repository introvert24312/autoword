#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
测试封面保护功能 - 确保分页符后封面不受样式修改影响
"""

import sys
import os
from pathlib import Path

# 添加项目根目录到Python路径
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from autoword.vnext.simple_pipeline import SimplePipeline
from autoword.vnext.core import VNextConfig

def test_cover_protection():
    """测试封面保护功能"""
    
    print("🛡️ 测试封面保护功能")
    print("=" * 50)
    
    # 配置
    config = VNextConfig()
    pipeline = SimplePipeline(config)
    
    # 测试文档路径
    test_doc = input("请输入你的测试文档路径: ").strip()
    if not test_doc or not os.path.exists(test_doc):
        print("❌ 文档不存在")
        return
    
    print(f"📄 使用文档: {test_doc}")
    
    # 测试意图：同时插入分页符和修改样式
    test_intent = "插入分页符在封面后，1级标题设置为楷体小四号2倍行距，正文设置为宋体小四号2倍行距"
    
    print(f"\n🎯 测试意图: {test_intent}")
    print("\n预期结果:")
    print("- 在封面后插入分页符")
    print("- 只修改分页符后的正文样式")
    print("- 封面样式保持不变")
    
    try:
        result = pipeline.process_document(test_doc, test_intent)
        
        if result.success:
            print(f"\n✅ 处理成功!")
            print(f"📁 输出文件: {result.output_path}")
            print(f"🔧 执行了 {len(result.plan.get('ops', []))} 个操作:")
            
            for i, op in enumerate(result.plan.get('ops', [])):
                op_type = op.get('operation_type')
                print(f"  {i+1}. {op_type}")
                
                if op_type == "insert_page_break":
                    print(f"     位置: {op.get('position', '未指定')}")
                elif op_type == "set_style_rule":
                    style_name = op.get('target_style_name', '未指定')
                    font_info = op.get('font', {})
                    print(f"     样式: {style_name}")
                    if 'east_asian' in font_info:
                        print(f"     字体: {font_info['east_asian']}")
                    if 'size_pt' in font_info:
                        print(f"     大小: {font_info['size_pt']}pt")
            
            print(f"\n🎉 请检查输出文档，确认:")
            print("1. 封面后有分页符")
            print("2. 封面样式未改变")
            print("3. 正文样式已更新")
            
        else:
            print(f"\n❌ 处理失败: {result.error}")
            
    except Exception as e:
        print(f"\n❌ 异常: {e}")
        import traceback
        traceback.print_exc()

def debug_paragraph_analysis():
    """调试段落分析"""
    print("\n" + "=" * 50)
    print("🔍 调试段落分析")
    
    config = VNextConfig()
    pipeline = SimplePipeline(config)
    
    test_doc = input("请输入测试文档路径: ").strip()
    if not test_doc or not os.path.exists(test_doc):
        print("❌ 文档不存在")
        return
    
    try:
        # 提取文档结构来分析段落
        structure = pipeline._extract_structure(test_doc)
        
        print(f"\n📊 文档分析结果:")
        print(f"总段落数: {len(structure['paragraphs'])}")
        print(f"标题数: {len(structure['headings'])}")
        
        print(f"\n📝 前10个段落:")
        for i, para in enumerate(structure['paragraphs'][:10]):
            print(f"  {i+1}. 页码:{para['page_number']} | 大纲级别:{para['outline_level']} | 样式:{para['style_name']}")
            print(f"      内容: {para['preview_text'][:50]}...")
            print(f"      是否封面: {para['is_cover']}")
            print()
            
    except Exception as e:
        print(f"❌ 分析失败: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    print("AutoWord 封面保护功能测试")
    print("=" * 50)
    
    choice = input("选择测试:\n1. 封面保护测试\n2. 段落分析调试\n请输入 1 或 2: ").strip()
    
    if choice == "1":
        test_cover_protection()
    elif choice == "2":
        debug_paragraph_analysis()
    else:
        print("无效选择，执行封面保护测试")
        test_cover_protection()