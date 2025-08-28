#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
测试分页符插入功能
"""

import sys
import os
from pathlib import Path

# 添加项目根目录到Python路径
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from autoword.vnext.simple_pipeline import SimplePipeline
from autoword.vnext.core import VNextConfig

def test_page_break_insertion():
    """测试分页符插入"""
    
    # 配置
    config = VNextConfig()
    pipeline = SimplePipeline(config)
    
    # 测试文档路径（请替换为你的实际文档路径）
    test_doc = "test_document.docx"  # 请替换为实际的测试文档
    
    if not os.path.exists(test_doc):
        print(f"测试文档不存在: {test_doc}")
        print("请将你的测试文档重命名为 test_document.docx 或修改此脚本中的路径")
        return
    
    # 测试用户意图
    test_intents = [
        "插入分页符在封面后",
        "在封面后面插入分页符",
        "添加分页符分隔封面和正文",
        "插入分页符"
    ]
    
    for i, intent in enumerate(test_intents):
        print(f"\n=== 测试 {i+1}: {intent} ===")
        
        try:
            result = pipeline.process_document(test_doc, intent)
            
            if result.success:
                print(f"✅ 处理成功!")
                print(f"输出文件: {result.output_path}")
                print(f"执行的操作: {len(result.plan.get('ops', []))} 个")
                
                # 显示执行的操作
                for j, op in enumerate(result.plan.get('ops', [])):
                    print(f"  操作 {j+1}: {op.get('operation_type')} - {op}")
                    
            else:
                print(f"❌ 处理失败: {result.error}")
                
        except Exception as e:
            print(f"❌ 异常: {e}")
        
        print("-" * 50)

if __name__ == "__main__":
    print("AutoWord 分页符插入功能测试")
    print("=" * 50)
    test_page_break_insertion()