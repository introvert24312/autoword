#!/usr/bin/env python3
"""
测试中文引号修复功能
"""

import sys
import json
from pathlib import Path

# 添加项目根目录到Python路径
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

def test_chinese_quotes_fix():
    """测试中文引号修复功能"""
    print("========================================")
    print("      测试中文引号修复功能")
    print("========================================")
    print()
    
    try:
        from autoword.core.llm_client import LLMClient
        
        # 创建LLM客户端实例
        llm_client = LLMClient()
        
        # 测试包含中文引号的JSON字符串
        test_json = '''{"tasks": [{"id": "task_1","source_comment_id": "comment_1","type": "delete","locator": {"by": "find","value": "摘要"},"instruction": "从目录中删除"摘要"项，保留其他内容和参考文献","risk": "low"}]}'''
        
        print("原始JSON（包含中文引号）:")
        print("=" * 50)
        print(test_json)
        print("=" * 50)
        print()
        
        # 尝试直接解析（应该失败）
        try:
            json.loads(test_json)
            print("❌ 意外：原始JSON解析成功了")
        except json.JSONDecodeError as e:
            print(f"✅ 预期：原始JSON解析失败 - {e}")
            print()
        
        # 使用修复函数
        print("应用中文引号修复...")
        fixed_json = llm_client._fix_chinese_quotes_in_json(test_json)
        
        print("修复后的JSON:")
        print("=" * 50)
        print(fixed_json)
        print("=" * 50)
        print()
        
        # 尝试解析修复后的JSON
        try:
            parsed = json.loads(fixed_json)
            print("✅ 修复后JSON解析成功!")
            print(f"任务数量: {len(parsed.get('tasks', []))}")
            
            # 显示修复后的内容
            for task in parsed.get('tasks', []):
                print(f"- 任务ID: {task.get('id')}")
                print(f"  指令: {task.get('instruction')}")
            
            return True
            
        except json.JSONDecodeError as e:
            print(f"❌ 修复后JSON仍然解析失败: {e}")
            return False
            
    except Exception as e:
        print(f"❌ 测试过程中发生异常: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    try:
        success = test_chinese_quotes_fix()
        if success:
            print()
            print("🎉 中文引号修复功能测试成功！")
            sys.exit(0)
        else:
            print()
            print("❌ 中文引号修复功能测试失败。")
            sys.exit(1)
    except KeyboardInterrupt:
        print("\n测试被用户中断")
        sys.exit(1)
    except Exception as e:
        print(f"测试过程中发生异常: {e}")
        sys.exit(1)