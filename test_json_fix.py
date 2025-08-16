#!/usr/bin/env python3
"""
测试JSON修复功能
"""

import sys
import json
from pathlib import Path

# 添加项目根目录到Python路径
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

def test_json_fixes():
    """测试JSON修复功能"""
    print("========================================")
    print("      JSON修复功能测试")
    print("========================================")
    print()
    
    from autoword.core.llm_client import LLMClient
    
    # 创建LLM客户端实例
    llm_client = LLMClient()
    
    # 测试各种有问题的JSON格式
    test_cases = [
        # 尾随逗号
        '{"tasks": [{"id": "task_1", "type": "rewrite",}]}',
        
        # 缺少引号
        '{tasks: [{id: "task_1", type: "rewrite"}]}',
        
        # 带markdown标记
        '```json\n{"tasks": [{"id": "task_1", "type": "rewrite"}]}\n```',
        
        # 带注释
        '{\n  // 这是任务列表\n  "tasks": [{"id": "task_1", "type": "rewrite"}]\n}',
        
        # 缺少逗号
        '{"tasks": [{"id": "task_1" "type": "rewrite"}]}',
        
        # 混合问题
        '```json\n{\n  tasks: [\n    {id: "task_1", type: "rewrite",}\n  ],\n}\n```'
    ]
    
    print("测试JSON修复策略:")
    print()
    
    success_count = 0
    
    for i, test_json in enumerate(test_cases, 1):
        print(f"测试用例 {i}:")
        print(f"原始: {repr(test_json[:50])}...")
        
        try:
            # 尝试原始解析
            json.loads(test_json)
            print("✅ 原始JSON有效")
            success_count += 1
        except json.JSONDecodeError:
            # 尝试清理修复
            try:
                cleaned = llm_client._clean_json_content(test_json)
                json.loads(cleaned)
                print("✅ 清理修复成功")
                success_count += 1
            except json.JSONDecodeError:
                # 尝试激进修复
                try:
                    fixed = llm_client._aggressive_json_fix(test_json)
                    parsed = json.loads(fixed)
                    print("✅ 激进修复成功")
                    print(f"修复后: {json.dumps(parsed, ensure_ascii=False)}")
                    success_count += 1
                except json.JSONDecodeError as e:
                    print(f"❌ 修复失败: {e}")
        
        print()
    
    print(f"修复成功率: {success_count}/{len(test_cases)} ({success_count/len(test_cases)*100:.1f}%)")
    
    return success_count == len(test_cases)


def test_real_llm_call():
    """测试真实的LLM调用和JSON解析"""
    print("========================================")
    print("      真实LLM调用测试")
    print("========================================")
    print()
    
    try:
        from autoword.core.llm_client import LLMClient, ModelType
        
        # 配置API密钥
        api_keys = {
            "claude": "sk-3w1JFbWUq7tKjpLlopdkISQ9F6fpLhHx5viD0frh43ESE9Io",
            "gpt": "sk-NhjnJtqlZMx4PGTqvkGlH4POT82HHBrBnBbWOat99Bs5VZXi"
        }
        
        llm_client = LLMClient(api_keys=api_keys)
        
        # 测试简单的JSON生成
        system_prompt = """你是一个任务规划助手。请根据用户的要求生成JSON格式的任务列表。

返回格式必须是有效的JSON，格式如下：
{
  "tasks": [
    {
      "id": "task_1",
      "type": "rewrite",
      "instruction": "重写内容",
      "locator": {"by": "find", "value": "目标文本"}
    }
  ]
}

重要：只返回JSON，不要添加任何解释文字或markdown标记。"""
        
        user_prompt = "请生成一个简单的重写任务"
        
        print("发送测试请求...")
        response = llm_client.call_with_json_retry(
            ModelType.CLAUDE37,
            system_prompt,
            user_prompt,
            max_json_retries=2
        )
        
        if response.success:
            try:
                parsed = json.loads(response.content)
                print("✅ JSON解析成功!")
                print(f"生成的任务数量: {len(parsed.get('tasks', []))}")
                return True
            except json.JSONDecodeError as e:
                print(f"❌ JSON解析仍然失败: {e}")
                print(f"响应内容: {response.content}")
                return False
        else:
            print(f"❌ LLM调用失败: {response.error}")
            return False
            
    except Exception as e:
        print(f"❌ 测试过程中发生异常: {e}")
        return False


if __name__ == "__main__":
    try:
        print("开始JSON修复功能测试...\n")
        
        # 测试修复策略
        fix_success = test_json_fixes()
        print()
        
        # 测试真实LLM调用
        llm_success = test_real_llm_call()
        print()
        
        if fix_success and llm_success:
            print("🎉 所有JSON测试通过！")
            print("   JSON解析问题已修复。")
            sys.exit(0)
        else:
            print("❌ 部分JSON测试失败。")
            sys.exit(1)
            
    except KeyboardInterrupt:
        print("\n测试被用户中断")
        sys.exit(1)
    except Exception as e:
        print(f"测试过程中发生异常: {e}")
        sys.exit(1)