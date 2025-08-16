#!/usr/bin/env python3
"""
调试JSON解析问题
"""

import sys
import json
from pathlib import Path

# 添加项目根目录到Python路径
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

def debug_json_issue():
    """调试JSON解析问题"""
    print("========================================")
    print("      调试JSON解析问题")
    print("========================================")
    print()
    
    from autoword.core.llm_client import LLMClient, ModelType
    
    # 配置API密钥
    api_keys = {
        "claude": "sk-3w1JFbWUq7tKjpLlopdkISQ9F6fpLhHx5viD0frh43ESE9Io",
        "gpt": "sk-NhjnJtqlZMx4PGTqvkGlH4POT82HHBrBnBbWOat99Bs5VZXi"
    }
    
    llm_client = LLMClient(api_keys=api_keys)
    
    # 使用更严格的系统提示词
    system_prompt = """你是一个任务规划助手。请根据用户的要求生成JSON格式的任务列表。

返回格式必须是严格的JSON，示例：
{
  "tasks": [
    {
      "id": "task_1",
      "type": "rewrite",
      "instruction": "重写内容",
      "locator": {"by": "find", "value": "目标文本"},
      "source_comment_id": "comment_1"
    }
  ]
}

严格要求：
1. 只返回JSON，不要任何解释文字
2. 不要使用markdown标记
3. 确保所有字符串都用双引号
4. 确保JSON格式完全正确
5. 不要在最后一个元素后加逗号"""
    
    user_prompt = "请生成一个简单的重写任务，用于处理文档批注"
    
    print("发送调试请求...")
    print("系统提示词长度:", len(system_prompt))
    print()
    
    # 直接调用模型，不使用重试机制，看看原始响应
    response = llm_client.call_model(ModelType.CLAUDE37, system_prompt, user_prompt)
    
    if response.success:
        print("原始响应:")
        print("=" * 50)
        print(repr(response.content))
        print("=" * 50)
        print()
        
        print("原始响应内容:")
        print(response.content)
        print()
        
        # 尝试直接解析
        try:
            parsed = json.loads(response.content)
            print("✅ 原始JSON解析成功!")
            return True
        except json.JSONDecodeError as e:
            print(f"❌ 原始JSON解析失败: {e}")
            print(f"错误位置: 第{e.lineno}行，第{e.colno}列")
            print()
            
            # 显示错误位置的内容
            lines = response.content.split('\n')
            if e.lineno <= len(lines):
                error_line = lines[e.lineno - 1]
                print(f"错误行内容: {repr(error_line)}")
                if e.colno <= len(error_line):
                    print(f"错误字符: {repr(error_line[e.colno-1:e.colno+5])}")
            print()
            
            # 尝试清理修复
            print("尝试清理修复...")
            cleaned = llm_client._clean_json_content(response.content)
            print("清理后内容:")
            print("=" * 50)
            print(cleaned)
            print("=" * 50)
            print()
            
            try:
                parsed = json.loads(cleaned)
                print("✅ 清理后JSON解析成功!")
                return True
            except json.JSONDecodeError as e2:
                print(f"❌ 清理后仍然失败: {e2}")
                
                # 尝试激进修复
                print("尝试激进修复...")
                fixed = llm_client._aggressive_json_fix(response.content)
                print("激进修复后内容:")
                print("=" * 50)
                print(fixed)
                print("=" * 50)
                print()
                
                try:
                    parsed = json.loads(fixed)
                    print("✅ 激进修复后JSON解析成功!")
                    return True
                except json.JSONDecodeError as e3:
                    print(f"❌ 激进修复后仍然失败: {e3}")
                    
                    # 手动修复
                    print("尝试手动修复...")
                    manual_fixed = manual_json_fix(response.content)
                    print("手动修复后内容:")
                    print("=" * 50)
                    print(manual_fixed)
                    print("=" * 50)
                    print()
                    
                    try:
                        parsed = json.loads(manual_fixed)
                        print("✅ 手动修复后JSON解析成功!")
                        return True
                    except json.JSONDecodeError as e4:
                        print(f"❌ 手动修复后仍然失败: {e4}")
                        return False
    else:
        print(f"❌ LLM调用失败: {response.error}")
        return False


def manual_json_fix(content: str) -> str:
    """手动JSON修复"""
    import re
    
    content = content.strip()
    
    # 移除所有markdown标记
    content = re.sub(r'```[a-zA-Z]*\n?', '', content)
    content = re.sub(r'```', '', content)
    
    # 移除前后的非JSON文本
    start_pos = content.find('{')
    if start_pos != -1:
        content = content[start_pos:]
    
    end_pos = content.rfind('}')
    if end_pos != -1:
        content = content[:end_pos + 1]
    
    # 修复常见问题
    content = re.sub(r',(\s*[}\]])', r'\1', content)  # 移除尾随逗号
    content = re.sub(r'([}\]])(\s*)([{\[])', r'\1,\2\3', content)  # 添加缺失的逗号
    
    # 修复特定的逗号问题
    content = re.sub(r'"\s*\n\s*"', '",\n"', content)
    content = re.sub(r'}\s*\n\s*"', '},\n"', content)
    content = re.sub(r'"\s*\n\s*{', '",\n{', content)
    
    # 修复未引用的键名
    content = re.sub(r'([{,]\s*)([a-zA-Z_][a-zA-Z0-9_]*)\s*:', r'\1"\2":', content)
    
    return content.strip()


if __name__ == "__main__":
    try:
        success = debug_json_issue()
        if success:
            print()
            print("🎉 JSON解析问题已解决!")
            sys.exit(0)
        else:
            print()
            print("❌ JSON解析问题仍然存在。")
            sys.exit(1)
    except KeyboardInterrupt:
        print("\n调试被用户中断")
        sys.exit(1)
    except Exception as e:
        print(f"调试过程中发生异常: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)