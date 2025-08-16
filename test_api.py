#!/usr/bin/env python3
"""
测试API连接
"""

import sys
sys.path.append('.')

from autoword.core.llm_client import LLMClient, ModelType

def test_api():
    print("测试API连接...")
    
    client = LLMClient()
    
    # 测试GPT5
    print("\n测试GPT5...")
    response = client.call_gpt5(
        system_prompt="你是一个有用的助手。",
        user_prompt="请回答：1+1等于多少？"
    )
    
    print(f"GPT5 响应:")
    print(f"  成功: {response.success}")
    print(f"  内容: {response.content[:100]}...")
    print(f"  错误: {response.error}")
    
    # 测试Claude 3.7
    print("\n测试Claude 3.7...")
    response = client.call_claude37(
        system_prompt="你是一个有用的助手。",
        user_prompt="请回答：1+1等于多少？"
    )
    
    print(f"Claude 3.7 响应:")
    print(f"  成功: {response.success}")
    print(f"  内容: {response.content[:100]}...")
    print(f"  错误: {response.error}")

if __name__ == "__main__":
    test_api()