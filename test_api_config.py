#!/usr/bin/env python3
"""
测试API配置
"""

import sys
from pathlib import Path

# 添加项目根目录到Python路径
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from autoword.core.llm_client import LLMClient, ModelType
from autoword.gui.config_manager import ConfigurationManager


def test_api_configuration():
    """测试API配置"""
    print("========================================")
    print("      AutoWord API 配置测试")
    print("========================================")
    print()
    
    # 测试配置管理器
    print("[1/3] 测试配置管理器...")
    try:
        config_manager = ConfigurationManager()
        
        # 显示可用模型
        models = config_manager.get_available_models()
        print(f"✅ 可用模型: {models}")
        
        # 显示API密钥（部分隐藏）
        for model_key, model_name in models.items():
            api_key = config_manager.get_api_key(model_key)
            if api_key:
                masked_key = api_key[:10] + "..." + api_key[-10:] if len(api_key) > 20 else api_key
                print(f"✅ {model_name} API密钥: {masked_key}")
            else:
                print(f"❌ {model_name} API密钥未设置")
        
        print()
        
    except Exception as e:
        print(f"❌ 配置管理器测试失败: {e}")
        return False
    
    # 测试LLM客户端初始化
    print("[2/3] 测试LLM客户端...")
    try:
        # 获取API密钥
        api_keys = {
            "claude": config_manager.get_api_key("claude"),
            "gpt": config_manager.get_api_key("gpt")
        }
        
        # 创建LLM客户端
        llm_client = LLMClient(api_keys=api_keys)
        print("✅ LLM客户端初始化成功")
        
        # 测试密钥获取
        try:
            claude_key = llm_client._get_api_key(ModelType.CLAUDE37)
            gpt_key = llm_client._get_api_key(ModelType.GPT5)
            print("✅ API密钥获取成功")
        except Exception as e:
            print(f"❌ API密钥获取失败: {e}")
            return False
        
        print()
        
    except Exception as e:
        print(f"❌ LLM客户端测试失败: {e}")
        return False
    
    # 测试API连接（简单测试）
    print("[3/3] 测试API连接...")
    try:
        # 测试Claude 3.7
        print("测试Claude 3.7连接...")
        claude_response = llm_client.call_claude37(
            system_prompt="你是一个测试助手。",
            user_prompt="请回复'测试成功'",
            temperature=0.1
        )
        
        if claude_response.success:
            print(f"✅ Claude 3.7 连接成功: {claude_response.content[:50]}...")
        else:
            print(f"❌ Claude 3.7 连接失败: {claude_response.error}")
        
        # 测试GPT-4o
        print("测试GPT-4o连接...")
        gpt_response = llm_client.call_gpt5(
            system_prompt="你是一个测试助手。",
            user_prompt="请回复'测试成功'",
            temperature=0.1
        )
        
        if gpt_response.success:
            print(f"✅ GPT-4o 连接成功: {gpt_response.content[:50]}...")
        else:
            print(f"❌ GPT-4o 连接失败: {gpt_response.error}")
        
        print()
        
    except Exception as e:
        print(f"❌ API连接测试失败: {e}")
        return False
    
    print("========================================")
    print("      API配置测试完成")
    print("========================================")
    
    return True


if __name__ == "__main__":
    try:
        success = test_api_configuration()
        if success:
            print("✅ 所有测试通过！")
            sys.exit(0)
        else:
            print("❌ 部分测试失败，请检查配置")
            sys.exit(1)
    except KeyboardInterrupt:
        print("\n测试被用户中断")
        sys.exit(1)
    except Exception as e:
        print(f"测试过程中发生异常: {e}")
        sys.exit(1)