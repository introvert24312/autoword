#!/usr/bin/env python3
"""
AutoWord vNext API密钥快速配置
直接配置你提供的API密钥
"""

import json
import os
from pathlib import Path

def create_config_with_keys():
    """使用提供的API密钥创建配置"""
    
    print("🔑 AutoWord vNext API密钥配置")
    print("=" * 40)
    print()
    
    # 你提供的API密钥
    openai_key = "sk-NhjnJtqlZMx4PGTqvkGlH4POT82HHBrBnBbWOat99Bs5VZXi"
    claude_key = "sk-3w1JFbWUq7tKjpLlopdkISQ9F6fpLhHx5viD0frh43ESE9Io"
    
    print("检测到以下API密钥:")
    print(f"🤖 OpenAI GPT: {openai_key[:20]}...")
    print(f"🧠 Claude: {claude_key[:20]}...")
    print()
    
    # 创建完整配置
    config = {
        "llm": {
            "provider": "openai",  # 默认使用OpenAI
            "model": "gpt-4",
            "api_key": openai_key,
            "temperature": 0.1,
            "max_tokens": 4000
        },
        "llm_backup": {
            "provider": "anthropic",
            "model": "claude-3-sonnet-20240229", 
            "api_key": claude_key,
            "temperature": 0.1,
            "max_tokens": 4000
        },
        "localization": {
            "language": "zh-CN",
            "style_aliases": {
                "Heading 1": "标题 1",
                "Heading 2": "标题 2", 
                "Heading 3": "标题 3",
                "Normal": "正文",
                "Title": "标题"
            },
            "font_fallbacks": {
                "楷体": ["楷体", "楷体_GB2312", "STKaiti"],
                "宋体": ["宋体", "SimSun", "NSimSun"],
                "黑体": ["黑体", "SimHei", "Microsoft YaHei"]
            }
        },
        "validation": {
            "strict_mode": True,
            "rollback_on_failure": True,
            "chapter_assertions": True,
            "style_assertions": True,
            "toc_assertions": True
        },
        "audit": {
            "save_snapshots": True,
            "generate_diff_reports": True,
            "output_directory": "./audit_output"
        },
        "performance": {
            "visible_word": False,
            "enable_monitoring": True,
            "memory_limit_mb": 2048
        }
    }
    
    # 保存主配置文件
    config_file = Path("vnext_config.json")
    try:
        with open(config_file, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
        print(f"✅ 主配置已保存: {config_file.absolute()}")
    except Exception as e:
        print(f"❌ 保存主配置失败: {e}")
        return False
    
    # 设置环境变量
    try:
        os.environ["OPENAI_API_KEY"] = openai_key
        os.environ["ANTHROPIC_API_KEY"] = claude_key
        print("✅ 环境变量已设置")
    except Exception as e:
        print(f"⚠️ 环境变量设置失败: {e}")
    
    # 创建简化的启动配置
    simple_config = {
        "openai_key": openai_key,
        "claude_key": claude_key,
        "default_model": "gpt-4",
        "language": "zh-CN"
    }
    
    simple_config_file = Path("simple_config.json")
    try:
        with open(simple_config_file, 'w', encoding='utf-8') as f:
            json.dump(simple_config, f, indent=2, ensure_ascii=False)
        print(f"✅ 简化配置已保存: {simple_config_file.absolute()}")
    except Exception as e:
        print(f"❌ 保存简化配置失败: {e}")
    
    return True

def test_api_keys():
    """测试API密钥"""
    print("\n🧪 测试API密钥...")
    
    try:
        # 测试OpenAI
        print("测试OpenAI连接...")
        import openai
        openai.api_key = "sk-NhjnJtqlZMx4PGTqvkGlH4POT82HHBrBnBbWOat99Bs5VZXi"
        
        # 简单测试请求
        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": "Hello"}],
            max_tokens=5
        )
        print("✅ OpenAI API连接成功")
        
    except Exception as e:
        print(f"⚠️ OpenAI API测试失败: {e}")
    
    try:
        # 测试Claude (需要anthropic库)
        print("测试Claude连接...")
        import anthropic
        
        client = anthropic.Anthropic(
            api_key="sk-3w1JFbWUq7tKjpLlopdkISQ9F6fpLhHx5viD0frh43ESE9Io"
        )
        
        message = client.messages.create(
            model="claude-3-sonnet-20240229",
            max_tokens=5,
            messages=[{"role": "user", "content": "Hello"}]
        )
        print("✅ Claude API连接成功")
        
    except Exception as e:
        print(f"⚠️ Claude API测试失败: {e}")

def main():
    """主函数"""
    print("🚀 开始配置API密钥...")
    print()
    
    if create_config_with_keys():
        print("\n🎉 API密钥配置完成！")
        print()
        print("配置文件:")
        print("- vnext_config.json (完整配置)")
        print("- simple_config.json (简化配置)")
        print()
        print("下一步:")
        print("1. 运行 'python 简化版GUI.py' 启动图形界面")
        print("2. 运行 'python 命令行版本.py' 启动命令行版本")
        print("3. 双击 '一键启动AutoWord.bat' 选择启动方式")
        print()
        
        # 询问是否测试API
        test_choice = input("是否测试API连接? (y/N): ").strip().lower()
        if test_choice == 'y':
            test_api_keys()
    else:
        print("❌ 配置失败")
    
    print(f"\n🏁 配置完成")
    input("按回车键退出...")

if __name__ == "__main__":
    main()