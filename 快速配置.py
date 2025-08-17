#!/usr/bin/env python3
"""
AutoWord vNext 快速配置向导
帮助用户快速设置API密钥和基本配置
"""

import sys
import os
import json
from pathlib import Path
from getpass import getpass

def print_banner():
    """打印欢迎横幅"""
    print("🎯 AutoWord vNext 快速配置向导")
    print("="*50)
    print("欢迎使用AutoWord！让我们快速设置您的配置。")
    print()

def get_api_key_info():
    """获取API密钥信息"""
    print("🔑 API密钥配置")
    print("-" * 30)
    print()
    print("AutoWord支持以下AI模型:")
    print("1. OpenAI GPT-4 (推荐)")
    print("2. Anthropic Claude")
    print("3. 跳过配置 (稍后在GUI中设置)")
    print()
    
    while True:
        choice = input("请选择模型 (1-3): ").strip()
        
        if choice == "1":
            return setup_openai()
        elif choice == "2":
            return setup_claude()
        elif choice == "3":
            print("✅ 跳过API配置，您可以稍后在GUI中设置")
            return None
        else:
            print("❌ 无效选择，请输入1-3")

def setup_openai():
    """设置OpenAI配置"""
    print("\n🤖 OpenAI GPT-4 配置")
    print("-" * 25)
    print()
    print("📝 如何获取OpenAI API密钥:")
    print("1. 访问 https://platform.openai.com/")
    print("2. 注册/登录账户")
    print("3. 进入 API Keys 页面")
    print("4. 创建新的API密钥")
    print()
    
    api_key = getpass("请输入您的OpenAI API密钥 (输入时不显示): ").strip()
    
    if not api_key:
        print("❌ API密钥不能为空")
        return None
    
    if not api_key.startswith("sk-"):
        print("⚠️  警告: OpenAI API密钥通常以'sk-'开头")
        confirm = input("确认继续? (y/N): ").strip().lower()
        if confirm != 'y':
            return None
    
    return {
        "provider": "openai",
        "model": "gpt-4",
        "api_key": api_key,
        "temperature": 0.1,
        "max_tokens": 4000
    }

def setup_claude():
    """设置Claude配置"""
    print("\n🧠 Anthropic Claude 配置")
    print("-" * 25)
    print()
    print("📝 如何获取Claude API密钥:")
    print("1. 访问 https://console.anthropic.com/")
    print("2. 注册/登录账户")
    print("3. 进入 API Keys 页面")
    print("4. 创建新的API密钥")
    print()
    
    api_key = getpass("请输入您的Claude API密钥 (输入时不显示): ").strip()
    
    if not api_key:
        print("❌ API密钥不能为空")
        return None
    
    if not api_key.startswith("sk-ant-"):
        print("⚠️  警告: Claude API密钥通常以'sk-ant-'开头")
        confirm = input("确认继续? (y/N): ").strip().lower()
        if confirm != 'y':
            return None
    
    return {
        "provider": "anthropic", 
        "model": "claude-3-sonnet-20240229",
        "api_key": api_key,
        "temperature": 0.1,
        "max_tokens": 4000
    }

def setup_localization():
    """设置本地化配置"""
    print("\n🌏 本地化设置")
    print("-" * 20)
    print()
    print("选择文档语言:")
    print("1. 中文 (推荐)")
    print("2. 英文")
    print("3. 自动检测")
    print()
    
    while True:
        choice = input("请选择 (1-3): ").strip()
        
        if choice == "1":
            return {
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
            }
        elif choice == "2":
            return {
                "language": "en-US",
                "style_aliases": {},
                "font_fallbacks": {
                    "Times New Roman": ["Times New Roman", "Times", "serif"],
                    "Arial": ["Arial", "Helvetica", "sans-serif"]
                }
            }
        elif choice == "3":
            return {
                "language": "auto",
                "style_aliases": {},
                "font_fallbacks": {}
            }
        else:
            print("❌ 无效选择，请输入1-3")

def setup_validation():
    """设置验证配置"""
    print("\n✅ 验证设置")
    print("-" * 15)
    print()
    print("选择验证模式:")
    print("1. 严格模式 (推荐) - 更安全，会回滚失败的操作")
    print("2. 宽松模式 - 更快速，但可能保留部分失败的修改")
    print()
    
    while True:
        choice = input("请选择 (1-2): ").strip()
        
        if choice == "1":
            return {
                "strict_mode": True,
                "rollback_on_failure": True,
                "chapter_assertions": True,
                "style_assertions": True,
                "toc_assertions": True
            }
        elif choice == "2":
            return {
                "strict_mode": False,
                "rollback_on_failure": False,
                "chapter_assertions": True,
                "style_assertions": False,
                "toc_assertions": True
            }
        else:
            print("❌ 无效选择，请输入1-2")

def save_configuration(config):
    """保存配置到文件"""
    config_file = Path("vnext_config.json")
    
    try:
        with open(config_file, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
        
        print(f"✅ 配置已保存到: {config_file.absolute()}")
        return True
        
    except Exception as e:
        print(f"❌ 保存配置失败: {e}")
        return False

def test_configuration(config):
    """测试配置"""
    print("\n🧪 测试配置...")
    
    try:
        # 测试导入
        sys.path.insert(0, str(Path(__file__).parent))
        from autoword.vnext.core import VNextConfig, LLMConfig
        
        # 创建配置对象
        if config.get("llm"):
            llm_config = LLMConfig(**config["llm"])
            vnext_config = VNextConfig(llm=llm_config)
            print("✅ 配置格式正确")
        else:
            print("⚠️  跳过LLM配置测试")
        
        return True
        
    except Exception as e:
        print(f"❌ 配置测试失败: {e}")
        return False

def main():
    """主函数"""
    print_banner()
    
    try:
        # 收集配置
        config = {}
        
        # API密钥配置
        llm_config = get_api_key_info()
        if llm_config:
            config["llm"] = llm_config
        
        # 本地化配置
        localization_config = setup_localization()
        config["localization"] = localization_config
        
        # 验证配置
        validation_config = setup_validation()
        config["validation"] = validation_config
        
        # 审计配置
        config["audit"] = {
            "save_snapshots": True,
            "generate_diff_reports": True,
            "output_directory": "./audit_output"
        }
        
        # 显示配置摘要
        print("\n📋 配置摘要")
        print("-" * 20)
        
        if "llm" in config:
            print(f"🤖 AI模型: {config['llm']['provider']} - {config['llm']['model']}")
        else:
            print("🤖 AI模型: 未配置 (稍后在GUI中设置)")
        
        print(f"🌏 语言: {config['localization']['language']}")
        print(f"✅ 验证模式: {'严格' if config['validation']['strict_mode'] else '宽松'}")
        print()
        
        # 确认保存
        save_confirm = input("保存此配置? (Y/n): ").strip().lower()
        if save_confirm in ['', 'y', 'yes']:
            if save_configuration(config):
                # 测试配置
                test_configuration(config)
                
                print("\n🎉 配置完成！")
                print()
                print("下一步:")
                print("1. 运行 '启动AutoWord测试.bat' 启动GUI")
                print("2. 或运行 'python 性能测试工具.py' 进行性能测试")
                print("3. 查看 'docs/USER_GUIDE.md' 了解详细使用方法")
            else:
                print("❌ 配置保存失败")
        else:
            print("❌ 配置未保存")
    
    except KeyboardInterrupt:
        print("\n\n⏹️  配置被用户中断")
    
    except Exception as e:
        print(f"\n❌ 配置过程出错: {e}")
        import traceback
        traceback.print_exc()
    
    print(f"\n🏁 配置向导结束")
    input("按回车键退出...")

if __name__ == "__main__":
    main()