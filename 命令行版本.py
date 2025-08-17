#!/usr/bin/env python3
"""
AutoWord vNext 命令行版本
最简单的使用方式，无需GUI依赖
"""

import sys
import os
import json
import time
from pathlib import Path
from typing import Optional

# 添加项目路径
sys.path.insert(0, str(Path(__file__).parent))

def print_banner():
    """打印横幅"""
    print("=" * 60)
    print("🎯 AutoWord vNext - 命令行版本")
    print("智能文档处理系统")
    print("=" * 60)
    print()

def get_config():
    """获取或创建配置"""
    config_file = Path("vnext_config.json")
    
    if config_file.exists():
        try:
            with open(config_file, 'r', encoding='utf-8') as f:
                config = json.load(f)
            print("✅ 已加载现有配置")
            return config
        except Exception as e:
            print(f"⚠️ 配置文件损坏: {e}")
    
    print("🔧 首次运行，请配置API密钥")
    print()
    
    # 选择模型
    print("选择AI模型:")
    print("1. OpenAI GPT-4")
    print("2. Anthropic Claude")
    
    while True:
        choice = input("请选择 (1-2): ").strip()
        if choice in ["1", "2"]:
            break
        print("❌ 请输入1或2")
    
    if choice == "1":
        provider = "openai"
        model = "gpt-4"
        print("\n🤖 OpenAI GPT-4 配置")
        print("获取API密钥: https://platform.openai.com/api-keys")
    else:
        provider = "anthropic"
        model = "claude-3-sonnet-20240229"
        print("\n🧠 Anthropic Claude 配置")
        print("获取API密钥: https://console.anthropic.com/")
    
    print()
    api_key = input("请输入API密钥: ").strip()
    
    if not api_key:
        print("❌ API密钥不能为空")
        return None
    
    # 创建配置
    config = {
        "llm": {
            "provider": provider,
            "model": model,
            "api_key": api_key,
            "temperature": 0.1,
            "max_tokens": 4000
        },
        "localization": {
            "language": "zh-CN",
            "style_aliases": {
                "Heading 1": "标题 1",
                "Heading 2": "标题 2",
                "Normal": "正文"
            },
            "font_fallbacks": {
                "楷体": ["楷体", "楷体_GB2312", "STKaiti"],
                "宋体": ["宋体", "SimSun", "NSimSun"]
            }
        },
        "validation": {
            "strict_mode": True,
            "rollback_on_failure": True
        },
        "audit": {
            "save_snapshots": True,
            "generate_diff_reports": True,
            "output_directory": "./audit_output"
        }
    }
    
    # 保存配置
    try:
        with open(config_file, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
        print(f"✅ 配置已保存到: {config_file}")
        return config
    except Exception as e:
        print(f"❌ 保存配置失败: {e}")
        return config

def get_input_file():
    """获取输入文件"""
    print("\n📁 选择输入文件")
    print("-" * 20)
    
    # 查找当前目录的docx文件
    docx_files = list(Path(".").glob("*.docx"))
    
    if docx_files:
        print("发现以下Word文档:")
        for i, file in enumerate(docx_files, 1):
            print(f"{i}. {file.name}")
        print(f"{len(docx_files) + 1}. 手动输入路径")
        print()
        
        while True:
            choice = input(f"请选择 (1-{len(docx_files) + 1}): ").strip()
            try:
                choice_num = int(choice)
                if 1 <= choice_num <= len(docx_files):
                    return str(docx_files[choice_num - 1])
                elif choice_num == len(docx_files) + 1:
                    break
                else:
                    print(f"❌ 请输入1-{len(docx_files) + 1}")
            except ValueError:
                print("❌ 请输入数字")
    
    # 手动输入路径
    while True:
        file_path = input("请输入Word文档路径: ").strip().strip('"')
        if not file_path:
            print("❌ 路径不能为空")
            continue
        
        path = Path(file_path)
        if not path.exists():
            print("❌ 文件不存在")
            continue
        
        if path.suffix.lower() not in ['.docx', '.doc']:
            print("❌ 请选择Word文档(.docx或.doc)")
            continue
        
        return str(path)

def get_user_intent():
    """获取用户意图"""
    print("\n💭 处理指令")
    print("-" * 15)
    
    presets = [
        ("删除摘要和参考文献", "删除摘要部分和参考文献部分，然后更新目录"),
        ("更新目录", "更新文档的目录"),
        ("标准化格式", "设置标题1为楷体12磅加粗，正文为宋体12磅，行距2倍"),
        ("删除摘要", "删除文档中的摘要部分"),
        ("自定义指令", "")
    ]
    
    print("常用指令:")
    for i, (name, _) in enumerate(presets, 1):
        print(f"{i}. {name}")
    print()
    
    while True:
        choice = input(f"请选择 (1-{len(presets)}): ").strip()
        try:
            choice_num = int(choice)
            if 1 <= choice_num <= len(presets):
                if choice_num == len(presets):  # 自定义指令
                    return input("请输入自定义指令: ").strip()
                else:
                    return presets[choice_num - 1][1]
            else:
                print(f"❌ 请输入1-{len(presets)}")
        except ValueError:
            print("❌ 请输入数字")

def process_document(config: dict, input_file: str, user_intent: str):
    """处理文档"""
    print("\n🚀 开始处理文档")
    print("=" * 30)
    
    try:
        print("📦 正在加载AutoWord...")
        from autoword.vnext import VNextPipeline
        from autoword.vnext.core import VNextConfig, LLMConfig
        
        print("⚙️ 正在配置系统...")
        llm_config = LLMConfig(**config["llm"])
        vnext_config = VNextConfig(
            llm=llm_config,
            localization=config.get("localization", {}),
            validation=config.get("validation", {}),
            audit=config.get("audit", {})
        )
        
        pipeline = VNextPipeline(vnext_config)
        
        print(f"📄 输入文件: {input_file}")
        print(f"💭 处理指令: {user_intent}")
        print()
        
        print("🔄 正在处理...")
        start_time = time.time()
        
        result = pipeline.process_document(input_file, user_intent)
        
        processing_time = time.time() - start_time
        
        print(f"⏱️ 处理耗时: {processing_time:.2f}秒")
        print()
        
        if result.status == "SUCCESS":
            print("✅ 处理成功！")
            print(f"📤 输出文件: {result.output_path}")
            if result.audit_directory:
                print(f"📊 审计目录: {result.audit_directory}")
            
            return True
            
        else:
            print(f"❌ 处理失败: {result.status}")
            if result.error:
                print(f"错误信息: {result.error}")
            if result.validation_errors:
                print("验证错误:")
                for error in result.validation_errors:
                    print(f"  - {error}")
            if result.plan_errors:
                print("计划错误:")
                for error in result.plan_errors:
                    print(f"  - {error}")
            
            return False
    
    except Exception as e:
        print(f"❌ 处理异常: {e}")
        import traceback
        traceback.print_exc()
        return False

def run_performance_test():
    """运行性能测试"""
    print("\n🧪 性能测试")
    print("-" * 15)
    
    try:
        import subprocess
        result = subprocess.run([sys.executable, "性能测试工具.py"], 
                              capture_output=False, text=True)
        return result.returncode == 0
    except Exception as e:
        print(f"❌ 启动性能测试失败: {e}")
        return False

def main_menu():
    """主菜单"""
    while True:
        print("\n📋 主菜单")
        print("-" * 10)
        print("1. 🚀 处理文档")
        print("2. 🧪 性能测试")
        print("3. ⚙️ 重新配置")
        print("4. 📚 查看文档")
        print("5. ❌ 退出")
        print()
        
        choice = input("请选择 (1-5): ").strip()
        
        if choice == "1":
            return "process"
        elif choice == "2":
            return "test"
        elif choice == "3":
            return "config"
        elif choice == "4":
            return "docs"
        elif choice == "5":
            return "exit"
        else:
            print("❌ 请输入1-5")

def main():
    """主函数"""
    print_banner()
    
    try:
        while True:
            action = main_menu()
            
            if action == "exit":
                print("\n👋 感谢使用AutoWord vNext！")
                break
            
            elif action == "process":
                # 获取配置
                config = get_config()
                if not config:
                    continue
                
                # 获取输入文件
                input_file = get_input_file()
                if not input_file:
                    continue
                
                # 获取用户意图
                user_intent = get_user_intent()
                if not user_intent:
                    continue
                
                # 处理文档
                success = process_document(config, input_file, user_intent)
                
                if success:
                    print("\n🎉 处理完成！")
                else:
                    print("\n😞 处理失败，请检查错误信息")
            
            elif action == "test":
                print("\n🧪 启动性能测试...")
                run_performance_test()
            
            elif action == "config":
                # 删除现有配置，强制重新配置
                config_file = Path("vnext_config.json")
                if config_file.exists():
                    config_file.unlink()
                print("✅ 配置已清除，下次处理时将重新配置")
            
            elif action == "docs":
                docs_dir = Path("docs")
                if docs_dir.exists():
                    print(f"\n📚 文档目录: {docs_dir.absolute()}")
                    print("主要文档:")
                    print("- USER_GUIDE.md: 用户指南")
                    print("- TECHNICAL_ARCHITECTURE.md: 技术架构")
                    print("- TROUBLESHOOTING.md: 故障排除")
                    print("- API_REFERENCE.md: API参考")
                else:
                    print("❌ 文档目录不存在")
            
            input("\n按回车键继续...")
    
    except KeyboardInterrupt:
        print("\n\n⏹️ 程序被用户中断")
    
    except Exception as e:
        print(f"\n❌ 程序异常: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()