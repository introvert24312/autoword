#!/usr/bin/env python3
"""
AutoWord vNext 预配置启动器
已预配置你的API密钥，可以直接使用
"""

import sys
import os
import json
import time
from pathlib import Path
from typing import Optional

# 添加项目路径
sys.path.insert(0, str(Path(__file__).parent))

def create_preconfigured_config():
    """创建预配置的配置文件"""
    config = {
        "llm": {
            "provider": "custom",
            "model": "claude-3-7-sonnet-20250219",
            "api_key": "sk-3w1JFbWUq7tKjpLlopdkISQ9F6fpLhHx5viD0frh43ESE9Io",
            "base_url": "globalai.vip",
            "temperature": 0.1,
            "max_tokens": 4000
        },
        "gpt_config": {
            "provider": "custom",
            "model": "gpt-4",
            "api_key": "sk-NhjnJtqlZMx4PGTqvkGlH4POT82HHBrBnBbWOat99Bs5VZXi",
            "base_url": "globalai.vip",
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
        }
    }
    
    # 保存配置
    with open("vnext_config.json", 'w', encoding='utf-8') as f:
        json.dump(config, f, indent=2, ensure_ascii=False)
    
    print("✅ 预配置文件已创建")
    return config

def test_api_connection(api_key: str, model: str):
    """测试API连接"""
    try:
        import http.client
        
        conn = http.client.HTTPSConnection("globalai.vip")
        
        payload = json.dumps({
            "model": model,
            "messages": [
                {
                    "role": "system",
                    "content": "You are a helpful assistant."
                },
                {
                    "role": "user",
                    "content": "Hello, please respond with 'API connection successful'"
                }
            ],
            "max_tokens": 50
        })
        
        headers = {
            'Accept': 'application/json',
            'Authorization': f'Bearer {api_key}',
            'Content-Type': 'application/json',
            'Host': 'globalai.vip',
            'Connection': 'keep-alive'
        }
        
        conn.request("POST", "/v1/chat/completions", payload, headers)
        res = conn.getresponse()
        data = res.read()
        
        response_json = json.loads(data.decode("utf-8"))
        
        if 'choices' in response_json and len(response_json['choices']) > 0:
            content = response_json['choices'][0]['message']['content']
            print(f"✅ API测试成功: {content}")
            return True
        else:
            print(f"❌ API响应格式错误: {response_json}")
            return False
            
    except Exception as e:
        print(f"❌ API测试失败: {e}")
        return False

def quick_document_test():
    """快速文档测试"""
    print("\n🧪 快速文档处理测试")
    print("-" * 30)
    
    try:
        # 导入AutoWord
        from autoword.vnext import VNextPipeline, VNextConfig, LLMConfig
        
        # 创建配置
        config = VNextConfig()
        config.llm = LLMConfig(
            provider="custom",
            model="claude-3-7-sonnet-20250219",
            api_key="sk-3w1JFbWUq7tKjpLlopdkISQ9F6fpLhHx5viD0frh43ESE9Io",
            base_url="globalai.vip"
        )
        
        # 创建管道
        pipeline = VNextPipeline(config)
        
        # 查找测试文档
        docx_files = list(Path(".").glob("*.docx"))
        
        if docx_files:
            test_doc = str(docx_files[0])
            print(f"📄 使用测试文档: {test_doc}")
            
            # 测试处理
            user_intent = "更新文档目录"
            print(f"💭 测试意图: {user_intent}")
            
            start_time = time.time()
            result = pipeline.process_document(test_doc, user_intent)
            processing_time = time.time() - start_time
            
            print(f"⏱️ 处理耗时: {processing_time:.2f}秒")
            print(f"📊 处理状态: {result.status}")
            
            if result.status == "SUCCESS":
                print("🎉 测试成功！")
                if hasattr(result, 'audit_directory') and result.audit_directory:
                    print(f"📁 输出目录: {result.audit_directory}")
                return True
            else:
                print(f"⚠️ 测试完成但有问题: {result.message}")
                if hasattr(result, 'errors') and result.errors:
                    for error in result.errors:
                        print(f"   错误: {error}")
                return False
        else:
            print("⚠️ 未找到测试文档，创建一个简单的测试文档...")
            return create_test_document_and_process(pipeline)
            
    except Exception as e:
        print(f"❌ 测试失败: {e}")
        import traceback
        traceback.print_exc()
        return False

def create_test_document_and_process(pipeline):
    """创建测试文档并处理"""
    try:
        import win32com.client
        
        print("📄 创建测试文档...")
        
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        
        doc = word.Documents.Add()
        
        # 添加内容
        doc.Range().Text = """AutoWord vNext 测试文档

摘要
这是一个测试文档，用于验证AutoWord vNext的功能。

目录
1. 引言
2. 方法
3. 结果
4. 结论

引言
本文档用于测试AutoWord vNext系统的文档处理能力。

方法
我们使用自动化测试来验证系统功能。

结果
测试结果将显示系统是否正常工作。

结论
AutoWord vNext是一个强大的文档处理工具。

参考文献
1. AutoWord技术文档
2. Microsoft Word自动化指南
"""
        
        # 保存文档
        test_doc_path = Path("AutoWord测试文档.docx").absolute()
        doc.SaveAs(str(test_doc_path))
        doc.Close()
        word.Quit()
        
        print(f"✅ 测试文档已创建: {test_doc_path}")
        
        # 处理文档
        user_intent = "删除摘要部分和参考文献部分，然后更新目录"
        print(f"💭 处理意图: {user_intent}")
        
        start_time = time.time()
        result = pipeline.process_document(str(test_doc_path), user_intent)
        processing_time = time.time() - start_time
        
        print(f"⏱️ 处理耗时: {processing_time:.2f}秒")
        print(f"📊 处理状态: {result.status}")
        
        if result.status == "SUCCESS":
            print("🎉 文档处理测试成功！")
            return True
        else:
            print(f"⚠️ 处理有问题: {result.message}")
            return False
            
    except Exception as e:
        print(f"❌ 测试文档创建失败: {e}")
        return False

def main():
    """主函数"""
    print("🎯 AutoWord vNext 预配置启动器")
    print("=" * 50)
    print("已预配置你的API密钥，可以直接使用！")
    print()
    
    try:
        # 1. 创建预配置
        print("⚙️ 创建预配置...")
        config = create_preconfigured_config()
        
        # 2. 测试API连接
        print("\n🔗 测试API连接...")
        claude_success = test_api_connection(
            config["llm"]["api_key"], 
            config["llm"]["model"]
        )
        
        gpt_success = test_api_connection(
            config["gpt_config"]["api_key"],
            config["gpt_config"]["model"]
        )
        
        if not (claude_success or gpt_success):
            print("❌ 所有API连接都失败了")
            print("请检查网络连接和API密钥")
            return
        
        # 3. 快速文档测试
        print("\n📄 开始文档处理测试...")
        success = quick_document_test()
        
        if success:
            print("\n🎉 所有测试通过！")
            print("\n下一步:")
            print("1. 运行 'python 简化版GUI.py' 启动图形界面")
            print("2. 运行 'python 命令行版本.py' 使用命令行")
            print("3. 双击 '一键启动AutoWord.bat' 选择启动方式")
        else:
            print("\n⚠️ 测试有问题，但基本功能可用")
            print("可以尝试使用GUI或命令行版本")
        
    except Exception as e:
        print(f"\n❌ 启动器异常: {e}")
        import traceback
        traceback.print_exc()
    
    print(f"\n🏁 预配置启动器完成")
    input("按回车键退出...")

if __name__ == "__main__":
    main()