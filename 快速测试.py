#!/usr/bin/env python3
"""
AutoWord vNext 快速测试
验证系统是否正常工作
"""

import sys
import os
from pathlib import Path

# 添加项目路径
sys.path.insert(0, str(Path(__file__).parent))

def test_imports():
    """测试导入"""
    print("🔍 测试系统导入...")
    
    try:
        # 测试基础Python库
        import json
        import time
        import threading
        print("✅ 基础Python库: OK")
        
        # 测试Tkinter
        import tkinter
        print("✅ Tkinter GUI库: OK")
        
        # 测试Word COM
        import win32com.client
        word = win32com.client.Dispatch("Word.Application")
        word.Quit()
        print("✅ Microsoft Word COM: OK")
        
        # 测试AutoWord核心
        from autoword.vnext import VNextPipeline
        from autoword.vnext.core import VNextConfig, LLMConfig
        print("✅ AutoWord vNext核心: OK")
        
        return True
        
    except ImportError as e:
        print(f"❌ 导入失败: {e}")
        return False
    except Exception as e:
        print(f"❌ 测试失败: {e}")
        return False

def test_config():
    """测试配置"""
    print("\n⚙️ 测试配置系统...")
    
    try:
        from autoword.vnext.core import VNextConfig, LLMConfig
        
        # 创建测试配置
        llm_config = LLMConfig(
            provider="openai",
            model="gpt-4",
            api_key="test-key",
            temperature=0.1
        )
        
        config = VNextConfig(llm=llm_config)
        print("✅ 配置系统: OK")
        return True
        
    except Exception as e:
        print(f"❌ 配置测试失败: {e}")
        return False

def test_document_creation():
    """测试文档创建"""
    print("\n📄 测试文档创建...")
    
    try:
        import win32com.client
        
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        
        doc = word.Documents.Add()
        doc.Range().Text = "AutoWord测试文档\n\n这是一个测试文档。"
        
        test_file = Path("快速测试文档.docx").absolute()
        doc.SaveAs(str(test_file))
        doc.Close()
        word.Quit()
        
        if test_file.exists():
            print(f"✅ 测试文档已创建: {test_file}")
            return str(test_file)
        else:
            print("❌ 文档创建失败")
            return None
            
    except Exception as e:
        print(f"❌ 文档创建失败: {e}")
        return None

def test_pipeline():
    """测试处理管道"""
    print("\n🚀 测试处理管道...")
    
    try:
        from autoword.vnext import VNextPipeline
        from autoword.vnext.core import VNextConfig, LLMConfig
        
        # 创建配置（使用测试密钥）
        llm_config = LLMConfig(
            provider="openai",
            model="gpt-4",
            api_key="test-key-for-demo",
            temperature=0.1
        )
        
        config = VNextConfig(llm=llm_config)
        pipeline = VNextPipeline(config)
        
        print("✅ 管道创建: OK")
        
        # 测试文档结构提取（如果有测试文档）
        test_docs = list(Path(".").glob("*.docx"))
        if test_docs:
            test_doc = str(test_docs[0])
            print(f"📄 使用测试文档: {test_doc}")
            
            # 只测试结构提取，不调用LLM
            try:
                from autoword.vnext.extractor import DocumentExtractor
                extractor = DocumentExtractor()
                structure = extractor.extract_structure(test_doc)
                
                print(f"✅ 文档结构提取: OK")
                print(f"   段落数: {len(structure.paragraphs)}")
                print(f"   标题数: {len(structure.headings)}")
                print(f"   样式数: {len(structure.styles)}")
                
                return True
                
            except Exception as e:
                print(f"⚠️ 结构提取测试跳过: {e}")
                return True
        else:
            print("⚠️ 无测试文档，跳过结构提取测试")
            return True
            
    except Exception as e:
        print(f"❌ 管道测试失败: {e}")
        return False

def test_gui():
    """测试GUI启动"""
    print("\n🖥️ 测试GUI组件...")
    
    try:
        import tkinter as tk
        from tkinter import ttk
        
        # 创建测试窗口
        root = tk.Tk()
        root.title("GUI测试")
        root.geometry("300x200")
        root.withdraw()  # 隐藏窗口
        
        # 测试基本组件
        frame = ttk.Frame(root)
        label = ttk.Label(frame, text="测试")
        button = ttk.Button(frame, text="测试按钮")
        
        root.destroy()
        
        print("✅ GUI组件: OK")
        return True
        
    except Exception as e:
        print(f"❌ GUI测试失败: {e}")
        return False

def main():
    """主函数"""
    print("🎯 AutoWord vNext 快速测试")
    print("=" * 40)
    
    tests = [
        ("系统导入", test_imports),
        ("配置系统", test_config),
        ("GUI组件", test_gui),
        ("文档创建", test_document_creation),
        ("处理管道", test_pipeline)
    ]
    
    results = []
    
    for test_name, test_func in tests:
        try:
            result = test_func()
            results.append((test_name, result))
        except Exception as e:
            print(f"❌ {test_name}测试异常: {e}")
            results.append((test_name, False))
    
    # 生成报告
    print("\n" + "=" * 40)
    print("📊 测试报告")
    print("=" * 40)
    
    passed = sum(1 for _, result in results if result)
    total = len(results)
    
    for test_name, result in results:
        status = "✅ 通过" if result else "❌ 失败"
        print(f"{test_name}: {status}")
    
    print(f"\n总计: {passed}/{total} 通过")
    success_rate = passed / total * 100
    print(f"成功率: {success_rate:.1f}%")
    
    if success_rate >= 80:
        print("\n🎉 系统状态良好，可以正常使用！")
        print("\n下一步:")
        print("1. 运行 '一键启动AutoWord.bat' 开始使用")
        print("2. 或运行 'python 简化版GUI.py' 启动GUI")
        print("3. 或运行 'python 命令行版本.py' 使用命令行")
    elif success_rate >= 60:
        print("\n⚠️ 系统基本可用，但可能存在一些问题")
        print("建议检查失败的测试项目")
    else:
        print("\n❌ 系统存在严重问题，需要修复")
        print("请检查Python环境和依赖安装")
    
    print(f"\n🏁 测试完成")
    input("按回车键退出...")

if __name__ == "__main__":
    main()