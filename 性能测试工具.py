#!/usr/bin/env python3
"""
AutoWord vNext 性能测试工具
快速测试AutoWord的各种功能和性能
"""

import sys
import time
import os
from pathlib import Path
from typing import Dict, List, Any
import json

# 添加项目路径
sys.path.insert(0, str(Path(__file__).parent))

def check_environment():
    """检查运行环境"""
    print("🔍 检查运行环境...")
    
    # 检查Python版本
    if sys.version_info < (3, 8):
        print(f"❌ Python版本过低: {sys.version}")
        print("需要Python 3.8+")
        return False
    
    print(f"✅ Python版本: {sys.version.split()[0]}")
    
    # 检查Word
    try:
        import win32com.client
        word = win32com.client.Dispatch("Word.Application")
        word.Quit()
        print("✅ Microsoft Word: 已安装")
    except Exception as e:
        print(f"❌ Microsoft Word: 未找到 ({e})")
        return False
    
    # 检查核心依赖
    try:
        from autoword.vnext import VNextPipeline
        print("✅ AutoWord vNext: 已安装")
    except ImportError as e:
        print(f"❌ AutoWord vNext: 导入失败 ({e})")
        return False
    
    return True


def create_test_document():
    """创建测试文档"""
    print("\n📄 创建测试文档...")
    
    try:
        import win32com.client
        
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        
        doc = word.Documents.Add()
        
        # 添加标题
        doc.Range().Text = "AutoWord vNext 性能测试文档\n\n"
        
        # 添加摘要
        doc.Range().InsertAfter("摘要\n")
        doc.Range().InsertAfter("这是一个用于测试AutoWord vNext性能的示例文档。本文档包含了各种常见的文档元素，用于验证系统的处理能力。\n\n")
        
        # 添加目录
        doc.Range().InsertAfter("目录\n")
        doc.Range().InsertAfter("1. 引言\n")
        doc.Range().InsertAfter("2. 方法\n") 
        doc.Range().InsertAfter("3. 结果\n")
        doc.Range().InsertAfter("4. 结论\n")
        doc.Range().InsertAfter("5. 参考文献\n\n")
        
        # 添加正文内容
        sections = [
            ("1. 引言", "本研究旨在评估AutoWord vNext系统的性能和功能。"),
            ("2. 方法", "我们采用了多种测试方法来评估系统性能。"),
            ("3. 结果", "测试结果显示系统具有良好的性能表现。"),
            ("4. 结论", "AutoWord vNext是一个高效的文档处理系统。"),
            ("参考文献", "1. AutoWord技术文档\n2. Microsoft Word自动化指南")
        ]
        
        for title, content in sections:
            doc.Range().InsertAfter(f"{title}\n")
            doc.Range().InsertAfter(f"{content}\n\n")
        
        # 保存文档
        test_doc_path = Path("测试文档.docx").absolute()
        doc.SaveAs(str(test_doc_path))
        doc.Close()
        word.Quit()
        
        print(f"✅ 测试文档已创建: {test_doc_path}")
        return str(test_doc_path)
        
    except Exception as e:
        print(f"❌ 创建测试文档失败: {e}")
        return None


def performance_test_suite():
    """性能测试套件"""
    print("\n🚀 开始性能测试...")
    
    # 检查环境
    if not check_environment():
        print("❌ 环境检查失败，无法继续测试")
        return
    
    # 创建测试文档
    test_doc = create_test_document()
    if not test_doc:
        print("❌ 无法创建测试文档，使用现有文档进行测试")
        # 寻找现有的docx文件
        docx_files = list(Path(".").glob("*.docx"))
        if docx_files:
            test_doc = str(docx_files[0])
            print(f"📄 使用现有文档: {test_doc}")
        else:
            print("❌ 未找到可用的测试文档")
            return
    
    # 导入AutoWord
    try:
        from autoword.vnext import VNextPipeline
        from autoword.vnext.core import VNextConfig, LLMConfig
    except ImportError as e:
        print(f"❌ 导入AutoWord失败: {e}")
        return
    
    # 测试用例
    test_cases = [
        {
            "name": "基础功能测试",
            "intent": "更新目录",
            "description": "测试基本的目录更新功能"
        },
        {
            "name": "删除章节测试", 
            "intent": "删除摘要部分",
            "description": "测试章节删除功能"
        },
        {
            "name": "样式设置测试",
            "intent": "设置标题1为楷体字体，12磅，加粗",
            "description": "测试样式修改功能"
        },
        {
            "name": "复合操作测试",
            "intent": "删除参考文献部分，更新目录，设置正文为宋体12磅",
            "description": "测试多个操作的组合执行"
        }
    ]
    
    results = []
    
    print(f"\n📊 开始执行 {len(test_cases)} 个测试用例...")
    
    for i, test_case in enumerate(test_cases, 1):
        print(f"\n🧪 测试 {i}/{len(test_cases)}: {test_case['name']}")
        print(f"📝 描述: {test_case['description']}")
        print(f"💭 意图: {test_case['intent']}")
        
        start_time = time.time()
        
        try:
            # 创建管道（使用模拟配置）
            config = VNextConfig(
                llm=LLMConfig(
                    provider="openai",
                    model="gpt-4",
                    api_key="test-key-for-demo"  # 演示用密钥
                )
            )
            
            pipeline = VNextPipeline(config)
            
            # 执行干运行（生成计划但不执行）
            print("🔄 生成执行计划...")
            plan_result = pipeline.generate_plan(test_doc, test_case['intent'])
            
            processing_time = time.time() - start_time
            
            if plan_result.is_valid:
                print(f"✅ 计划生成成功 ({processing_time:.2f}s)")
                print(f"📋 操作数量: {len(plan_result.plan.ops)}")
                
                for j, op in enumerate(plan_result.plan.ops, 1):
                    print(f"   {j}. {op.op_type}")
                
                results.append({
                    "test": test_case['name'],
                    "status": "SUCCESS",
                    "time": processing_time,
                    "operations": len(plan_result.plan.ops),
                    "details": [op.op_type for op in plan_result.plan.ops]
                })
            else:
                print(f"❌ 计划生成失败 ({processing_time:.2f}s)")
                print(f"错误: {plan_result.errors}")
                
                results.append({
                    "test": test_case['name'],
                    "status": "FAILED",
                    "time": processing_time,
                    "errors": plan_result.errors
                })
                
        except Exception as e:
            processing_time = time.time() - start_time
            print(f"❌ 测试异常 ({processing_time:.2f}s): {e}")
            
            results.append({
                "test": test_case['name'],
                "status": "ERROR",
                "time": processing_time,
                "error": str(e)
            })
    
    # 生成测试报告
    print("\n" + "="*60)
    print("📊 性能测试报告")
    print("="*60)
    
    total_tests = len(results)
    successful_tests = sum(1 for r in results if r['status'] == 'SUCCESS')
    total_time = sum(r['time'] for r in results)
    avg_time = total_time / total_tests if total_tests > 0 else 0
    
    print(f"总测试数: {total_tests}")
    print(f"成功测试: {successful_tests}")
    print(f"失败测试: {total_tests - successful_tests}")
    print(f"成功率: {successful_tests/total_tests*100:.1f}%")
    print(f"总耗时: {total_time:.2f}秒")
    print(f"平均耗时: {avg_time:.2f}秒/测试")
    
    print(f"\n📋 详细结果:")
    for result in results:
        status_icon = "✅" if result['status'] == 'SUCCESS' else "❌"
        print(f"{status_icon} {result['test']}: {result['status']} ({result['time']:.2f}s)")
        
        if result['status'] == 'SUCCESS' and 'operations' in result:
            print(f"   操作数: {result['operations']}")
        elif 'errors' in result:
            print(f"   错误: {result['errors']}")
        elif 'error' in result:
            print(f"   异常: {result['error']}")
    
    # 保存测试报告
    report_file = f"性能测试报告_{time.strftime('%Y%m%d_%H%M%S')}.json"
    with open(report_file, 'w', encoding='utf-8') as f:
        json.dump({
            "timestamp": time.strftime('%Y-%m-%d %H:%M:%S'),
            "summary": {
                "total_tests": total_tests,
                "successful_tests": successful_tests,
                "success_rate": successful_tests/total_tests*100,
                "total_time": total_time,
                "average_time": avg_time
            },
            "results": results
        }, f, indent=2, ensure_ascii=False)
    
    print(f"\n💾 测试报告已保存: {report_file}")
    
    # 给出建议
    print(f"\n💡 性能建议:")
    if avg_time < 2.0:
        print("🚀 性能优秀！平均响应时间很快")
    elif avg_time < 5.0:
        print("👍 性能良好，响应时间合理")
    else:
        print("⚠️  性能需要优化，响应时间较长")
    
    if successful_tests == total_tests:
        print("🎉 所有测试通过！系统功能正常")
    else:
        print("🔧 部分测试失败，建议检查配置和环境")


def main():
    """主函数"""
    print("🎯 AutoWord vNext 性能测试工具")
    print("="*50)
    
    try:
        performance_test_suite()
        
    except KeyboardInterrupt:
        print("\n\n⏹️  测试被用户中断")
        
    except Exception as e:
        print(f"\n❌ 测试工具异常: {e}")
        import traceback
        traceback.print_exc()
    
    print(f"\n🏁 测试完成")
    input("按回车键退出...")


if __name__ == "__main__":
    main()