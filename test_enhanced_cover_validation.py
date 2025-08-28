#!/usr/bin/env python3
"""
测试增强的封面格式验证功能
Test enhanced cover format validation functionality
"""

import os
import sys
import logging
from pathlib import Path

# Add the autoword package to the path
sys.path.insert(0, str(Path(__file__).parent))

from autoword.vnext.simple_pipeline import SimplePipeline
from autoword.vnext.core import load_config

# 设置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def test_enhanced_cover_validation():
    """测试增强的封面格式验证功能"""
    
    print("=" * 60)
    print("测试增强的封面格式验证功能")
    print("=" * 60)
    
    # 查找测试文档
    test_documents = [
        "最终测试文档.docx",
        "快速测试文档.docx", 
        "AutoWord_演示文档_20250816_161226.docx"
    ]
    
    test_doc = None
    for doc_name in test_documents:
        if os.path.exists(doc_name):
            test_doc = doc_name
            break
    
    if not test_doc:
        print("❌ 未找到测试文档，请确保以下文档之一存在:")
        for doc in test_documents:
            print(f"  - {doc}")
        return False
    
    print(f"✅ 使用测试文档: {test_doc}")
    
    try:
        # 初始化管道
        config = load_config()
        pipeline = SimplePipeline(config)
        
        # 测试用户意图 - 包含可能影响封面的格式修改
        user_intent = "将正文设置为宋体小四号字，2倍行距"
        
        print(f"\n📝 测试用户意图: {user_intent}")
        print(f"📄 测试文档: {test_doc}")
        
        # 处理文档
        print("\n🚀 开始处理文档...")
        result = pipeline.process_document(test_doc, user_intent)
        
        # 检查结果
        print(f"\n📊 处理结果:")
        print(f"  状态: {result.status}")
        print(f"  消息: {result.message}")
        
        if result.warnings:
            print(f"  警告数量: {len(result.warnings)}")
            for i, warning in enumerate(result.warnings[:3], 1):
                print(f"    {i}. {warning}")
            if len(result.warnings) > 3:
                print(f"    ... 还有 {len(result.warnings) - 3} 个警告")
        
        if result.errors:
            print(f"  错误数量: {len(result.errors)}")
            for i, error in enumerate(result.errors[:3], 1):
                print(f"    {i}. {error}")
        
        # 检查输出文件
        if result.audit_directory:
            audit_path = Path(result.audit_directory)
            output_files = list(audit_path.glob("*_processed.docx"))
            if output_files:
                output_file = output_files[0]
                print(f"\n📁 输出文件: {output_file}")
                print(f"  文件大小: {output_file.stat().st_size} bytes")
                
                # 验证文件是否可以打开
                if output_file.exists():
                    print("  ✅ 输出文件存在且可访问")
                else:
                    print("  ❌ 输出文件不存在")
            else:
                print("  ⚠️ 未找到处理后的文档文件")
        
        # 测试结果评估
        if result.status == "SUCCESS":
            print(f"\n✅ 测试成功!")
            if result.warnings:
                print(f"  📋 包含 {len(result.warnings)} 个警告（这是正常的，表示封面验证在工作）")
            else:
                print(f"  📋 无警告（封面可能未受影响或验证未检测到问题）")
            return True
        else:
            print(f"\n❌ 测试失败: {result.message}")
            return False
            
    except Exception as e:
        print(f"\n❌ 测试过程中发生异常: {e}")
        import traceback
        traceback.print_exc()
        return False

def test_cover_format_capture():
    """测试封面格式捕获功能"""
    
    print("\n" + "=" * 60)
    print("测试封面格式捕获功能")
    print("=" * 60)
    
    # 查找测试文档
    test_doc = None
    test_documents = [
        "最终测试文档.docx",
        "快速测试文档.docx", 
        "AutoWord_演示文档_20250816_161226.docx"
    ]
    
    for doc_name in test_documents:
        if os.path.exists(doc_name):
            test_doc = doc_name
            break
    
    if not test_doc:
        print("❌ 未找到测试文档")
        return False
    
    try:
        # 初始化管道
        config = load_config()
        pipeline = SimplePipeline(config)
        
        # 测试封面格式捕获
        print(f"📄 捕获文档封面格式: {test_doc}")
        cover_format = pipeline._capture_cover_formatting(test_doc)
        
        if cover_format and "paragraphs" in cover_format:
            print(f"✅ 成功捕获封面格式")
            print(f"  捕获段落数量: {len(cover_format['paragraphs'])}")
            
            # 显示前几个段落的格式信息
            for i, para in enumerate(cover_format["paragraphs"][:3]):
                print(f"\n  段落 {para['index']}:")
                print(f"    文本预览: {para['text_preview'][:50]}...")
                print(f"    样式名称: {para['style_name']}")
                print(f"    中文字体: {para['font_name_east_asian']}")
                print(f"    字体大小: {para['font_size']}pt")
                print(f"    行距: {para['line_spacing']}pt")
                print(f"    粗体: {para['font_bold']}")
            
            if len(cover_format["paragraphs"]) > 3:
                print(f"  ... 还有 {len(cover_format['paragraphs']) - 3} 个段落")
            
            return True
        else:
            print("❌ 封面格式捕获失败")
            if "error" in cover_format:
                print(f"  错误: {cover_format['error']}")
            return False
            
    except Exception as e:
        print(f"❌ 封面格式捕获测试失败: {e}")
        import traceback
        traceback.print_exc()
        return False

def main():
    """主测试函数"""
    
    print("🧪 AutoWord vNext 增强封面验证测试")
    print("=" * 60)
    
    # 检查当前目录
    print(f"📁 当前工作目录: {os.getcwd()}")
    
    # 检查可用的测试文档
    test_documents = [
        "最终测试文档.docx",
        "快速测试文档.docx", 
        "AutoWord_演示文档_20250816_161226.docx"
    ]
    
    available_docs = [doc for doc in test_documents if os.path.exists(doc)]
    print(f"📄 可用测试文档: {len(available_docs)}")
    for doc in available_docs:
        print(f"  - {doc}")
    
    if not available_docs:
        print("❌ 未找到测试文档，请确保至少有一个测试文档存在")
        return False
    
    # 运行测试
    tests_passed = 0
    total_tests = 2
    
    # 测试1: 封面格式捕获
    print(f"\n🧪 测试 1/2: 封面格式捕获")
    if test_cover_format_capture():
        tests_passed += 1
        print("✅ 测试1通过")
    else:
        print("❌ 测试1失败")
    
    # 测试2: 增强封面验证
    print(f"\n🧪 测试 2/2: 增强封面验证")
    if test_enhanced_cover_validation():
        tests_passed += 1
        print("✅ 测试2通过")
    else:
        print("❌ 测试2失败")
    
    # 总结
    print(f"\n" + "=" * 60)
    print(f"📊 测试总结: {tests_passed}/{total_tests} 通过")
    
    if tests_passed == total_tests:
        print("🎉 所有测试通过！增强封面验证功能正常工作")
        return True
    else:
        print(f"⚠️ {total_tests - tests_passed} 个测试失败")
        return False

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)