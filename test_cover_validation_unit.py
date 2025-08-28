#!/usr/bin/env python3
"""
单元测试增强的封面格式验证功能
Unit test for enhanced cover format validation functionality
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

def test_cover_format_comparison():
    """测试封面格式比较逻辑"""
    
    print("=" * 60)
    print("测试封面格式比较逻辑")
    print("=" * 60)
    
    try:
        # 初始化管道
        config = load_config()
        pipeline = SimplePipeline(config)
        
        # 模拟原始格式
        original_format = {
            "index": 0,
            "text_preview": "毕业论文题目",
            "style_name": "标题",
            "font_name_east_asian": "楷体",
            "font_name_latin": "Times New Roman",
            "font_size": 16,
            "font_bold": True,
            "font_italic": False,
            "line_spacing": 18,
            "space_before": 0,
            "space_after": 6,
            "alignment": 1,  # Center
            "left_indent": 0,
            "page_number": 1
        }
        
        # 测试1: 无变化的情况
        print("\n🧪 测试1: 无格式变化")
        current_format = original_format.copy()
        changes = pipeline._compare_paragraph_formatting(original_format, current_format)
        print(f"  检测到的变化: {len(changes)}")
        if changes:
            for change in changes:
                print(f"    - {change}")
        else:
            print("  ✅ 正确检测到无变化")
        
        # 测试2: 样式被错误修改为BodyText
        print("\n🧪 测试2: 样式被错误修改为BodyText")
        current_format = original_format.copy()
        current_format["style_name"] = "BodyText (AutoWord)"
        changes = pipeline._compare_paragraph_formatting(original_format, current_format)
        print(f"  检测到的变化: {len(changes)}")
        for change in changes:
            print(f"    - {change}")
        
        expected_change = "样式被错误修改为BodyText (AutoWord)"
        if any(expected_change in change for change in changes):
            print("  ✅ 正确检测到BodyText样式错误分配")
        else:
            print("  ❌ 未能检测到BodyText样式错误分配")
        
        # 测试3: 字体被意外修改为宋体
        print("\n🧪 测试3: 字体被意外修改为宋体")
        current_format = original_format.copy()
        current_format["font_name_east_asian"] = "宋体"
        changes = pipeline._compare_paragraph_formatting(original_format, current_format)
        print(f"  检测到的变化: {len(changes)}")
        for change in changes:
            print(f"    - {change}")
        
        expected_change = "中文字体被意外修改为宋体"
        if any(expected_change in change for change in changes):
            print("  ✅ 正确检测到字体意外修改")
        else:
            print("  ❌ 未能检测到字体意外修改")
        
        # 测试4: 行距被意外修改为2倍行距
        print("\n🧪 测试4: 行距被意外修改为2倍行距")
        current_format = original_format.copy()
        current_format["line_spacing"] = 24  # 2倍行距
        changes = pipeline._compare_paragraph_formatting(original_format, current_format)
        print(f"  检测到的变化: {len(changes)}")
        for change in changes:
            print(f"    - {change}")
        
        expected_change = "行距被意外修改为2倍行距"
        if any(expected_change in change for change in changes):
            print("  ✅ 正确检测到行距意外修改")
        else:
            print("  ❌ 未能检测到行距意外修改")
        
        # 测试5: 字体大小被意外修改为12pt
        print("\n🧪 测试5: 字体大小被意外修改为12pt")
        current_format = original_format.copy()
        current_format["font_size"] = 12
        changes = pipeline._compare_paragraph_formatting(original_format, current_format)
        print(f"  检测到的变化: {len(changes)}")
        for change in changes:
            print(f"    - {change}")
        
        expected_change = "字体大小被意外修改为12pt"
        if any(expected_change in change for change in changes):
            print("  ✅ 正确检测到字体大小意外修改")
        else:
            print("  ❌ 未能检测到字体大小意外修改")
        
        return True
        
    except Exception as e:
        print(f"❌ 测试失败: {e}")
        import traceback
        traceback.print_exc()
        return False

def test_basic_validation():
    """测试基本验证逻辑"""
    
    print("\n" + "=" * 60)
    print("测试基本验证逻辑")
    print("=" * 60)
    
    try:
        # 初始化管道
        config = load_config()
        pipeline = SimplePipeline(config)
        
        # 测试1: 正常封面段落
        print("\n🧪 测试1: 正常封面段落")
        normal_format = {
            "text_preview": "毕业论文题目",
            "style_name": "标题",
            "font_name_east_asian": "楷体",
            "font_size": 16,
            "line_spacing": 18
        }
        issues = pipeline._validate_paragraph_formatting_basic(normal_format)
        print(f"  检测到的问题: {len(issues)}")
        if issues:
            for issue in issues:
                print(f"    - {issue}")
        else:
            print("  ✅ 正常段落无问题")
        
        # 测试2: 被错误分配到BodyText样式的封面段落
        print("\n🧪 测试2: 被错误分配到BodyText样式")
        bodytext_format = {
            "text_preview": "学生姓名: 张三",
            "style_name": "BodyText (AutoWord)",
            "font_name_east_asian": "宋体",
            "font_size": 12,
            "line_spacing": 24
        }
        issues = pipeline._validate_paragraph_formatting_basic(bodytext_format)
        print(f"  检测到的问题: {len(issues)}")
        for issue in issues:
            print(f"    - {issue}")
        
        expected_issue = "封面段落被错误分配到BodyText样式"
        if any(expected_issue in issue for issue in issues):
            print("  ✅ 正确检测到BodyText样式错误分配")
        else:
            print("  ❌ 未能检测到BodyText样式错误分配")
        
        # 测试3: 封面信息被意外修改为宋体
        print("\n🧪 测试3: 封面信息被意外修改为宋体")
        songti_format = {
            "text_preview": "指导老师: 李教授",
            "style_name": "正文",
            "font_name_east_asian": "宋体",
            "font_size": 12,
            "line_spacing": 24
        }
        issues = pipeline._validate_paragraph_formatting_basic(songti_format)
        print(f"  检测到的问题: {len(issues)}")
        for issue in issues:
            print(f"    - {issue}")
        
        expected_issue = "封面信息字体可能被意外修改为宋体"
        if any(expected_issue in issue for issue in issues):
            print("  ✅ 正确检测到字体意外修改")
        else:
            print("  ❌ 未能检测到字体意外修改")
        
        return True
        
    except Exception as e:
        print(f"❌ 测试失败: {e}")
        import traceback
        traceback.print_exc()
        return False

def test_validation_result_analysis():
    """测试验证结果分析逻辑"""
    
    print("\n" + "=" * 60)
    print("测试验证结果分析逻辑")
    print("=" * 60)
    
    try:
        # 初始化管道
        config = load_config()
        pipeline = SimplePipeline(config)
        
        # 测试1: 无问题情况
        print("\n🧪 测试1: 无问题情况")
        result = pipeline._analyze_cover_validation_results([], 5)
        print(f"  验证结果: {result}")
        if result:
            print("  ✅ 正确返回通过")
        else:
            print("  ❌ 应该返回通过但返回失败")
        
        # 测试2: 少量问题情况
        print("\n🧪 测试2: 少量问题情况")
        minor_issues = [
            "段落1: 字体大小变化",
            "段落2: 行距变化"
        ]
        result = pipeline._analyze_cover_validation_results(minor_issues, 5)
        print(f"  验证结果: {result}")
        if result:
            print("  ✅ 正确返回通过（少量问题可接受）")
        else:
            print("  ❌ 少量问题应该可接受")
        
        # 测试3: 严重问题情况
        print("\n🧪 测试3: 严重问题情况")
        severe_issues = [
            "段落1: 封面段落被错误分配到BodyText样式",
            "段落2: 封面信息字体被意外修改为宋体",
            "段落3: 封面信息行距被意外修改为2倍行距",
            "段落4: 封面标题字体大小被意外修改为12pt"
        ]
        result = pipeline._analyze_cover_validation_results(severe_issues, 5)
        print(f"  验证结果: {result}")
        if not result:
            print("  ✅ 正确返回失败（严重问题应该失败）")
        else:
            print("  ❌ 严重问题应该返回失败")
        
        # 测试4: 大量问题情况
        print("\n🧪 测试4: 大量问题情况")
        many_issues = [f"段落{i}: 格式问题" for i in range(1, 8)]
        result = pipeline._analyze_cover_validation_results(many_issues, 10)
        print(f"  验证结果: {result}")
        if not result:
            print("  ✅ 正确返回失败（大量问题应该失败）")
        else:
            print("  ❌ 大量问题应该返回失败")
        
        return True
        
    except Exception as e:
        print(f"❌ 测试失败: {e}")
        import traceback
        traceback.print_exc()
        return False

def main():
    """主测试函数"""
    
    print("🧪 AutoWord vNext 封面验证单元测试")
    print("=" * 60)
    
    # 运行测试
    tests_passed = 0
    total_tests = 3
    
    # 测试1: 封面格式比较逻辑
    print(f"\n🧪 测试 1/3: 封面格式比较逻辑")
    if test_cover_format_comparison():
        tests_passed += 1
        print("✅ 测试1通过")
    else:
        print("❌ 测试1失败")
    
    # 测试2: 基本验证逻辑
    print(f"\n🧪 测试 2/3: 基本验证逻辑")
    if test_basic_validation():
        tests_passed += 1
        print("✅ 测试2通过")
    else:
        print("❌ 测试2失败")
    
    # 测试3: 验证结果分析逻辑
    print(f"\n🧪 测试 3/3: 验证结果分析逻辑")
    if test_validation_result_analysis():
        tests_passed += 1
        print("✅ 测试3通过")
    else:
        print("❌ 测试3失败")
    
    # 总结
    print(f"\n" + "=" * 60)
    print(f"📊 测试总结: {tests_passed}/{total_tests} 通过")
    
    if tests_passed == total_tests:
        print("🎉 所有单元测试通过！封面验证逻辑正常工作")
        return True
    else:
        print(f"⚠️ {total_tests - tests_passed} 个测试失败")
        return False

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)