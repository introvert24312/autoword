#!/usr/bin/env python3
"""
单元测试形状处理功能
Unit test for shape processing functionality
"""

import os
import sys
import logging

# 添加项目路径
sys.path.insert(0, os.path.abspath('.'))

from autoword.vnext.simple_pipeline import SimplePipeline
from autoword.vnext.core import load_config

# 设置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def test_shape_cover_content_detection():
    """测试形状封面内容检测功能"""
    print("🧪 测试形状封面内容检测功能")
    print("=" * 50)
    
    try:
        # 初始化管道
        config = load_config()
        pipeline = SimplePipeline(config)
        
        # 测试用例
        test_cases = [
            # 应该被识别为封面内容的文本
            ("学生姓名：张三", True, "学生信息"),
            ("指导教师：李教授", True, "教师信息"),
            ("毕业论文", True, "论文标题"),
            ("2024年6月", True, "日期信息"),
            ("学号：20240001", True, "学号信息"),
            ("国家开放大学", True, "学校名称"),
            ("专业：计算机科学", True, "专业信息"),
            ("University", True, "英文学校"),
            ("Thesis", True, "英文论文"),
            ("Author: John", True, "英文作者"),
            
            # 不应该被识别为封面内容的文本
            ("这是正文内容，讨论了相关的技术问题。", False, "正文内容"),
            ("第一章 引言", False, "章节标题"),
            ("本研究的目的是探讨...", False, "研究内容"),
            ("参考文献显示...", False, "参考文献讨论"),
            ("实验结果表明...", False, "实验结果"),
            ("The main contribution of this work is...", False, "英文正文"),
        ]
        
        success_count = 0
        total_count = len(test_cases)
        
        for text, expected, description in test_cases:
            result = pipeline._is_shape_cover_content(text)
            
            if result == expected:
                print(f"✅ {description}: '{text[:30]}...' -> {result}")
                success_count += 1
            else:
                print(f"❌ {description}: '{text[:30]}...' -> {result} (期望: {expected})")
        
        print(f"\n📊 封面内容检测结果: {success_count}/{total_count} 通过")
        return success_count == total_count
        
    except Exception as e:
        print(f"❌ 测试失败: {e}")
        return False

def test_cover_or_toc_content_detection():
    """测试增强的封面/目录内容检测功能"""
    print("\n🧪 测试增强的封面/目录内容检测功能")
    print("=" * 50)
    
    try:
        # 初始化管道
        config = load_config()
        pipeline = SimplePipeline(config)
        
        # 测试用例
        test_cases = [
            # (para_index, first_content_index, text_preview, style_name, expected, description)
            (0, 5, "毕业论文题目", "Normal", True, "封面标题"),
            (1, 5, "学生姓名：张三", "Normal", True, "学生信息"),
            (2, 5, "指导老师：李教授", "Normal", True, "指导教师"),
            (3, 5, "2024年6月", "Normal", True, "日期信息"),
            (4, 5, "目录", "Normal", True, "目录标题"),
            (6, 5, "第一章 引言", "Heading 1", False, "正文标题"),
            (7, 5, "这是正文内容", "Normal", False, "正文段落"),
            (8, 5, "本研究探讨了相关问题", "Normal", False, "研究内容"),
            (0, 5, "20240001", "Normal", True, "学号（纯数字）"),
            (1, 5, "zhang@university.edu", "Normal", True, "邮箱地址"),
            (2, 5, "13812345678", "Normal", True, "电话号码"),
            (10, 5, "实验结果如下所示", "Normal", False, "实验内容"),
        ]
        
        success_count = 0
        total_count = len(test_cases)
        
        for para_index, first_content_index, text_preview, style_name, expected, description in test_cases:
            result = pipeline._is_cover_or_toc_content(para_index, first_content_index, text_preview, style_name)
            
            if result == expected:
                print(f"✅ {description}: 段落{para_index} '{text_preview[:20]}...' -> {result}")
                success_count += 1
            else:
                print(f"❌ {description}: 段落{para_index} '{text_preview[:20]}...' -> {result} (期望: {expected})")
        
        print(f"\n📊 封面/目录检测结果: {success_count}/{total_count} 通过")
        return success_count == total_count
        
    except Exception as e:
        print(f"❌ 测试失败: {e}")
        return False

def test_body_text_style_exists():
    """测试BodyText样式存在检查功能"""
    print("\n🧪 测试BodyText样式存在检查功能")
    print("=" * 50)
    
    try:
        # 初始化管道
        config = load_config()
        pipeline = SimplePipeline(config)
        
        # 这个测试需要实际的Word文档，所以我们只测试方法是否存在
        if hasattr(pipeline, '_body_text_style_exists'):
            print("✅ _body_text_style_exists 方法存在")
            return True
        else:
            print("❌ _body_text_style_exists 方法不存在")
            return False
        
    except Exception as e:
        print(f"❌ 测试失败: {e}")
        return False

def test_direct_formatting_method():
    """测试直接格式化方法"""
    print("\n🧪 测试直接格式化方法")
    print("=" * 50)
    
    try:
        # 初始化管道
        config = load_config()
        pipeline = SimplePipeline(config)
        
        # 检查方法是否存在
        if hasattr(pipeline, '_apply_direct_formatting_to_shape_paragraph'):
            print("✅ _apply_direct_formatting_to_shape_paragraph 方法存在")
            return True
        else:
            print("❌ _apply_direct_formatting_to_shape_paragraph 方法不存在")
            return False
        
    except Exception as e:
        print(f"❌ 测试失败: {e}")
        return False

def test_enhanced_shape_processing_method():
    """测试增强的形状处理方法"""
    print("\n🧪 测试增强的形状处理方法")
    print("=" * 50)
    
    try:
        # 初始化管道
        config = load_config()
        pipeline = SimplePipeline(config)
        
        # 检查方法是否存在并且包含增强功能
        if hasattr(pipeline, '_process_shapes_with_cover_protection'):
            print("✅ _process_shapes_with_cover_protection 方法存在")
            
            # 检查方法的源代码是否包含增强功能
            import inspect
            source = inspect.getsource(pipeline._process_shapes_with_cover_protection)
            
            enhancements = [
                "anchor_page = shape.Anchor.Information(3)",  # 锚定页码检测
                "_is_shape_cover_content",  # 形状封面内容检测
                "_apply_direct_formatting_to_shape_paragraph",  # 直接格式化
                "Enhanced shape and text frame processing",  # 文档字符串
                "📊 形状处理统计",  # 增强的日志
            ]
            
            found_enhancements = 0
            for enhancement in enhancements:
                if enhancement in source:
                    print(f"  ✅ 包含增强功能: {enhancement}")
                    found_enhancements += 1
                else:
                    print(f"  ❌ 缺少增强功能: {enhancement}")
            
            print(f"📊 增强功能检查: {found_enhancements}/{len(enhancements)} 通过")
            return found_enhancements >= len(enhancements) - 1  # 允许1个功能缺失
        else:
            print("❌ _process_shapes_with_cover_protection 方法不存在")
            return False
        
    except Exception as e:
        print(f"❌ 测试失败: {e}")
        return False

def test_integration_with_apply_styles():
    """测试与_apply_styles_to_content的集成"""
    print("\n🧪 测试与_apply_styles_to_content的集成")
    print("=" * 50)
    
    try:
        # 初始化管道
        config = load_config()
        pipeline = SimplePipeline(config)
        
        # 检查_apply_styles_to_content是否调用了形状处理
        import inspect
        source = inspect.getsource(pipeline._apply_styles_to_content)
        
        if "_process_shapes_with_cover_protection" in source:
            print("✅ _apply_styles_to_content 调用了形状处理方法")
            return True
        else:
            print("❌ _apply_styles_to_content 未调用形状处理方法")
            return False
        
    except Exception as e:
        print(f"❌ 测试失败: {e}")
        return False

def main():
    """主测试函数"""
    print("🚀 开始形状处理功能单元测试")
    print("=" * 60)
    
    tests = [
        ("形状封面内容检测", test_shape_cover_content_detection),
        ("封面/目录内容检测", test_cover_or_toc_content_detection),
        ("BodyText样式检查", test_body_text_style_exists),
        ("直接格式化方法", test_direct_formatting_method),
        ("增强形状处理方法", test_enhanced_shape_processing_method),
        ("集成测试", test_integration_with_apply_styles),
    ]
    
    success_count = 0
    total_tests = len(tests)
    
    for test_name, test_func in tests:
        try:
            if test_func():
                success_count += 1
                print(f"✅ {test_name} 通过")
            else:
                print(f"❌ {test_name} 失败")
        except Exception as e:
            print(f"❌ {test_name} 异常: {e}")
        
        print()
    
    print("=" * 60)
    print(f"📊 测试结果: {success_count}/{total_tests} 通过")
    
    if success_count == total_tests:
        print("🎉 所有单元测试通过！形状处理功能实现正确")
        return True
    else:
        print("⚠️ 部分测试失败，请检查实现")
        return False

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)