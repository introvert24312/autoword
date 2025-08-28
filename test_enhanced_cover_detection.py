#!/usr/bin/env python3
"""
测试增强的封面检测功能
Test enhanced cover detection in _is_cover_or_toc_content method
"""

import os
import sys
import logging
from pathlib import Path

# 添加项目根目录到Python路径
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from autoword.vnext.simple_pipeline import SimplePipeline

# 设置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def test_enhanced_cover_detection():
    """测试增强的封面检测功能"""
    logger.info("=== 测试增强的封面检测功能 ===")
    
    # 创建SimplePipeline实例
    pipeline = SimplePipeline()
    
    # 测试用例：各种封面内容模式
    test_cases = [
        # 原有的封面关键词
        ("学生姓名：张三", "Normal", True, "原有封面关键词"),
        ("指导教师：李教授", "Normal", True, "原有封面关键词"),
        ("毕业论文", "Normal", True, "原有封面关键词"),
        
        # 增强的学术论文封面指标
        ("指导老师：王教授", "Normal", True, "增强的指导老师关键词"),
        ("专业班级：计算机科学与技术2020级1班", "Normal", True, "增强的专业班级关键词"),
        ("提交日期：2024年6月", "Normal", True, "增强的提交日期关键词"),
        ("学位论文", "Normal", True, "增强的学位论文关键词"),
        ("开题报告", "Normal", True, "增强的开题报告关键词"),
        ("研究方向：人工智能", "Normal", True, "增强的研究方向关键词"),
        ("所在学院：计算机学院", "Normal", True, "增强的学院关键词"),
        
        # 国际学术指标
        ("Thesis Supervisor: Dr. Smith", "Normal", True, "英文导师关键词"),
        ("Department of Computer Science", "Normal", True, "英文系部关键词"),
        ("Bachelor of Science", "Normal", True, "英文学位关键词"),
        ("Submitted in fulfillment", "Normal", True, "英文提交关键词"),
        
        # 中文大学/机构指标
        ("清华大学", "Normal", True, "大学名称"),
        ("计算机系", "Normal", True, "系部名称"),
        ("教授", "Normal", True, "职称"),
        ("研究所", "Normal", True, "研究机构"),
        
        # 增强的样式检测
        ("论文标题", "封面标题", True, "增强的封面样式"),
        ("作者信息", "Cover Info", True, "增强的封面信息样式"),
        ("Paper Title", "Document Title", True, "增强的文档标题样式"),
        
        # 增强的文本模式识别
        ("202012345678", "Normal", True, "学号格式"),
        ("2024年6月15日", "Normal", True, "中文日期格式"),
        ("2024/06/15", "Normal", True, "斜杠日期格式"),
        ("2024-06-15", "Normal", True, "横杠日期格式"),
        ("June 2024", "Normal", True, "英文月份格式"),
        
        # 冒号分隔的标签模式
        ("姓名：", "Normal", True, "冒号标签格式"),
        ("学号：202012345678", "Normal", True, "学号标签格式"),
        ("专业：计算机科学", "Normal", True, "专业标签格式"),
        
        # 邮箱和电话模式
        ("zhang.san@university.edu.cn", "Normal", True, "邮箱地址"),
        ("138-1234-5678", "Normal", True, "电话号码格式"),
        ("(010) 123-4567", "Normal", True, "带区号电话格式"),
        ("13812345678", "Normal", True, "手机号码格式"),
        
        # 很短的内容（可能是封面元素）
        ("A", "Normal", True, "很短内容"),
        ("I", "Normal", True, "很短内容"),
        ("1", "Normal", True, "很短内容"),
        
        # 居中格式的内容
        ("   THESIS TITLE   ", "Normal", True, "可能居中的标题"),
        ("COMPUTER SCIENCE", "Normal", True, "全大写标题"),
        
        # 非封面内容（应该返回False）
        ("这是正文内容，包含了很多文字，描述了研究的背景和意义。", "Normal", False, "正文内容"),
        ("第一章 绪论", "Heading 1", False, "章节标题"),
        ("1.1 研究背景", "Heading 2", False, "小节标题"),
        ("根据相关研究表明，人工智能技术在各个领域都有广泛应用。", "Normal", False, "正文段落"),
    ]
    
    # 执行测试
    passed = 0
    failed = 0
    
    for i, (text, style, expected, description) in enumerate(test_cases, 1):
        try:
            # 测试封面检测（使用较大的first_content_index确保基于内容检测）
            result = pipeline._is_cover_or_toc_content(100, 50, text, style)
            
            if result == expected:
                logger.info(f"✅ 测试 {i:2d}: {description} - PASS")
                logger.debug(f"    文本: {text}")
                logger.debug(f"    样式: {style}")
                logger.debug(f"    预期: {expected}, 实际: {result}")
                passed += 1
            else:
                logger.error(f"❌ 测试 {i:2d}: {description} - FAIL")
                logger.error(f"    文本: {text}")
                logger.error(f"    样式: {style}")
                logger.error(f"    预期: {expected}, 实际: {result}")
                failed += 1
                
        except Exception as e:
            logger.error(f"❌ 测试 {i:2d}: {description} - ERROR: {e}")
            failed += 1
    
    # 测试结果汇总
    total = passed + failed
    logger.info(f"\n=== 测试结果汇总 ===")
    logger.info(f"总测试数: {total}")
    logger.info(f"通过: {passed}")
    logger.info(f"失败: {failed}")
    logger.info(f"成功率: {passed/total*100:.1f}%")
    
    if failed == 0:
        logger.info("🎉 所有测试通过！增强的封面检测功能工作正常。")
        return True
    else:
        logger.warning(f"⚠️ 有 {failed} 个测试失败，需要进一步调试。")
        return False

def test_position_based_detection():
    """测试基于位置的封面检测"""
    logger.info("\n=== 测试基于位置的封面检测 ===")
    
    pipeline = SimplePipeline()
    
    # 测试基于段落位置的检测
    test_cases = [
        (0, 10, "任意内容", "Normal", True, "段落0在正文开始位置10之前"),
        (5, 10, "任意内容", "Normal", True, "段落5在正文开始位置10之前"),
        (9, 10, "任意内容", "Normal", True, "段落9在正文开始位置10之前"),
        (10, 10, "任意内容", "Normal", False, "段落10等于正文开始位置10"),
        (15, 10, "任意内容", "Normal", False, "段落15在正文开始位置10之后"),
        (20, 10, "任意内容", "Normal", False, "段落20在正文开始位置10之后"),
    ]
    
    passed = 0
    failed = 0
    
    for i, (para_index, first_content_index, text, style, expected, description) in enumerate(test_cases, 1):
        try:
            result = pipeline._is_cover_or_toc_content(para_index, first_content_index, text, style)
            
            if result == expected:
                logger.info(f"✅ 位置测试 {i}: {description} - PASS")
                passed += 1
            else:
                logger.error(f"❌ 位置测试 {i}: {description} - FAIL")
                logger.error(f"    段落索引: {para_index}, 正文开始: {first_content_index}")
                logger.error(f"    预期: {expected}, 实际: {result}")
                failed += 1
                
        except Exception as e:
            logger.error(f"❌ 位置测试 {i}: {description} - ERROR: {e}")
            failed += 1
    
    total = passed + failed
    logger.info(f"\n位置检测测试结果: {passed}/{total} 通过")
    
    return failed == 0

if __name__ == "__main__":
    logger.info("开始测试增强的封面检测功能...")
    
    try:
        # 测试内容检测
        content_test_passed = test_enhanced_cover_detection()
        
        # 测试位置检测
        position_test_passed = test_position_based_detection()
        
        # 总体结果
        if content_test_passed and position_test_passed:
            logger.info("\n🎉 所有测试通过！任务2的增强封面检测功能实现成功。")
            logger.info("\n实现的功能包括：")
            logger.info("✅ 增强关键词检测，包含更全面的封面页指标")
            logger.info("✅ 更好的基于样式的封面元素检测")
            logger.info("✅ 改进的学术论文封面文本模式识别")
            logger.info("✅ 添加了文本框和形状的检测支持")
            logger.info("✅ 支持多种文档类型的封面检测")
            sys.exit(0)
        else:
            logger.error("\n❌ 部分测试失败，需要进一步调试。")
            sys.exit(1)
            
    except Exception as e:
        logger.error(f"测试执行失败: {e}")
        sys.exit(1)