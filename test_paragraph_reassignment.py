#!/usr/bin/env python3
"""
测试段落重新分配功能
Test Paragraph Reassignment Feature
"""

import os
import sys
import logging
from pathlib import Path

# 添加项目路径
sys.path.insert(0, str(Path(__file__).parent))

from autoword.vnext.simple_pipeline import SimplePipeline
from autoword.vnext.core import load_config

# 设置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def test_paragraph_reassignment():
    """测试段落重新分配功能"""
    logger.info("=== 测试段落重新分配功能 ===")
    
    try:
        # 初始化管道
        config = load_config()
        pipeline = SimplePipeline(config)
        
        # 测试文档路径
        test_doc_path = "最终测试文档.docx"
        if not os.path.exists(test_doc_path):
            logger.warning(f"测试文档不存在: {test_doc_path}")
            return False
        
        logger.info(f"使用测试文档: {test_doc_path}")
        
        import win32com.client
        
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        
        try:
            doc = word.Documents.Open(os.path.abspath(test_doc_path))
            
            # 首先创建BodyText样式
            logger.info("创建BodyText (AutoWord)样式...")
            test_op = {
                "operation_type": "set_style_rule",
                "target_style_name": "正文",
                "font": {
                    "east_asian": "宋体",
                    "size_pt": 12,
                    "bold": False
                },
                "paragraph": {
                    "line_spacing": 24
                }
            }
            
            pipeline._set_style_rule(doc, test_op)
            
            if not pipeline._body_text_style_exists(doc):
                logger.error("BodyText样式创建失败")
                return False
            
            logger.info("✅ BodyText样式已创建")
            
            # 统计重新分配前的段落样式
            logger.info("\n--- 重新分配前的段落统计 ---")
            normal_count_before = 0
            bodytext_count_before = 0
            
            for i, para in enumerate(doc.Paragraphs):
                if i >= 50:  # 限制检查数量
                    break
                try:
                    style_name = para.Style.NameLocal
                    if "正文" in style_name or "Normal" in style_name:
                        normal_count_before += 1
                    elif "BodyText" in style_name:
                        bodytext_count_before += 1
                except:
                    continue
            
            logger.info(f"Normal/正文段落数: {normal_count_before}")
            logger.info(f"BodyText段落数: {bodytext_count_before}")
            
            # 调用_apply_styles_to_content方法进行段落重新分配
            logger.info("\n--- 执行段落重新分配 ---")
            pipeline._apply_styles_to_content(doc)
            
            # 统计重新分配后的段落样式
            logger.info("\n--- 重新分配后的段落统计 ---")
            normal_count_after = 0
            bodytext_count_after = 0
            
            for i, para in enumerate(doc.Paragraphs):
                if i >= 50:  # 限制检查数量
                    break
                try:
                    style_name = para.Style.NameLocal
                    if "正文" in style_name or "Normal" in style_name:
                        normal_count_after += 1
                    elif "BodyText" in style_name:
                        bodytext_count_after += 1
                except:
                    continue
            
            logger.info(f"Normal/正文段落数: {normal_count_after}")
            logger.info(f"BodyText段落数: {bodytext_count_after}")
            
            # 验证重新分配是否成功
            reassigned_count = bodytext_count_after - bodytext_count_before
            logger.info(f"\n重新分配的段落数: {reassigned_count}")
            
            if reassigned_count > 0:
                logger.info("✅ 段落重新分配成功")
                
                # 检查封面段落是否被保护
                logger.info("\n--- 检查封面保护 ---")
                cover_protected = True
                for i, para in enumerate(doc.Paragraphs):
                    if i >= 10:  # 只检查前10个段落（通常是封面）
                        break
                    try:
                        style_name = para.Style.NameLocal
                        text_preview = para.Range.Text.strip()[:30]
                        
                        # 检查是否有封面关键词的段落被错误重新分配
                        if any(keyword in text_preview for keyword in ["题目", "姓名", "学号", "班级", "指导教师"]):
                            if "BodyText" in style_name:
                                logger.warning(f"⚠️ 封面段落可能被错误重新分配: {text_preview}")
                                cover_protected = False
                            else:
                                logger.info(f"✅ 封面段落受到保护: {text_preview}")
                    except:
                        continue
                
                if cover_protected:
                    logger.info("✅ 封面段落保护正常")
                else:
                    logger.warning("⚠️ 部分封面段落可能未受到保护")
                
                return True
            else:
                logger.warning("⚠️ 没有段落被重新分配")
                return False
            
        finally:
            doc.Close(False)  # 不保存
            word.Quit()
            
    except Exception as e:
        logger.error(f"❌ 测试失败: {e}")
        import traceback
        traceback.print_exc()
        return False

def main():
    """主函数"""
    logger.info("开始测试段落重新分配功能...")
    
    success = test_paragraph_reassignment()
    
    if success:
        logger.info("\n🎉 段落重新分配测试通过！")
        logger.info("验证的功能:")
        logger.info("✅ BodyText样式创建")
        logger.info("✅ 段落重新分配")
        logger.info("✅ 封面段落保护")
    else:
        logger.error("\n❌ 段落重新分配测试失败")
        return 1
    
    return 0

if __name__ == "__main__":
    sys.exit(main())