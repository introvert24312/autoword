"""
AutoWord vNext 简化管道
直接可用的文档处理管道，集成自定义API
"""

import os
import json
import time
import logging
from pathlib import Path
from typing import Optional, Dict, Any

from .core import VNextConfig, CustomLLMClient, load_config
from .models import ProcessingResult, StructureV1, PlanV1
from .exceptions import VNextError, ExtractionError, PlanningError, ExecutionError

# 设置日志
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class SimplePipeline:
    """简化的AutoWord vNext管道"""
    
    def __init__(self, config: Optional[VNextConfig] = None):
        """初始化管道"""
        self.config = config or load_config()
        self.llm_client = CustomLLMClient(self.config.llm)
        self._body_text_style_created = False  # Track if BodyText (AutoWord) style was created
        
    def process_document(self, docx_path: str, user_intent: str) -> ProcessingResult:
        """处理文档"""
        logger.info(f"开始处理文档: {docx_path}")
        logger.info(f"用户意图: {user_intent}")
        
        # Initialize cover format tracking
        self._original_cover_format = None
        
        try:
            # 0. Capture original cover formatting for validation
            logger.info("步骤0: 捕获原始封面格式...")
            self._original_cover_format = self._capture_cover_formatting(docx_path)
            
            # 1. 提取文档结构
            logger.info("步骤1: 提取文档结构...")
            structure = self._extract_structure(docx_path)
            
            # 2. 生成执行计划
            logger.info("步骤2: 生成执行计划...")
            plan = self._generate_plan(structure, user_intent)
            
            # 3. 执行计划
            logger.info("步骤3: 执行计划...")
            result_path = self._execute_plan(docx_path, plan)
            
            # 4. 验证结果
            logger.info("步骤4: 验证结果...")
            validation_result = self._validate_result(result_path)
            
            if validation_result:
                logger.info("✅ 文档处理成功！")
                return ProcessingResult(
                    status="SUCCESS",
                    message="文档处理成功",
                    audit_directory=str(Path(result_path).parent)
                )
            else:
                logger.warning("⚠️ 验证失败，但文档已生成")
                warnings = ["部分验证检查未通过"]
                
                # Add specific cover format warnings if available
                if hasattr(self, '_cover_validation_warnings'):
                    warnings.extend(self._cover_validation_warnings)
                
                return ProcessingResult(
                    status="SUCCESS",
                    message="文档处理完成，但验证有警告",
                    warnings=warnings,
                    audit_directory=str(Path(result_path).parent)
                )
                
        except Exception as e:
            logger.error(f"❌ 处理失败: {e}")
            return ProcessingResult(
                status="FAILED_VALIDATION",
                message=f"处理失败: {str(e)}",
                errors=[str(e)]
            )
    
    def _extract_structure(self, docx_path: str) -> Dict[str, Any]:
        """提取文档结构（简化版）"""
        try:
            import win32com.client
            
            # 启动Word应用
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            
            try:
                # 打开文档 - 使用绝对路径避免COM路径问题
                abs_path = os.path.abspath(docx_path)
                doc = word.Documents.Open(abs_path)
                
                # 提取基本信息
                structure = {
                    "metadata": {
                        "title": doc.BuiltInDocumentProperties("Title").Value or "未知标题",
                        "author": doc.BuiltInDocumentProperties("Author").Value or "未知作者",
                        "paragraph_count": doc.Paragraphs.Count,
                        "word_count": doc.Words.Count
                    },
                    "paragraphs": [],
                    "headings": [],
                    "styles": []
                }
                
                # 提取段落信息
                for i, para in enumerate(doc.Paragraphs):
                    if i >= 100:  # 限制段落数量，避免过大
                        break
                    
                    text = para.Range.Text.strip()[:120]  # 限制预览文本长度
                    if text:
                        # 获取outline level和页码信息
                        outline_level = para.OutlineLevel
                        
                        try:
                            page_number = para.Range.Information(3)  # wdActiveEndPageNumber
                        except:
                            page_number = 1
                        
                        para_info = {
                            "index": i,
                            "style_name": para.Style.NameLocal,
                            "preview_text": text,
                            "outline_level": outline_level,
                            "page_number": page_number,
                            "is_heading": outline_level <= 9,  # outline level 1-9 通常是标题
                            "is_cover": page_number == 1  # 第1页标记为封面
                        }
                        structure["paragraphs"].append(para_info)
                        
                        # 基于outline level判断是否为标题
                        if outline_level <= 9:  # 1-9级都是标题
                            structure["headings"].append({
                                "paragraph_index": i,
                                "level": outline_level,
                                "text": text,
                                "style_name": para.Style.NameLocal,
                                "outline_level": outline_level
                            })
                
                # 提取样式信息
                for style in doc.Styles:
                    if len(structure["styles"]) >= 20:  # 限制样式数量
                        break
                    
                    try:
                        structure["styles"].append({
                            "name": style.NameLocal,
                            "type": "paragraph"  # 简化处理
                        })
                    except:
                        continue
                
                doc.Close()
                return structure
                
            finally:
                word.Quit()
                
        except Exception as e:
            raise ExtractionError(f"文档结构提取失败: {e}")
    
    def _generate_plan(self, structure: Dict[str, Any], user_intent: str) -> Dict[str, Any]:
        """生成执行计划"""
        try:
            # 对于批注驱动的处理，优先使用基本计划逻辑，因为它更准确
            if "基于文档批注的处理指令" in user_intent:
                logger.info("检测到批注驱动处理，使用优化的基本计划")
                return self._create_basic_plan(user_intent)
            
            # 将结构转换为JSON字符串
            structure_json = json.dumps(structure, ensure_ascii=False, indent=2)
            
            # 调用LLM生成计划
            plan_text = self.llm_client.generate_plan(structure_json, user_intent)
            
            # 解析JSON响应
            try:
                plan = json.loads(plan_text)
                
                # 验证计划格式
                if "ops" not in plan:
                    plan = {"schema_version": "plan.v1", "ops": []}
                
                logger.info(f"生成了 {len(plan['ops'])} 个操作")
                for i, op in enumerate(plan['ops']):
                    logger.info(f"  {i+1}. {op.get('operation_type', '未知操作')}")
                
                return plan
                
            except json.JSONDecodeError:
                # 如果JSON解析失败，创建一个基本计划
                logger.warning("LLM响应不是有效JSON，创建基本计划")
                return self._create_basic_plan(user_intent)
                
        except Exception as e:
            logger.warning(f"计划生成失败，使用基本计划: {e}")
            return self._create_basic_plan(user_intent)
    
    def _create_basic_plan(self, user_intent: str) -> Dict[str, Any]:
        """创建基本执行计划"""
        ops = []
        
        intent_lower = user_intent.lower()
        
        # 处理目录相关操作
        if "删掉目录里的摘要" in intent_lower or "删掉目录里的参考文献" in intent_lower:
            # 删除目录中的特定条目，而不是删除整个章节
            ops.append({
                "operation_type": "delete_toc",
                "target": "摘要和参考文献"
            })
        
        # 处理章节删除
        if "删除" in intent_lower and "摘要" in intent_lower and "目录" not in intent_lower:
            ops.append({
                "operation_type": "delete_section_by_heading",
                "heading_text": "摘要",
                "level": 1,
                "match": "CONTAINS",
                "case_sensitive": False
            })
        
        if "删除" in intent_lower and ("参考文献" in intent_lower or "references" in intent_lower) and "目录" not in intent_lower:
            ops.append({
                "operation_type": "delete_section_by_heading", 
                "heading_text": "参考文献",
                "level": 1,
                "match": "CONTAINS",
                "case_sensitive": False
            })
        
        # 处理目录更新
        if "更新" in intent_lower and "目录" in intent_lower:
            ops.append({
                "operation_type": "update_toc"
            })
        
        # 处理分页符插入
        if "插入分页符" in intent_lower or "分页符" in intent_lower:
            # 默认在封面后插入分页符
            ops.append({
                "operation_type": "insert_page_break",
                "position": "after_cover"  # 在封面后插入
            })
        
        # 处理1级标题格式
        if "1级标题" in intent_lower or "一级标题" in intent_lower:
            font_size = 14  # 小四号
            if "小四" in intent_lower:
                font_size = 12
            
            ops.append({
                "operation_type": "set_style_rule",
                "target_style_name": "标题 1",
                "font": {
                    "east_asian": "楷体",
                    "size_pt": font_size,
                    "bold": True
                },
                "paragraph": {
                    "line_spacing": 24 if "2倍行距" in intent_lower else 12
                }
            })
        
        # 处理2级标题格式
        if "二级标题" in intent_lower or "2级标题" in intent_lower:
            font_size = 12  # 小四号
            
            ops.append({
                "operation_type": "set_style_rule",
                "target_style_name": "标题 2",
                "font": {
                    "east_asian": "宋体",
                    "size_pt": font_size,
                    "bold": True
                },
                "paragraph": {
                    "line_spacing": 24 if "2倍行距" in intent_lower else 12
                }
            })
        
        # 处理正文格式
        if "正文" in intent_lower:
            font_size = 12  # 小四号
            
            ops.append({
                "operation_type": "set_style_rule",
                "target_style_name": "正文",
                "font": {
                    "east_asian": "宋体",
                    "size_pt": font_size,
                    "bold": False
                },
                "paragraph": {
                    "line_spacing": 24 if "2倍行距" in intent_lower else 12
                }
            })
        
        # 处理目录左对齐
        if "左顶格" in intent_lower or "左对齐" in intent_lower:
            ops.append({
                "operation_type": "set_style_rule",
                "target_style_name": "TOC 1",
                "paragraph": {
                    "alignment": "LEFT",
                    "indent_left": 0
                }
            })
            ops.append({
                "operation_type": "set_style_rule",
                "target_style_name": "TOC 2",
                "paragraph": {
                    "alignment": "LEFT",
                    "indent_left": 0
                }
            })
        
        return {
            "schema_version": "plan.v1",
            "ops": ops
        }
    
    def _execute_plan(self, docx_path: str, plan: Dict[str, Any]) -> str:
        """执行计划（简化版）"""
        try:
            import win32com.client
            
            # 创建输出文件路径
            input_path = Path(docx_path)
            output_path = input_path.parent / f"{input_path.stem}_processed{input_path.suffix}"
            
            # 复制文件
            import shutil
            shutil.copy2(docx_path, output_path)
            
            # 启动Word应用
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            
            try:
                # 打开文档 - 使用绝对路径
                abs_output_path = os.path.abspath(str(output_path))
                doc = word.Documents.Open(abs_output_path)
                
                # 执行操作
                for i, op in enumerate(plan.get("ops", [])):
                    op_type = op.get("operation_type")
                    logger.info(f"执行操作 {i+1}: {op_type}")
                    
                    try:
                        if op_type == "delete_section_by_heading":
                            self._delete_section_by_heading(doc, op)
                        elif op_type == "delete_toc":
                            self._delete_toc(doc, op)
                        elif op_type == "update_toc":
                            self._update_toc(doc)
                        elif op_type == "set_style_rule":
                            self._set_style_rule(doc, op)
                        elif op_type == "insert_page_break":
                            self._insert_page_break(doc, op)
                        else:
                            logger.warning(f"未知操作类型: {op_type}")
                    
                    except Exception as e:
                        logger.warning(f"操作 {op_type} 执行失败: {e}")
                        continue
                
                # 强制应用样式更改到所有段落
                self._apply_styles_to_content(doc)
                
                # 更新所有域和重新计算
                doc.Fields.Update()
                doc.Repaginate()
                
                # 保存文档
                doc.Save()
                doc.Close()
                
                logger.info(f"文档已保存到: {output_path}")
                return str(output_path)
                
            finally:
                word.Quit()
                
        except Exception as e:
            raise ExecutionError(f"计划执行失败: {e}")
    
    def _delete_section_by_heading(self, doc, op):
        """删除章节"""
        heading_text = op.get("heading_text", "")
        
        for para in doc.Paragraphs:
            if heading_text in para.Range.Text:
                # 找到匹配的标题，删除到下一个同级标题
                start_para = para
                current_para = para.Next()
                
                # 删除段落直到找到下一个标题或文档结束
                while current_para and not (current_para.Style.NameLocal.startswith(("标题", "Heading"))):
                    next_para = current_para.Next()
                    current_para.Range.Delete()
                    current_para = next_para
                
                # 删除标题段落本身
                start_para.Range.Delete()
                break
    
    def _delete_toc(self, doc, op):
        """删除目录中的特定部分"""
        try:
            # 根据操作参数决定删除什么
            target = op.get("target", "")
            
            if "摘要" in target or "参考文献" in target:
                # 删除目录中的摘要和参考文献条目
                paras_to_delete = []
                for para in doc.Paragraphs:
                    text = para.Range.Text.strip()
                    if ("摘要" in text or "参考文献" in text) and len(text) < 50:
                        # 这可能是目录条目
                        paras_to_delete.append(para)
                
                # 安全删除段落
                for para in paras_to_delete:
                    try:
                        logger.info(f"删除目录条目: {para.Range.Text.strip()}")
                        para.Range.Text = ""  # 清空内容而不是删除段落
                    except:
                        continue
            else:
                # 尝试删除整个目录
                try:
                    # 查找TOC域并删除
                    fields_to_delete = []
                    for field in doc.Fields:
                        if field.Type == 13:  # wdFieldTOC
                            fields_to_delete.append(field)
                    
                    for field in fields_to_delete:
                        field.Select()
                        field.Delete()
                        logger.info("目录域已删除")
                        
                except Exception as e:
                    logger.warning(f"目录域删除失败: {e}")
                    # 尝试查找并清空目录段落
                    for para in doc.Paragraphs:
                        text = para.Range.Text.strip()
                        if "目录" in text and len(text) < 10:
                            para.Range.Text = ""
                            logger.info("目录段落已清空")
                            break
                    
        except Exception as e:
            logger.warning(f"目录删除失败: {e}")
    
    def _update_toc(self, doc):
        """更新目录"""
        try:
            # 更新所有域
            doc.Fields.Update()
            logger.info("目录已更新")
        except Exception as e:
            logger.warning(f"目录更新失败: {e}")
    
    def _insert_page_break(self, doc, op):
        """插入分页符"""
        try:
            position = op.get("position", "after_cover")
            
            if position == "after_cover":
                # 智能查找封面结束位置
                page_break_inserted = False
                
                # 方法1: 查找第一个标题前插入分页符
                for i, para in enumerate(doc.Paragraphs):
                    if para.Style.NameLocal in ["标题 1", "Heading 1", "标题1"]:
                        # 检查是否已经有分页符
                        if i > 0:
                            prev_para = doc.Paragraphs[i-1]
                            if not prev_para.Range.Text.strip().endswith('\f'):  # \f 是分页符
                                para.Range.InsertBreak(7)  # 7 = wdPageBreak
                                logger.info(f"已在第一个标题 '{para.Range.Text.strip()[:20]}...' 前插入分页符")
                                page_break_inserted = True
                        break
                
                # 方法2: 查找典型的正文开始标志
                if not page_break_inserted:
                    for i, para in enumerate(doc.Paragraphs):
                        text = para.Range.Text.strip()
                        
                        # 检查是否为正文开始的典型标志
                        if any(keyword in text for keyword in [
                            "摘要", "Abstract", "引言", "前言", "第一章", "第1章", 
                            "1.", "一、", "1 ", "Chapter", "目录"
                        ]):
                            # 确保不是在第一段就插入
                            if i > 3:  # 至少跳过前几段（通常是封面内容）
                                para.Range.InsertBreak(7)
                                logger.info(f"已在段落 '{text[:20]}...' 前插入分页符")
                                page_break_inserted = True
                                break
                
                # 方法3: 基于页码判断
                if not page_break_inserted:
                    for i, para in enumerate(doc.Paragraphs):
                        try:
                            page_number = para.Range.Information(3)  # wdActiveEndPageNumber
                            # 如果找到第2页的第一个段落
                            if page_number == 2 and i > 0:
                                para.Range.InsertBreak(7)
                                logger.info(f"已在第2页第一个段落前插入分页符")
                                page_break_inserted = True
                                break
                        except:
                            continue
                
                # 方法4: 如果以上都失败，在合理位置插入
                if not page_break_inserted and doc.Paragraphs.Count > 5:
                    # 在第6个段落前插入（通常封面不会超过5段）
                    insert_position = min(6, doc.Paragraphs.Count - 1)
                    doc.Paragraphs[insert_position].Range.InsertBreak(7)
                    logger.info(f"已在第{insert_position}个段落前插入分页符")
                    page_break_inserted = True
                
                if not page_break_inserted:
                    logger.warning("未能找到合适的分页符插入位置")
            
        except Exception as e:
            logger.error(f"插入分页符失败: {e}")
    
    def _set_style_rule(self, doc, op):
        """设置样式规则"""
        try:
            # 尝试多种方式获取样式名称
            style_name = op.get("target_style_name") or op.get("target_style") or op.get("style_name", "")
            font_spec = op.get("font", {})
            
            # 如果样式名称为空，跳过
            if not style_name:
                logger.warning("样式名称为空，跳过样式设置")
                return
            
            # NEW: Check if targeting Normal/正文 styles for cover protection
            if style_name in ["Normal", "正文", "Body Text", "正文文本"]:
                logger.info(f"检测到Normal/正文样式修改，创建BodyText (AutoWord)样式以保护封面")
                return self._create_and_apply_body_text_style(doc, op)
            
            # 查找样式（支持中英文样式名）
            style = None
            style_mappings = {
                "Heading 1": "标题 1",
                "Heading 2": "标题 2", 
                "Heading 3": "标题 3",
                "Normal": "正文",
                "标题 1": "Heading 1",
                "标题 2": "Heading 2",
                "标题 3": "Heading 3",
                "正文": "Normal",
                "BodyText (AutoWord)": "BodyText (AutoWord)"  # Add body text style to mappings
            }
            
            # 首先尝试直接匹配
            for s in doc.Styles:
                if s.NameLocal == style_name:
                    style = s
                    break
            
            # 如果没找到，尝试映射
            if not style and style_name in style_mappings:
                mapped_name = style_mappings[style_name]
                for s in doc.Styles:
                    if s.NameLocal == mapped_name:
                        style = s
                        break
            
            if style:
                # 设置字体属性
                if "east_asian" in font_spec:
                    style.Font.NameFarEast = font_spec["east_asian"]
                if "latin" in font_spec:
                    style.Font.Name = font_spec["latin"]
                if "size_pt" in font_spec:
                    style.Font.Size = font_spec["size_pt"]
                if "bold" in font_spec:
                    style.Font.Bold = font_spec["bold"]
                
                # 设置段落格式
                paragraph_spec = op.get("paragraph", {})
                if "line_spacing" in paragraph_spec:
                    style.ParagraphFormat.LineSpacing = paragraph_spec["line_spacing"]
                if "space_before" in paragraph_spec:
                    style.ParagraphFormat.SpaceBefore = paragraph_spec["space_before"]
                if "space_after" in paragraph_spec:
                    style.ParagraphFormat.SpaceAfter = paragraph_spec["space_after"]
                if "alignment" in paragraph_spec:
                    alignment_map = {
                        "LEFT": 0,    # wdAlignParagraphLeft
                        "CENTER": 1,  # wdAlignParagraphCenter
                        "RIGHT": 2,   # wdAlignParagraphRight
                        "JUSTIFY": 3  # wdAlignParagraphJustify
                    }
                    if paragraph_spec["alignment"] in alignment_map:
                        style.ParagraphFormat.Alignment = alignment_map[paragraph_spec["alignment"]]
                if "indent_left" in paragraph_spec:
                    style.ParagraphFormat.LeftIndent = paragraph_spec["indent_left"]
                
                logger.info(f"样式 {style_name} 已更新")
            else:
                logger.warning(f"未找到样式: {style_name}")
                
        except Exception as e:
            logger.warning(f"样式设置失败: {e}")
    
    def _create_and_apply_body_text_style(self, doc, op):
        """Create BodyText (AutoWord) style and reassign paragraphs"""
        body_style_name = "BodyText (AutoWord)"
        
        try:
            # Check if style already exists
            body_style = None
            for s in doc.Styles:
                if s.NameLocal == body_style_name:
                    body_style = s
                    break
            
            # Create if doesn't exist
            if not body_style:
                logger.info(f"创建新样式: {body_style_name}")
                body_style = doc.Styles.Add(body_style_name, 1)  # wdStyleTypeParagraph
                
                # Clone from Normal style as base
                try:
                    normal_style = None
                    for s in doc.Styles:
                        if s.NameLocal in ["Normal", "正文"]:
                            normal_style = s
                            break
                    
                    if normal_style:
                        body_style.BaseStyle = normal_style
                        logger.info(f"基于 {normal_style.NameLocal} 样式创建 {body_style_name}")
                    else:
                        logger.warning("未找到Normal/正文样式作为基础样式")
                        
                except Exception as e:
                    logger.warning(f"设置基础样式失败: {e}")
            else:
                logger.info(f"使用现有样式: {body_style_name}")
            
            # Apply formatting from operation to the body text style
            font_spec = op.get("font", {})
            if "east_asian" in font_spec:
                body_style.Font.NameFarEast = font_spec["east_asian"]
                logger.info(f"设置中文字体: {font_spec['east_asian']}")
            if "latin" in font_spec:
                body_style.Font.Name = font_spec["latin"]
                logger.info(f"设置西文字体: {font_spec['latin']}")
            if "size_pt" in font_spec:
                body_style.Font.Size = font_spec["size_pt"]
                logger.info(f"设置字体大小: {font_spec['size_pt']}pt")
            if "bold" in font_spec:
                body_style.Font.Bold = font_spec["bold"]
                logger.info(f"设置粗体: {font_spec['bold']}")
            
            # Apply paragraph formatting
            paragraph_spec = op.get("paragraph", {})
            if "line_spacing" in paragraph_spec:
                body_style.ParagraphFormat.LineSpacing = paragraph_spec["line_spacing"]
                logger.info(f"设置行距: {paragraph_spec['line_spacing']}")
            if "space_before" in paragraph_spec:
                body_style.ParagraphFormat.SpaceBefore = paragraph_spec["space_before"]
            if "space_after" in paragraph_spec:
                body_style.ParagraphFormat.SpaceAfter = paragraph_spec["space_after"]
            if "alignment" in paragraph_spec:
                alignment_map = {
                    "LEFT": 0,    # wdAlignParagraphLeft
                    "CENTER": 1,  # wdAlignParagraphCenter
                    "RIGHT": 2,   # wdAlignParagraphRight
                    "JUSTIFY": 3  # wdAlignParagraphJustify
                }
                if paragraph_spec["alignment"] in alignment_map:
                    body_style.ParagraphFormat.Alignment = alignment_map[paragraph_spec["alignment"]]
            if "indent_left" in paragraph_spec:
                body_style.ParagraphFormat.LeftIndent = paragraph_spec["indent_left"]
            
            logger.info(f"样式 {body_style_name} 已创建/更新")
            
            # Mark that body text style exists for later paragraph reassignment
            # This will be used in _apply_styles_to_content method
            self._body_text_style_created = True
            
        except Exception as e:
            logger.warning(f"BodyText样式创建失败: {e}")
            # Fallback to original behavior if body text style creation fails
            original_style_name = op.get("target_style_name") or op.get("target_style") or op.get("style_name", "")
            if original_style_name in ["Normal", "正文"]:
                op_copy = op.copy()
                op_copy["target_style_name"] = "Normal"  # Force to Normal for fallback
                return self._apply_style_directly(doc, op_copy)
    
    def _apply_style_directly(self, doc, op):
        """Apply style directly (fallback method)"""
        style_name = op.get("target_style_name") or op.get("target_style") or op.get("style_name", "")
        font_spec = op.get("font", {})
        
        # Find the style
        style = None
        for s in doc.Styles:
            if s.NameLocal == style_name:
                style = s
                break
        
        if style:
            # Apply formatting directly to the style
            if "east_asian" in font_spec:
                style.Font.NameFarEast = font_spec["east_asian"]
            if "latin" in font_spec:
                style.Font.Name = font_spec["latin"]
            if "size_pt" in font_spec:
                style.Font.Size = font_spec["size_pt"]
            if "bold" in font_spec:
                style.Font.Bold = font_spec["bold"]
            
            # Apply paragraph formatting
            paragraph_spec = op.get("paragraph", {})
            if "line_spacing" in paragraph_spec:
                style.ParagraphFormat.LineSpacing = paragraph_spec["line_spacing"]
            if "space_before" in paragraph_spec:
                style.ParagraphFormat.SpaceBefore = paragraph_spec["space_before"]
            if "space_after" in paragraph_spec:
                style.ParagraphFormat.SpaceAfter = paragraph_spec["space_after"]
            
            logger.info(f"直接应用样式 {style_name}")
        else:
            logger.warning(f"未找到样式进行直接应用: {style_name}")
    
    def _body_text_style_exists(self, doc):
        """Check if BodyText (AutoWord) style exists"""
        for s in doc.Styles:
            if s.NameLocal == "BodyText (AutoWord)":
                return True
        return False
    
    def _is_cover_content(self, page_number: int, text_preview: str, style_name: str) -> bool:
        """判断是否为封面内容"""
        # 第1页内容都是封面
        if page_number == 1:
            return True
        
        # 扩展封面关键词检测（更严格）
        cover_keywords = [
            "题目", "姓名", "学号", "班级", "指导教师", "分校", 
            "毕业论文", "毕业设计", "国家开放大学", "学院", "专业",
            "年月日", "学生姓名", "导师", "指导老师", "作者：",
            "2024年", "2023年", "2025年",  # 年份
            "指导老", "学号：", "班级：", "分校：", "题目：",  # 带冒号的标签
            "开放大学", "教学中心"
        ]
        
        text_lower = text_preview.lower()
        for keyword in cover_keywords:
            if keyword in text_lower:
                return True
        
        # 检查是否为纯数字（学号等）
        if text_preview.strip().isdigit() and len(text_preview.strip()) > 6:
            return True
        
        # 检查封面样式
        cover_styles = ["封面", "cover", "title", "标题页"]
        style_lower = style_name.lower()
        for style in cover_styles:
            if style in style_lower:
                return True
        
        # 检查是否为日期格式
        import re
        if re.match(r'.*\d{4}年\d{1,2}月.*', text_preview):
            return True
        
        return False
    
    def _find_first_content_section(self, doc):
        """查找第一个正文节的开始位置"""
        try:
            # 方法1: 查找分页符
            for i, para in enumerate(doc.Paragraphs):
                # 检查段落是否包含分页符
                if '\f' in para.Range.Text:  # \f 是分页符字符
                    logger.info(f"找到分页符在段落 {i}")
                    return i + 1  # 返回分页符后的段落索引
                
                # 检查段落后是否有分页符
                try:
                    if para.Range.ParagraphFormat.PageBreakBefore:
                        logger.info(f"找到分页符设置在段落 {i}")
                        return i
                except:
                    pass
            
            # 方法2: 查找第一个有页码的段落
            for i, para in enumerate(doc.Paragraphs):
                try:
                    # 检查段落是否在有页码的页面上
                    page_number = para.Range.Information(3)  # wdActiveEndPageNumber
                    if page_number > 1:  # 如果页码大于1，说明进入了正文区域
                        logger.info(f"找到第一个有页码的段落在索引 {i}, 页码 {page_number}")
                        return i
                except:
                    continue
            
            # 方法3: 查找典型的正文开始标志
            for i, para in enumerate(doc.Paragraphs):
                text = para.Range.Text.strip().lower()
                # 查找正文开始的典型标志
                if any(keyword in text for keyword in [
                    "摘要", "abstract", "引言", "前言", "第一章", "第1章", 
                    "1.", "一、", "1 ", "chapter"
                ]) and i > 5:  # 确保不是在文档开头
                    logger.info(f"找到正文开始标志在段落 {i}: {text[:30]}")
                    return i
            
            # 方法4: 基于节（Section）判断
            try:
                if doc.Sections.Count > 1:
                    # 如果有多个节，第二个节通常是正文开始
                    second_section_start = doc.Sections[2].Range.Start
                    for i, para in enumerate(doc.Paragraphs):
                        if para.Range.Start >= second_section_start:
                            logger.info(f"基于节判断，正文开始于段落 {i}")
                            return i
            except:
                pass
            
            # 如果都没找到，返回一个较大的数字（表示整个文档都是封面）
            logger.warning("未找到正文开始位置，将整个文档视为封面")
            return 999999
            
        except Exception as e:
            logger.warning(f"查找正文开始位置失败: {e}")
            return 999999
    
    def _is_cover_or_toc_content(self, para_index, first_content_index, text_preview, style_name):
        """Enhanced cover/TOC content detection with comprehensive indicators"""
        
        # If paragraph is before main content starts, it's cover/TOC content
        if para_index < first_content_index:
            logger.debug(f"段落 {para_index} 在正文开始位置 {first_content_index} 之前，视为封面/目录内容")
            return True
        
        # Additional check: even in main content area, protect obvious cover/TOC content
        text_lower = text_preview.lower()
        
        # Enhanced cover keywords with more comprehensive academic paper indicators
        cover_keywords = [
            # Original keywords
            "题目", "姓名", "学号", "班级", "指导教师", "分校", 
            "毕业论文", "毕业设计", "国家开放大学", "学院", "专业",
            "年月日", "学生姓名", "导师", "指导老师", "作者：",
            "2024年", "2023年", "2025年",  # 年份
            "指导老", "学号：", "班级：", "分校：", "题目：",  # 带冒号的标签
            "开放大学", "教学中心",
            
            # Enhanced academic paper cover indicators
            "指导老师", "学生姓名", "专业班级", "提交日期",
            "毕业设计", "课程设计", "学位论文", "开题报告",
            "学士学位", "硕士学位", "博士学位", "研究生",
            "本科生", "学位论文", "毕业答辩", "论文题目",
            "研究方向", "所在学院", "所在系", "研究生院",
            "答辩委员会", "评阅教师", "论文作者", "完成时间",
            "学科专业", "研究领域", "学位类型", "培养单位",
            
            # International academic indicators
            "thesis", "dissertation", "supervisor", "advisor",
            "department", "university", "college", "faculty",
            "degree", "bachelor", "master", "doctor", "phd",
            "submitted", "presented", "fulfillment", "requirements",
            
            # Chinese university/institution indicators
            "大学", "学院", "系", "专业", "班级", "届",
            "教授", "副教授", "讲师", "博导", "硕导",
            "研究所", "实验室", "中心", "院系",
            
            # Date and time indicators
            "年", "月", "日", "时间", "日期", "完成于",
            "提交于", "答辩时间", "完成时间"
        ]
        
        for keyword in cover_keywords:
            if keyword in text_lower:
                logger.debug(f"段落包含封面关键词 '{keyword}': {text_preview[:30]}")
                return True
        
        # Enhanced TOC keywords
        toc_keywords = [
            "目录", "contents", "目 录", "table of contents",
            "content", "索引", "index", "章节", "目次"
        ]
        for keyword in toc_keywords:
            if keyword in text_lower:
                logger.debug(f"段落包含目录关键词 '{keyword}': {text_preview[:30]}")
                return True
        
        # Enhanced style-based detection for cover elements
        cover_styles = [
            # Original styles
            "封面", "cover", "title", "标题页", "目录",
            
            # Enhanced cover styles
            "封面标题", "封面副标题", "封面信息", "cover title",
            "cover subtitle", "cover info", "title page", "front page",
            "document title", "paper title", "thesis title",
            "author info", "author name", "student info",
            "supervisor info", "institution info", "date info"
        ]
        style_lower = style_name.lower()
        for style in cover_styles:
            if style in style_lower:
                logger.debug(f"段落使用封面样式 '{style}': {text_preview[:30]}")
                return True
        
        # Enhanced text pattern recognition for academic papers
        import re
        
        # Check for student ID patterns (various formats)
        if re.match(r'.*\d{8,12}.*', text_preview.strip()):  # 8-12 digit student IDs
            logger.debug(f"段落包含学号格式: {text_preview}")
            return True
        
        # Check for date formats (multiple patterns)
        date_patterns = [
            r'.*\d{4}年\d{1,2}月.*',  # Chinese date format
            r'.*\d{4}/\d{1,2}/\d{1,2}.*',  # Slash date format
            r'.*\d{4}-\d{1,2}-\d{1,2}.*',  # Dash date format
            r'.*\d{1,2}/\d{1,2}/\d{4}.*',  # US date format
            r'.*\d{1,2}-\d{1,2}-\d{4}.*',  # Alternative dash format
            r'.*\d{4}\s*年.*',  # Year only in Chinese
            r'.*(january|february|march|april|may|june|july|august|september|october|november|december).*\d{4}.*'  # English months
        ]
        
        for pattern in date_patterns:
            if re.match(pattern, text_preview, re.IGNORECASE):
                logger.debug(f"段落包含日期格式: {text_preview}")
                return True
        
        # Check for colon-separated label patterns (common in cover pages)
        colon_patterns = [
            r'.*[：:]\s*$',  # Lines ending with colon (labels)
            r'^[^：:]*[：:][^：:]*$',  # Single colon pattern (label: value)
            r'.*姓名[：:].*', r'.*学号[：:].*', r'.*专业[：:].*',
            r'.*班级[：:].*', r'.*指导[：:].*', r'.*题目[：:].*'
        ]
        
        for pattern in colon_patterns:
            if re.match(pattern, text_preview):
                logger.debug(f"段落包含标签格式: {text_preview}")
                return True
        
        # Check for pure numeric content (likely student IDs, phone numbers, etc.)
        if text_preview.strip().isdigit() and len(text_preview.strip()) >= 6:
            logger.debug(f"段落为长数字（可能是学号或联系方式）: {text_preview}")
            return True
        
        # Check for email patterns (common on cover pages)
        if re.match(r'.*[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}.*', text_preview):
            logger.debug(f"段落包含邮箱地址: {text_preview}")
            return True
        
        # Check for phone number patterns
        phone_patterns = [
            r'.*\d{3}-\d{4}-\d{4}.*',  # xxx-xxxx-xxxx
            r'.*\d{3}\s\d{4}\s\d{4}.*',  # xxx xxxx xxxx
            r'.*\(\d{3}\)\s*\d{3}-\d{4}.*',  # (xxx) xxx-xxxx
            r'.*1[3-9]\d{9}.*'  # Chinese mobile numbers
        ]
        
        for pattern in phone_patterns:
            if re.match(pattern, text_preview):
                logger.debug(f"段落包含电话号码格式: {text_preview}")
                return True
        
        # Check for very short content that might be cover elements
        if len(text_preview.strip()) <= 3 and text_preview.strip():
            # Very short content on early pages is likely cover formatting
            logger.debug(f"段落内容很短，可能是封面元素: {text_preview}")
            return True
        
        # Check for centered or specially formatted text (common on covers)
        # This would need to be enhanced with actual paragraph formatting checks
        # For now, we use text patterns that suggest centered content
        centered_patterns = [
            r'^\s+.*\s+$',  # Text with leading and trailing spaces
            r'^[A-Z\s]+$',  # All caps text (often titles)
            r'^[一二三四五六七八九十]+、.*',  # Chinese numbered items
        ]
        
        for pattern in centered_patterns:
            if re.match(pattern, text_preview) and len(text_preview.strip()) < 50:
                logger.debug(f"段落可能是居中格式的封面内容: {text_preview}")
                return True
        
        return False
    
    def _process_shapes_with_cover_protection(self, doc, first_content_index):
        """Enhanced shape and text frame processing with comprehensive cover protection"""
        try:
            logger.info("处理文档中的形状和文本框...")
            
            shape_count = 0
            protected_count = 0
            processed_count = 0
            error_count = 0
            
            # Process all shapes in the document
            for shape in doc.Shapes:
                shape_count += 1
                
                # Skip shapes without text content
                if not hasattr(shape, 'TextFrame') or not shape.TextFrame.HasText:
                    logger.debug(f"跳过无文本内容的形状 (类型: {getattr(shape, 'Type', 'Unknown')})")
                    continue
                
                try:
                    # Enhanced anchor page detection using Word COM Information property
                    anchor_page = None
                    try:
                        # Primary method: Use Information property to get anchor page
                        anchor_page = shape.Anchor.Information(3)  # wdActiveEndPageNumber
                        logger.debug(f"形状锚定页码: {anchor_page}")
                    except Exception as e:
                        logger.debug(f"无法获取形状锚定页码，尝试备用方法: {e}")
                        
                        # Fallback method: Try to get page from anchor range
                        try:
                            anchor_page = shape.Anchor.Range.Information(3)
                        except:
                            # Last resort: Check if shape is in first few paragraphs
                            try:
                                anchor_start = shape.Anchor.Start
                                # If anchor is in first 1000 characters, likely on cover page
                                if anchor_start < 1000:
                                    anchor_page = 1
                                else:
                                    anchor_page = 2  # Assume non-cover page
                            except:
                                anchor_page = 1  # Default to cover page for safety
                    
                    # Enhanced filtering to skip shapes anchored to cover pages
                    if anchor_page == 1:
                        logger.debug(f"🛡️ 跳过封面页形状 (锚定页码: {anchor_page})")
                        protected_count += 1
                        continue
                    
                    # Additional cover content detection based on shape text
                    try:
                        shape_text = shape.TextFrame.TextRange.Text.strip()
                        
                        # Use enhanced cover detection on shape text content
                        if self._is_cover_or_toc_content(0, first_content_index, shape_text, ""):
                            logger.debug(f"🛡️ 跳过包含封面内容的形状: {shape_text[:30]}...")
                            protected_count += 1
                            continue
                        
                        # Additional check for cover-specific text patterns in shapes
                        if self._is_shape_cover_content(shape_text):
                            logger.debug(f"🛡️ 跳过封面特征形状: {shape_text[:30]}...")
                            protected_count += 1
                            continue
                            
                    except Exception as e:
                        logger.debug(f"形状文本检查失败: {e}")
                        # If we can't read the text, be conservative and skip if on page 1
                        if anchor_page == 1:
                            protected_count += 1
                            continue
                    
                    # Enhanced text frame paragraph style reassignment
                    paragraph_count = 0
                    reassigned_count = 0
                    
                    try:
                        # Process each paragraph in the text frame
                        for paragraph in shape.TextFrame.TextRange.Paragraphs:
                            paragraph_count += 1
                            
                            try:
                                style_name = str(paragraph.Range.Style.NameLocal)
                                para_text = paragraph.Range.Text.strip()
                                
                                # Skip empty paragraphs
                                if not para_text:
                                    continue
                                
                                logger.debug(f"处理形状段落: 样式='{style_name}', 文本='{para_text[:20]}...'")
                                
                                # Enhanced paragraph style reassignment logic
                                if style_name in ["Normal", "正文", "Body Text", "正文文本"]:
                                    # Try to reassign to BodyText (AutoWord) style if it exists
                                    if self._body_text_style_exists(doc):
                                        try:
                                            paragraph.Range.Style = doc.Styles("BodyText (AutoWord)")
                                            logger.debug(f"✅ 形状段落已重新分配到BodyText样式: {para_text[:20]}...")
                                            reassigned_count += 1
                                        except Exception as e:
                                            logger.warning(f"形状段落BodyText样式分配失败: {e}")
                                            # Fallback to direct formatting
                                            self._apply_direct_formatting_to_shape_paragraph(paragraph)
                                            reassigned_count += 1
                                    else:
                                        # Apply direct formatting if BodyText style doesn't exist
                                        self._apply_direct_formatting_to_shape_paragraph(paragraph)
                                        reassigned_count += 1
                                        logger.debug(f"✅ 形状段落已应用直接格式: {para_text[:20]}...")
                                
                                # Handle other common styles that might need formatting
                                elif style_name in ["默认段落字体", "Default Paragraph Font"]:
                                    # These might also need formatting
                                    if self._body_text_style_exists(doc):
                                        try:
                                            paragraph.Range.Style = doc.Styles("BodyText (AutoWord)")
                                            reassigned_count += 1
                                        except:
                                            self._apply_direct_formatting_to_shape_paragraph(paragraph)
                                            reassigned_count += 1
                                
                            except Exception as e:
                                logger.warning(f"形状段落处理失败: {e}")
                                error_count += 1
                                continue
                        
                        if reassigned_count > 0:
                            processed_count += 1
                            logger.debug(f"形状处理完成: {reassigned_count}/{paragraph_count} 段落已处理")
                        
                    except Exception as e:
                        logger.warning(f"形状文本框段落遍历失败: {e}")
                        error_count += 1
                        continue
                        
                except Exception as e:
                    logger.warning(f"形状处理失败: {e}")
                    error_count += 1
                    continue
            
            # Enhanced logging with detailed statistics
            logger.info(f"📊 形状处理统计:")
            logger.info(f"  - 总形状数: {shape_count}")
            logger.info(f"  - 保护的形状: {protected_count} (封面/目录)")
            logger.info(f"  - 处理的形状: {processed_count}")
            logger.info(f"  - 处理错误: {error_count}")
            
            if processed_count > 0:
                logger.info(f"✅ 成功处理 {processed_count} 个形状中的文本框内容")
            if protected_count > 0:
                logger.info(f"🛡️ 成功保护 {protected_count} 个封面/目录形状")
            
        except Exception as e:
            logger.error(f"❌ 形状处理过程失败: {e}")
    
    def _is_shape_cover_content(self, shape_text):
        """Check if shape contains cover-specific content patterns"""
        if not shape_text or len(shape_text.strip()) == 0:
            return False
        
        text_lower = shape_text.lower().strip()
        
        # Shape-specific cover indicators (often found in text boxes on covers)
        shape_cover_patterns = [
            # Institution logos and names
            "大学", "学院", "university", "college", "institute",
            "school", "department", "faculty", "研究所",
            
            # Thesis/paper identifiers in text boxes
            "毕业论文", "毕业设计", "学位论文", "课程设计",
            "thesis", "dissertation", "project", "research",
            
            # Author/student info boxes
            "作者", "学生", "姓名", "author", "student", "name",
            "指导教师", "导师", "supervisor", "advisor",
            
            # Date/time info boxes
            "年", "月", "日", "时间", "日期", "date", "time",
            "提交", "完成", "submitted", "completed",
            
            # Academic info boxes
            "专业", "班级", "学号", "major", "class", "id",
            "学位", "degree", "bachelor", "master", "doctor",
            
            # Cover page decorative elements
            "封面", "cover", "title page", "front page"
        ]
        
        for pattern in shape_cover_patterns:
            if pattern in text_lower:
                return True
        
        # Check for typical cover page text box content (short, formal text)
        if len(text_lower) < 50 and any(char in text_lower for char in ["：", ":", "年", "月", "日"]):
            return True
        
        return False
    
    def _apply_direct_formatting_to_shape_paragraph(self, paragraph):
        """Apply direct formatting to shape paragraph when style reassignment fails"""
        try:
            # Apply standard body text formatting
            paragraph.Range.Font.NameFarEast = "宋体"
            paragraph.Range.Font.Name = "Times New Roman"  # For Latin text
            paragraph.Range.Font.Size = 12
            paragraph.Range.Font.Bold = False
            paragraph.Range.ParagraphFormat.LineSpacing = 24  # 2倍行距
            paragraph.Range.ParagraphFormat.SpaceAfter = 0
            paragraph.Range.ParagraphFormat.SpaceBefore = 0
            
        except Exception as e:
            logger.warning(f"形状段落直接格式化失败: {e}")
    
    def _apply_styles_to_content(self, doc):
        """Enhanced style application with paragraph reassignment and cover protection"""
        try:
            logger.info("强制应用样式到文档内容...")
            
            # 首先找到第一个正文节的位置
            first_content_index = self._find_first_content_section(doc)
            logger.info(f"正文开始位置: 段落索引 {first_content_index}")
            
            # Process shapes with cover protection
            self._process_shapes_with_cover_protection(doc, first_content_index)
            
            # NEW: Enhanced paragraph reassignment logic
            reassignment_count = 0
            protected_count = 0
            
            # 遍历所有段落，基于outline level识别和应用样式
            for i, para in enumerate(doc.Paragraphs):
                try:
                    # 获取段落的大纲级别和页码信息
                    outline_level = para.OutlineLevel
                    style_name = para.Style.NameLocal
                    text_preview = para.Range.Text.strip()[:30]
                    
                    # 获取段落所在页码（可能为0，表示无页码）
                    try:
                        page_number = para.Range.Information(3)  # wdActiveEndPageNumber
                    except:
                        page_number = 0  # 无页码
                    
                    # Enhanced cover/TOC detection using existing _is_cover_or_toc_content filtering
                    is_cover_or_toc = self._is_cover_or_toc_content(i, first_content_index, text_preview, style_name)
                    if is_cover_or_toc:
                        # 封面或目录内容，跳过格式应用和段落重新分配
                        if text_preview and len(text_preview) > 2:  # 只对有实际内容的段落记录
                            logger.debug(f"🛡️ 保护封面/目录内容: {text_preview}... (段落={i}, 页码={page_number})")
                            protected_count += 1
                        continue
                    
                    # NEW: Enhanced paragraph reassignment logic for main content
                    # Reassign Normal/正文 paragraphs to BodyText (AutoWord) style if it exists
                    if style_name in ["Normal", "正文"] and self._body_text_style_exists(doc):
                        try:
                            para.Range.Style = doc.Styles("BodyText (AutoWord)")
                            logger.debug(f"段落已重新分配到BodyText样式: {text_preview}... (段落={i})")
                            reassignment_count += 1
                        except Exception as e:
                            logger.warning(f"段落重新分配失败 (段落={i}): {e}")
                            # Fallback to direct formatting
                            para.Range.Font.NameFarEast = "宋体"
                            para.Range.Font.Size = 12
                            para.Range.Font.Bold = False
                            para.Range.ParagraphFormat.LineSpacing = 24
                        continue  # Skip further processing for reassigned paragraphs
                    
                    # 基于outline level判断标题级别
                    if outline_level == 1:  # 1级标题
                        para.Range.Font.NameFarEast = "楷体"
                        para.Range.Font.Size = 12  # 小四
                        para.Range.Font.Bold = True
                        para.Range.ParagraphFormat.LineSpacing = 24  # 2倍行距
                        logger.debug(f"应用1级标题格式到: {text_preview}... (outline_level={outline_level})")
                        
                    elif outline_level == 2:  # 2级标题
                        para.Range.Font.NameFarEast = "宋体"
                        para.Range.Font.Size = 12  # 小四
                        para.Range.Font.Bold = True
                        para.Range.ParagraphFormat.LineSpacing = 24  # 2倍行距
                        logger.debug(f"应用2级标题格式到: {text_preview}... (outline_level={outline_level})")
                        
                    elif outline_level == 3:  # 3级标题
                        para.Range.Font.NameFarEast = "宋体"
                        para.Range.Font.Size = 12  # 小四
                        para.Range.Font.Bold = True
                        para.Range.ParagraphFormat.LineSpacing = 24  # 2倍行距
                        logger.debug(f"应用3级标题格式到: {text_preview}... (outline_level={outline_level})")
                        
                    elif outline_level == 10 or outline_level == 0:  # 正文级别
                        # 再次确认不是封面内容（双重保护）
                        if page_number == 1:
                            logger.debug(f"跳过封面正文内容: {text_preview}... (page={page_number}, outline_level={outline_level})")
                            protected_count += 1
                            continue
                            
                        # Apply formatting to body text paragraphs
                        if self._body_text_style_exists(doc):
                            try:
                                para.Range.Style = doc.Styles("BodyText (AutoWord)")
                                logger.debug(f"正文段落已重新分配到BodyText样式: {text_preview}... (outline_level={outline_level})")
                                reassignment_count += 1
                            except Exception as e:
                                logger.warning(f"正文段落重新分配失败: {e}")
                                # Fallback to direct formatting
                                para.Range.Font.NameFarEast = "宋体"
                                para.Range.Font.Size = 12  # 小四
                                para.Range.Font.Bold = False
                                para.Range.ParagraphFormat.LineSpacing = 24  # 2倍行距
                        else:
                            # Original behavior when BodyText style doesn't exist
                            para.Range.Font.NameFarEast = "宋体"
                            para.Range.Font.Size = 12  # 小四
                            para.Range.Font.Bold = False
                            para.Range.ParagraphFormat.LineSpacing = 24  # 2倍行距
                    
                    # 如果outline level不明确，回退到样式名判断
                    elif "标题" in style_name or "Heading" in style_name:
                        if "1" in style_name:
                            para.Range.Font.NameFarEast = "楷体"
                            para.Range.Font.Size = 12
                            para.Range.Font.Bold = True
                            para.Range.ParagraphFormat.LineSpacing = 24
                            logger.debug(f"基于样式应用1级标题格式到: {text_preview}...")
                        elif "2" in style_name:
                            para.Range.Font.NameFarEast = "宋体"
                            para.Range.Font.Size = 12
                            para.Range.Font.Bold = True
                            para.Range.ParagraphFormat.LineSpacing = 24
                            logger.debug(f"基于样式应用2级标题格式到: {text_preview}...")
                        
                except Exception as e:
                    # 单个段落处理失败不影响其他段落
                    logger.warning(f"段落处理失败 (段落={i}): {e}")
                    continue
            
            # Enhanced logging for reassignment operations
            logger.info(f"样式应用完成 - 重新分配段落: {reassignment_count}, 保护段落: {protected_count}")
            
        except Exception as e:
            logger.warning(f"强制应用样式失败: {e}")
    
    def _validate_result(self, result_path: str) -> bool:
        """Enhanced validation with cover format checking"""
        try:
            # Basic validation: check file exists and has reasonable size
            if not os.path.exists(result_path):
                logger.warning("验证失败: 输出文件不存在")
                return False
            
            file_size = os.path.getsize(result_path)
            if file_size < 1000:  # File too small might indicate problems
                logger.warning(f"验证失败: 文件大小过小 ({file_size} bytes)")
                return False
            
            # Enhanced validation: check cover page formatting preservation
            cover_validation_passed = self._validate_cover_formatting(result_path)
            
            if not cover_validation_passed:
                logger.warning("⚠️ 封面格式验证失败，但文档已生成")
                # Note: We don't return False here as the document is still usable
                # The warning will be captured in the ProcessingResult
            
            return True
            
        except Exception as e:
            logger.warning(f"验证失败: {e}")
            return False
    
    def _validate_cover_formatting(self, result_path: str) -> bool:
        """Enhanced validation with before/after cover format comparison"""
        try:
            import win32com.client
            
            logger.info("开始验证封面格式保护...")
            
            # Initialize validation warnings list
            self._cover_validation_warnings = []
            
            # Check if we have original cover format for comparison
            if not hasattr(self, '_original_cover_format') or not self._original_cover_format:
                logger.warning("缺少原始封面格式信息，使用基本验证")
                return self._validate_cover_formatting_basic(result_path)
            
            # Open the processed document for validation
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            
            try:
                abs_result_path = os.path.abspath(result_path)
                doc = word.Documents.Open(abs_result_path)
                
                # Collect current cover page formatting information
                current_cover_format = []
                cover_format_issues = []
                cover_paragraphs_checked = 0
                
                # Check first page paragraphs (cover page)
                for i, para in enumerate(doc.Paragraphs):
                    try:
                        # Get paragraph page number
                        page_number = para.Range.Information(3)  # wdActiveEndPageNumber
                        
                        # Only check first page (cover page)
                        if page_number != 1:
                            break
                        
                        text_preview = para.Range.Text.strip()[:100]
                        if not text_preview or len(text_preview) < 2:
                            continue  # Skip empty or very short paragraphs
                        
                        cover_paragraphs_checked += 1
                        
                        # Check if this is cover content using existing detection logic
                        style_name = para.Style.NameLocal
                        is_cover_content = self._is_cover_or_toc_content(i, 999, text_preview, style_name)
                        
                        if is_cover_content:
                            # Capture current formatting for comparison
                            current_format = {
                                "index": i,
                                "text_preview": text_preview,
                                "style_name": style_name,
                                "font_name_east_asian": para.Range.Font.NameFarEast,
                                "font_name_latin": para.Range.Font.Name,
                                "font_size": para.Range.Font.Size,
                                "font_bold": para.Range.Font.Bold,
                                "font_italic": para.Range.Font.Italic,
                                "line_spacing": para.Range.ParagraphFormat.LineSpacing,
                                "space_before": para.Range.ParagraphFormat.SpaceBefore,
                                "space_after": para.Range.ParagraphFormat.SpaceAfter,
                                "alignment": para.Range.ParagraphFormat.Alignment,
                                "left_indent": para.Range.ParagraphFormat.LeftIndent,
                                "page_number": page_number
                            }
                            
                            current_cover_format.append(current_format)
                            
                            # Find matching original paragraph for comparison
                            original_format = self._find_matching_original_paragraph(text_preview, i)
                            
                            if original_format:
                                # Compare before/after formatting
                                format_changes = self._compare_paragraph_formatting(original_format, current_format)
                                
                                if format_changes:
                                    issue_summary = f"段落 {i}: '{text_preview[:30]}...' - " + "; ".join(format_changes)
                                    cover_format_issues.append(issue_summary)
                                    logger.warning(f"🚨 封面格式变化: {issue_summary}")
                            else:
                                # No original format found, use basic validation
                                basic_issues = self._validate_paragraph_formatting_basic(current_format)
                                if basic_issues:
                                    issue_summary = f"段落 {i}: '{text_preview[:30]}...' - " + "; ".join(basic_issues)
                                    cover_format_issues.append(issue_summary)
                                    logger.warning(f"🚨 封面格式问题: {issue_summary}")
                        
                        # Limit checking to first 20 paragraphs to avoid performance issues
                        if i >= 20:
                            break
                            
                    except Exception as e:
                        logger.debug(f"段落 {i} 格式检查失败: {e}")
                        continue
                
                doc.Close()
                
                # Analyze validation results
                return self._analyze_cover_validation_results(cover_format_issues, cover_paragraphs_checked)
                
            finally:
                word.Quit()
                
        except Exception as e:
            logger.warning(f"封面格式验证失败: {e}")
            # Don't fail the entire validation if cover checking fails
            return True
    
    def _find_matching_original_paragraph(self, text_preview: str, current_index: int) -> Dict[str, Any]:
        """Find matching paragraph in original cover format"""
        if not self._original_cover_format or "paragraphs" not in self._original_cover_format:
            return None
        
        # First try to match by text content
        for orig_para in self._original_cover_format["paragraphs"]:
            if orig_para["text_preview"][:50] == text_preview[:50]:
                return orig_para
        
        # If no text match, try to match by index (less reliable but better than nothing)
        for orig_para in self._original_cover_format["paragraphs"]:
            if orig_para["index"] == current_index:
                return orig_para
        
        return None
    
    def _compare_paragraph_formatting(self, original: Dict[str, Any], current: Dict[str, Any]) -> list:
        """Compare original and current paragraph formatting, return list of changes"""
        changes = []
        
        # Check style changes
        if original["style_name"] != current["style_name"]:
            if current["style_name"] == "BodyText (AutoWord)":
                changes.append(f"样式被错误修改为BodyText (AutoWord)，原样式: {original['style_name']}")
            else:
                changes.append(f"样式从 {original['style_name']} 改为 {current['style_name']}")
        
        # Check font changes
        if original["font_name_east_asian"] != current["font_name_east_asian"]:
            if current["font_name_east_asian"] == "宋体" and original["font_name_east_asian"] != "宋体":
                changes.append(f"中文字体被意外修改为宋体，原字体: {original['font_name_east_asian']}")
            else:
                changes.append(f"中文字体从 {original['font_name_east_asian']} 改为 {current['font_name_east_asian']}")
        
        # Check font size changes
        if original["font_size"] != current["font_size"]:
            if current["font_size"] == 12 and original["font_size"] != 12:
                changes.append(f"字体大小被意外修改为12pt，原大小: {original['font_size']}pt")
            else:
                changes.append(f"字体大小从 {original['font_size']}pt 改为 {current['font_size']}pt")
        
        # Check line spacing changes
        if original["line_spacing"] != current["line_spacing"]:
            if current["line_spacing"] == 24 and original["line_spacing"] != 24:
                changes.append(f"行距被意外修改为2倍行距(24pt)，原行距: {original['line_spacing']}pt")
            else:
                changes.append(f"行距从 {original['line_spacing']}pt 改为 {current['line_spacing']}pt")
        
        # Check bold changes
        if original["font_bold"] != current["font_bold"]:
            if current["font_bold"] and not original["font_bold"]:
                changes.append("字体被意外设置为粗体")
            elif not current["font_bold"] and original["font_bold"]:
                changes.append("字体粗体被意外取消")
        
        return changes
    
    def _validate_paragraph_formatting_basic(self, current_format: Dict[str, Any]) -> list:
        """Basic validation when no original format is available"""
        issues = []
        text_preview = current_format["text_preview"].lower()
        
        # Check if paragraph was incorrectly assigned to BodyText style
        if current_format["style_name"] == "BodyText (AutoWord)":
            issues.append("封面段落被错误分配到BodyText样式")
        
        # Check for unexpected font changes (common signs of corruption)
        if current_format["font_name_east_asian"] == "宋体" and "标题" not in current_format["style_name"]:
            # Cover content shouldn't be forced to 宋体 unless it's intentional
            if any(keyword in text_preview for keyword in [
                "题目", "姓名", "学号", "指导", "大学", "学院", "专业", "班级"
            ]):
                issues.append("封面信息字体可能被意外修改为宋体")
        
        # Check for unexpected line spacing (24pt = 2倍行距)
        if current_format["line_spacing"] == 24:
            # Cover content shouldn't have forced 2倍行距 unless intentional
            if any(keyword in text_preview for keyword in [
                "题目", "姓名", "学号", "指导", "大学", "学院"
            ]):
                issues.append("封面信息行距可能被意外修改为2倍行距")
        
        # Check for unexpected font size changes
        if current_format["font_size"] == 12:
            # Cover titles shouldn't be forced to 12pt unless intentional
            if any(keyword in text_preview for keyword in [
                "毕业论文", "毕业设计", "题目", "大学"
            ]) and len(current_format["text_preview"]) > 10:
                issues.append("封面标题字体大小可能被意外修改为12pt")
        
        return issues
    
    def _validate_cover_formatting_basic(self, result_path: str) -> bool:
        """Basic cover format validation when no original format is available"""
        try:
            import win32com.client
            
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            
            try:
                abs_result_path = os.path.abspath(result_path)
                doc = word.Documents.Open(abs_result_path)
                
                cover_format_issues = []
                cover_paragraphs_checked = 0
                
                # Check first page paragraphs (cover page)
                for i, para in enumerate(doc.Paragraphs):
                    try:
                        page_number = para.Range.Information(3)  # wdActiveEndPageNumber
                        
                        if page_number != 1:
                            break
                        
                        text_preview = para.Range.Text.strip()[:50]
                        if not text_preview or len(text_preview) < 2:
                            continue
                        
                        cover_paragraphs_checked += 1
                        
                        style_name = para.Style.NameLocal
                        is_cover_content = self._is_cover_or_toc_content(i, 999, text_preview, style_name)
                        
                        if is_cover_content:
                            current_format = {
                                "text_preview": text_preview,
                                "style_name": style_name,
                                "font_name_east_asian": para.Range.Font.NameFarEast,
                                "font_size": para.Range.Font.Size,
                                "line_spacing": para.Range.ParagraphFormat.LineSpacing
                            }
                            
                            basic_issues = self._validate_paragraph_formatting_basic(current_format)
                            if basic_issues:
                                issue_summary = f"段落 {i}: '{text_preview}' - " + "; ".join(basic_issues)
                                cover_format_issues.append(issue_summary)
                                logger.warning(f"🚨 封面格式问题: {issue_summary}")
                        
                        if i >= 20:
                            break
                            
                    except Exception as e:
                        logger.debug(f"段落 {i} 格式检查失败: {e}")
                        continue
                
                doc.Close()
                
                return self._analyze_cover_validation_results(cover_format_issues, cover_paragraphs_checked)
                
            finally:
                word.Quit()
                
        except Exception as e:
            logger.warning(f"基本封面格式验证失败: {e}")
            return True
    
    def _analyze_cover_validation_results(self, cover_format_issues: list, cover_paragraphs_checked: int) -> bool:
        """Analyze cover validation results and determine if validation passes"""
        logger.info(f"封面格式验证完成 - 检查段落: {cover_paragraphs_checked}, 发现问题: {len(cover_format_issues)}")
        
        # Initialize warnings list if not exists
        if not hasattr(self, '_cover_validation_warnings'):
            self._cover_validation_warnings = []
        
        if cover_format_issues:
            logger.warning("🚨 检测到封面格式问题:")
            for issue in cover_format_issues[:5]:  # Limit to first 5 issues to avoid log spam
                logger.warning(f"  - {issue}")
                self._cover_validation_warnings.append(f"封面格式问题: {issue}")
            
            if len(cover_format_issues) > 5:
                logger.warning(f"  - ... 还有 {len(cover_format_issues) - 5} 个其他问题")
                self._cover_validation_warnings.append(f"还有 {len(cover_format_issues) - 5} 个其他封面格式问题")
            
            # Check if issues are severe enough to trigger rollback warning
            severe_issues = [issue for issue in cover_format_issues if 
                           "错误分配到BodyText样式" in issue or "意外修改" in issue]
            
            if len(severe_issues) >= 3:
                logger.error("🚨 严重封面格式问题检测到！建议检查封面保护逻辑")
                logger.error("💡 提示: 如果封面格式严重损坏，可以考虑使用原始文档重新处理")
                self._cover_validation_warnings.append("严重封面格式问题检测到，建议重新处理")
                return False
            elif len(cover_format_issues) >= 5:
                logger.warning("⚠️ 多个封面格式问题检测到，请检查结果")
                self._cover_validation_warnings.append("多个封面格式问题检测到")
                return False
            else:
                logger.info("✅ 封面格式问题较少，在可接受范围内")
                return True
        else:
            logger.info("✅ 封面格式验证通过，未发现问题")
            return True
    
    def _capture_cover_formatting(self, docx_path: str) -> Dict[str, Any]:
        """Capture original cover page formatting for before/after comparison"""
        try:
            import win32com.client
            
            logger.info("捕获原始封面格式信息...")
            
            # Convert to absolute path to avoid Word COM path issues
            abs_path = os.path.abspath(docx_path)
            logger.debug(f"使用绝对路径: {abs_path}")
            
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            
            cover_format = {
                "paragraphs": [],
                "capture_time": time.time(),
                "document_path": abs_path
            }
            
            try:
                doc = word.Documents.Open(abs_path)
                
                # Capture formatting of first page paragraphs (cover page)
                for i, para in enumerate(doc.Paragraphs):
                    try:
                        # Get paragraph page number
                        page_number = para.Range.Information(3)  # wdActiveEndPageNumber
                        
                        # Only capture first page (cover page)
                        if page_number != 1:
                            break
                        
                        text_preview = para.Range.Text.strip()[:100]
                        if not text_preview or len(text_preview) < 2:
                            continue  # Skip empty or very short paragraphs
                        
                        # Check if this is cover content
                        style_name = para.Style.NameLocal
                        is_cover_content = self._is_cover_or_toc_content(i, 999, text_preview, style_name)
                        
                        if is_cover_content:
                            # Capture detailed formatting information
                            para_format = {
                                "index": i,
                                "text_preview": text_preview,
                                "style_name": style_name,
                                "font_name_east_asian": para.Range.Font.NameFarEast,
                                "font_name_latin": para.Range.Font.Name,
                                "font_size": para.Range.Font.Size,
                                "font_bold": para.Range.Font.Bold,
                                "font_italic": para.Range.Font.Italic,
                                "line_spacing": para.Range.ParagraphFormat.LineSpacing,
                                "space_before": para.Range.ParagraphFormat.SpaceBefore,
                                "space_after": para.Range.ParagraphFormat.SpaceAfter,
                                "alignment": para.Range.ParagraphFormat.Alignment,
                                "left_indent": para.Range.ParagraphFormat.LeftIndent,
                                "page_number": page_number
                            }
                            
                            cover_format["paragraphs"].append(para_format)
                            logger.debug(f"捕获封面段落格式: {text_preview[:30]}...")
                        
                        # Limit to first 20 paragraphs to avoid performance issues
                        if i >= 20:
                            break
                            
                    except Exception as e:
                        logger.debug(f"段落 {i} 格式捕获失败: {e}")
                        continue
                
                doc.Close()
                
                logger.info(f"原始封面格式捕获完成 - 捕获段落: {len(cover_format['paragraphs'])}")
                return cover_format
                
            finally:
                word.Quit()
                
        except Exception as e:
            logger.warning(f"原始封面格式捕获失败: {e}")
            return {"paragraphs": [], "capture_time": time.time(), "error": str(e)}

# 为了兼容性，创建VNextPipeline别名
VNextPipeline = SimplePipeline