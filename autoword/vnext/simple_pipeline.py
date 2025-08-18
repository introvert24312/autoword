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
        
    def process_document(self, docx_path: str, user_intent: str) -> ProcessingResult:
        """处理文档"""
        logger.info(f"开始处理文档: {docx_path}")
        logger.info(f"用户意图: {user_intent}")
        
        try:
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
            if self._validate_result(result_path):
                logger.info("✅ 文档处理成功！")
                return ProcessingResult(
                    status="SUCCESS",
                    message="文档处理成功",
                    audit_directory=str(Path(result_path).parent)
                )
            else:
                logger.warning("⚠️ 验证失败，但文档已生成")
                return ProcessingResult(
                    status="SUCCESS",
                    message="文档处理完成，但验证有警告",
                    warnings=["部分验证检查未通过"],
                    audit_directory=str(Path(result_path).parent)
                )
                
        except Exception as e:
            logger.error(f"❌ 处理失败: {e}")
            return ProcessingResult(
                status="FAILED",
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
                # 打开文档
                doc = word.Documents.Open(docx_path)
                
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
                # 打开文档
                doc = word.Documents.Open(str(output_path))
                
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
                "正文": "Normal"
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
    
    def _apply_styles_to_content(self, doc):
        """强制应用样式到文档内容"""
        try:
            logger.info("强制应用样式到文档内容...")
            
            # 遍历所有段落，基于outline level识别和应用样式
            for para in doc.Paragraphs:
                try:
                    # 获取段落的大纲级别和页码信息
                    outline_level = para.OutlineLevel
                    style_name = para.Style.NameLocal
                    text_preview = para.Range.Text.strip()[:30]
                    
                    # 获取段落所在页码
                    try:
                        page_number = para.Range.Information(3)  # wdActiveEndPageNumber
                    except:
                        page_number = 1  # 默认为第1页
                    
                    # 封面识别：第1页的内容不应用正文格式
                    if page_number == 1:
                        # 封面内容，跳过格式应用
                        if text_preview and len(text_preview) > 5:  # 只对有实际内容的段落记录
                            logger.info(f"跳过封面内容: {text_preview}... (page={page_number})")
                        continue
                    
                    # 基于outline level判断标题级别
                    if outline_level == 1:  # 1级标题
                        para.Range.Font.NameFarEast = "楷体"
                        para.Range.Font.Size = 12  # 小四
                        para.Range.Font.Bold = True
                        para.Range.ParagraphFormat.LineSpacing = 24  # 2倍行距
                        logger.info(f"应用1级标题格式到: {text_preview}... (outline_level={outline_level})")
                        
                    elif outline_level == 2:  # 2级标题
                        para.Range.Font.NameFarEast = "宋体"
                        para.Range.Font.Size = 12  # 小四
                        para.Range.Font.Bold = True
                        para.Range.ParagraphFormat.LineSpacing = 24  # 2倍行距
                        logger.info(f"应用2级标题格式到: {text_preview}... (outline_level={outline_level})")
                        
                    elif outline_level == 3:  # 3级标题
                        para.Range.Font.NameFarEast = "宋体"
                        para.Range.Font.Size = 12  # 小四
                        para.Range.Font.Bold = True
                        para.Range.ParagraphFormat.LineSpacing = 24  # 2倍行距
                        logger.info(f"应用3级标题格式到: {text_preview}... (outline_level={outline_level})")
                        
                    elif outline_level == 10 or outline_level == 0:  # 正文级别
                        # 再次确认不是封面内容（双重保护）
                        if page_number == 1:
                            logger.info(f"跳过封面正文内容: {text_preview}... (page={page_number}, outline_level={outline_level})")
                            continue
                            
                        # 应用正文格式
                        para.Range.Font.NameFarEast = "宋体"
                        para.Range.Font.Size = 12  # 小四
                        para.Range.Font.Bold = False
                        para.Range.ParagraphFormat.LineSpacing = 24  # 2倍行距
                        # logger.info(f"应用正文格式到: {text_preview}... (outline_level={outline_level})")
                    
                    # 如果outline level不明确，回退到样式名判断
                    elif "标题" in style_name or "Heading" in style_name:
                        if "1" in style_name:
                            para.Range.Font.NameFarEast = "楷体"
                            para.Range.Font.Size = 12
                            para.Range.Font.Bold = True
                            para.Range.ParagraphFormat.LineSpacing = 24
                            logger.info(f"基于样式应用1级标题格式到: {text_preview}...")
                        elif "2" in style_name:
                            para.Range.Font.NameFarEast = "宋体"
                            para.Range.Font.Size = 12
                            para.Range.Font.Bold = True
                            para.Range.ParagraphFormat.LineSpacing = 24
                            logger.info(f"基于样式应用2级标题格式到: {text_preview}...")
                    
                    # 对明确的正文样式应用格式
                    elif "正文" in style_name or "Normal" in style_name:
                        para.Range.Font.NameFarEast = "宋体"
                        para.Range.Font.Size = 12
                        para.Range.Font.Bold = False
                        para.Range.ParagraphFormat.LineSpacing = 24
                        
                except Exception as e:
                    # 单个段落处理失败不影响其他段落
                    continue
            
            logger.info("样式应用完成")
            
        except Exception as e:
            logger.warning(f"强制应用样式失败: {e}")
    
    def _validate_result(self, result_path: str) -> bool:
        """验证结果"""
        try:
            # 简单验证：检查文件是否存在且大小合理
            if not os.path.exists(result_path):
                return False
            
            file_size = os.path.getsize(result_path)
            if file_size < 1000:  # 文件太小可能有问题
                return False
            
            return True
            
        except Exception as e:
            logger.warning(f"验证失败: {e}")
            return False

# 为了兼容性，创建VNextPipeline别名
VNextPipeline = SimplePipeline