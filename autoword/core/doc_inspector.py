"""
AutoWord Document Inspector
Word 文档检查器，提取批注和文档结构
"""

import logging
from typing import List, Any, Optional
from datetime import datetime

import win32com.client as win32
from win32com.client import constants as win32_constants

from .models import (
    Comment, Heading, Style, TocEntry, Hyperlink, Reference,
    DocumentStructure, DocumentSnapshot
)
from .utils import calculate_file_checksum, truncate_text
from .exceptions import DocumentError, COMError


logger = logging.getLogger(__name__)


class DocInspector:
    """Word 文档检查器"""
    
    def __init__(self):
        """初始化文档检查器"""
        pass
    
    def extract_comments(self, document: Any) -> List[Comment]:
        """
        提取文档中的所有批注
        
        Args:
            document: Word 文档对象
            
        Returns:
            批注列表
            
        Raises:
            DocumentError: 提取失败时抛出
        """
        try:
            comments = []
            
            # 遍历文档中的所有批注
            for i, com in enumerate(document.Comments, 1):
                try:
                    # 获取批注基本信息
                    comment_id = f"comment_{i}"
                    author = getattr(com, 'Author', 'Unknown')
                    
                    # 获取批注内容
                    comment_text = com.Range.Text.strip() if com.Range else ""
                    
                    # 获取批注锚点范围
                    scope = com.Scope
                    range_start = scope.Start if scope else 0
                    range_end = scope.End if scope else 0
                    
                    # 获取锚点文本（批注所指向的文本）
                    anchor_text = ""
                    if scope and scope.Text:
                        anchor_text = truncate_text(scope.Text.strip(), max_length=100)
                    
                    # 获取页码
                    page_number = 1
                    try:
                        if scope:
                            page_number = scope.Information(win32_constants.wdActiveEndPageNumber)
                    except Exception:
                        logger.warning(f"Could not get page number for comment {i}")
                    
                    # 获取创建时间
                    created_date = None
                    try:
                        if hasattr(com, 'Date'):
                            created_date = com.Date
                    except Exception:
                        pass
                    
                    # 创建批注对象
                    comment = Comment(
                        id=comment_id,
                        author=author,
                        page=max(1, page_number),  # 确保页码至少为1
                        anchor_text=anchor_text,
                        comment_text=comment_text,
                        range_start=range_start,
                        range_end=range_end,
                        created_date=created_date
                    )
                    
                    comments.append(comment)
                    logger.debug(f"Extracted comment {i}: {author} - {truncate_text(comment_text, 50)}")
                    
                except Exception as e:
                    logger.warning(f"Failed to extract comment {i}: {e}")
                    continue
            
            logger.info(f"Extracted {len(comments)} comments")
            return comments
            
        except Exception as e:
            raise DocumentError(f"Failed to extract comments: {e}")
    
    def extract_headings(self, document: Any) -> List[Heading]:
        """
        提取文档标题
        
        Args:
            document: Word 文档对象
            
        Returns:
            标题列表
        """
        try:
            headings = []
            
            # 遍历文档中的所有段落
            for para in document.Paragraphs:
                try:
                    style_name = para.Style.NameLocal
                    
                    # 检查是否为标题样式
                    if self._is_heading_style(style_name):
                        level = self._get_heading_level(style_name)
                        if level > 0:
                            text = para.Range.Text.strip().replace('\r', '')
                            
                            if text:  # 忽略空标题
                                heading = Heading(
                                    level=level,
                                    text=text,
                                    style=style_name,
                                    range_start=para.Range.Start,
                                    range_end=para.Range.End
                                )
                                headings.append(heading)
                                
                except Exception as e:
                    logger.debug(f"Error processing paragraph for headings: {e}")
                    continue
            
            logger.info(f"Extracted {len(headings)} headings")
            return headings
            
        except Exception as e:
            logger.warning(f"Failed to extract headings: {e}")
            return []
    
    def extract_styles(self, document: Any) -> List[Style]:
        """
        提取文档样式
        
        Args:
            document: Word 文档对象
            
        Returns:
            样式列表
        """
        try:
            styles = []
            
            for style in document.Styles:
                try:
                    style_obj = Style(
                        name=style.NameLocal,
                        type=self._get_style_type(style.Type),
                        built_in=style.BuiltIn,
                        in_use=style.InUse
                    )
                    styles.append(style_obj)
                    
                except Exception as e:
                    logger.debug(f"Error processing style: {e}")
                    continue
            
            logger.info(f"Extracted {len(styles)} styles")
            return styles
            
        except Exception as e:
            logger.warning(f"Failed to extract styles: {e}")
            return []
    
    def extract_toc_entries(self, document: Any) -> List[TocEntry]:
        """
        提取目录条目
        
        Args:
            document: Word 文档对象
            
        Returns:
            目录条目列表
        """
        try:
            toc_entries = []
            
            # 查找目录
            for toc in document.TablesOfContents:
                try:
                    toc_range = toc.Range
                    
                    # 遍历目录中的超链接（目录条目通常是超链接）
                    for hyperlink in toc_range.Hyperlinks:
                        try:
                            text = hyperlink.TextToDisplay.strip()
                            if text:
                                # 尝试提取页码（通常在文本末尾）
                                page_number = self._extract_page_number_from_toc_text(text)
                                
                                # 估算目录级别（基于缩进或样式）
                                level = self._estimate_toc_level(hyperlink.Range)
                                
                                entry = TocEntry(
                                    level=level,
                                    text=text,
                                    page_number=page_number,
                                    range_start=hyperlink.Range.Start,
                                    range_end=hyperlink.Range.End
                                )
                                toc_entries.append(entry)
                                
                        except Exception as e:
                            logger.debug(f"Error processing TOC hyperlink: {e}")
                            continue
                            
                except Exception as e:
                    logger.debug(f"Error processing TOC: {e}")
                    continue
            
            logger.info(f"Extracted {len(toc_entries)} TOC entries")
            return toc_entries
            
        except Exception as e:
            logger.warning(f"Failed to extract TOC entries: {e}")
            return []
    
    def extract_hyperlinks(self, document: Any) -> List[Hyperlink]:
        """
        提取超链接
        
        Args:
            document: Word 文档对象
            
        Returns:
            超链接列表
        """
        try:
            hyperlinks = []
            
            for hyperlink in document.Hyperlinks:
                try:
                    text = hyperlink.TextToDisplay or ""
                    address = hyperlink.Address or ""
                    
                    # 确定链接类型
                    link_type = "external"
                    if not address:
                        if hyperlink.SubAddress:
                            link_type = "internal"
                            address = hyperlink.SubAddress
                    elif address.startswith("#"):
                        link_type = "internal"
                    elif address.startswith("mailto:"):
                        link_type = "email"
                    elif address.startswith(("http://", "https://")):
                        link_type = "web"
                    elif address.startswith("file://"):
                        link_type = "file"
                    
                    link = Hyperlink(
                        text=text,
                        address=address,
                        type=link_type,
                        range_start=hyperlink.Range.Start,
                        range_end=hyperlink.Range.End
                    )
                    hyperlinks.append(link)
                    
                except Exception as e:
                    logger.debug(f"Error processing hyperlink: {e}")
                    continue
            
            logger.info(f"Extracted {len(hyperlinks)} hyperlinks")
            return hyperlinks
            
        except Exception as e:
            logger.warning(f"Failed to extract hyperlinks: {e}")
            return []
    
    def extract_references(self, document: Any) -> List[Reference]:
        """
        提取引用（交叉引用、书签等）
        
        Args:
            document: Word 文档对象
            
        Returns:
            引用列表
        """
        try:
            references = []
            
            # 提取书签
            for bookmark in document.Bookmarks:
                try:
                    ref = Reference(
                        type="bookmark",
                        text=bookmark.Name,
                        target=bookmark.Name,
                        range_start=bookmark.Range.Start,
                        range_end=bookmark.Range.End
                    )
                    references.append(ref)
                    
                except Exception as e:
                    logger.debug(f"Error processing bookmark: {e}")
                    continue
            
            # 提取域（Fields）中的引用
            for field in document.Fields:
                try:
                    field_type = field.Type
                    field_code = field.Code.Text if field.Code else ""
                    
                    # 检查是否为引用类型的域
                    if field_type in [win32_constants.wdFieldRef, win32_constants.wdFieldPageRef]:
                        ref = Reference(
                            type="field_reference",
                            text=field.Result.Text if field.Result else "",
                            target=field_code,
                            range_start=field.Code.Start,
                            range_end=field.Code.End
                        )
                        references.append(ref)
                        
                except Exception as e:
                    logger.debug(f"Error processing field: {e}")
                    continue
            
            logger.info(f"Extracted {len(references)} references")
            return references
            
        except Exception as e:
            logger.warning(f"Failed to extract references: {e}")
            return []
    
    def extract_structure(self, document: Any) -> DocumentStructure:
        """
        提取完整的文档结构
        
        Args:
            document: Word 文档对象
            
        Returns:
            文档结构对象
        """
        try:
            # 获取基本统计信息
            page_count = document.ComputeStatistics(win32_constants.wdStatisticPages)
            word_count = document.ComputeStatistics(win32_constants.wdStatisticWords)
            
            # 提取各种结构元素
            headings = self.extract_headings(document)
            styles = self.extract_styles(document)
            toc_entries = self.extract_toc_entries(document)
            hyperlinks = self.extract_hyperlinks(document)
            references = self.extract_references(document)
            
            structure = DocumentStructure(
                headings=headings,
                styles=styles,
                toc_entries=toc_entries,
                hyperlinks=hyperlinks,
                references=references,
                page_count=max(1, page_count),
                word_count=max(0, word_count)
            )
            
            logger.info(f"Document structure extracted: {page_count} pages, {word_count} words")
            return structure
            
        except Exception as e:
            raise DocumentError(f"Failed to extract document structure: {e}")
    
    def create_snapshot(self, document: Any, document_path: str) -> DocumentSnapshot:
        """
        创建文档快照
        
        Args:
            document: Word 文档对象
            document_path: 文档路径
            
        Returns:
            文档快照对象
        """
        try:
            # 提取结构和批注
            structure = self.extract_structure(document)
            comments = self.extract_comments(document)
            
            # 计算校验和
            checksum = calculate_file_checksum(document_path)
            
            snapshot = DocumentSnapshot(
                timestamp=datetime.now(),
                document_path=document_path,
                structure=structure,
                comments=comments,
                checksum=checksum
            )
            
            logger.info(f"Document snapshot created for: {document_path}")
            return snapshot
            
        except Exception as e:
            raise DocumentError(f"Failed to create document snapshot: {e}")
    
    # 辅助方法
    
    def _is_heading_style(self, style_name: str) -> bool:
        """检查是否为标题样式"""
        heading_patterns = [
            "标题", "Heading", "Title", "题目",
            "一级标题", "二级标题", "三级标题",
            "Heading 1", "Heading 2", "Heading 3"
        ]
        
        style_lower = style_name.lower()
        return any(pattern.lower() in style_lower for pattern in heading_patterns)
    
    def _get_heading_level(self, style_name: str) -> int:
        """从样式名称推断标题级别"""
        style_lower = style_name.lower()
        
        # 数字级别
        for i in range(1, 10):
            if f"{i}" in style_lower or f"heading {i}" in style_lower:
                return i
        
        # 中文级别
        level_map = {
            "一级": 1, "二级": 2, "三级": 3, "四级": 4, "五级": 5,
            "1级": 1, "2级": 2, "3级": 3, "4级": 4, "5级": 5
        }
        
        for pattern, level in level_map.items():
            if pattern in style_lower:
                return level
        
        # 默认为1级
        return 1
    
    def _get_style_type(self, style_type: int) -> str:
        """获取样式类型字符串"""
        type_map = {
            win32_constants.wdStyleTypeParagraph: "paragraph",
            win32_constants.wdStyleTypeCharacter: "character",
            win32_constants.wdStyleTypeTable: "table",
            win32_constants.wdStyleTypeList: "list"
        }
        return type_map.get(style_type, "unknown")
    
    def _extract_page_number_from_toc_text(self, text: str) -> int:
        """从目录文本中提取页码"""
        import re
        
        # 查找文本末尾的数字
        match = re.search(r'(\d+)\s*$', text)
        if match:
            return int(match.group(1))
        
        return 1
    
    def _estimate_toc_level(self, range_obj: Any) -> int:
        """估算目录条目的级别"""
        try:
            # 尝试从段落样式推断级别
            para = range_obj.Paragraphs(1)
            style_name = para.Style.NameLocal.lower()
            
            # 查找级别指示符
            for i in range(1, 6):
                if f"toc {i}" in style_name or f"目录 {i}" in style_name:
                    return i
            
            # 默认级别
            return 1
            
        except Exception:
            return 1