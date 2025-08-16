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


class DocumentInspector:
    """Word 文档检查器"""
    
    def __init__(self):
        """初始化文档检查器"""
        pass
    
    def get_document_info(self, document):
        """获取文档信息"""
        try:
            # 提取批注
            comments = self.extract_comments(document)
            
            # 提取文档结构
            structure = self.extract_structure(document)
            
            # 创建文档信息对象
            from .models import Document
            return Document(
                path=document.FullName,
                title=document.BuiltInDocumentProperties("Title").Value or "Untitled",
                page_count=document.Range().Information(4),  # wdNumberOfPagesInDocument
                word_count=document.Words.Count,
                comments=comments
            )
            
        except Exception as e:
            raise DocumentError(f"获取文档信息失败: {str(e)}")
    
    def extract_comments(self, document) -> List[Comment]:
        """提取文档批注"""
        comments = []
        
        try:
            for i, comment in enumerate(document.Comments, 1):
                try:
                    # 获取批注信息
                    comment_obj = Comment(
                        id=f"comment_{i}",
                        author=comment.Author,
                        comment_text=comment.Range.Text.strip(),
                        anchor_text=comment.Scope.Text.strip()[:50],  # 限制长度
                        page=comment.Scope.Information(1),  # wdActiveEndPageNumber
                        range_start=comment.Scope.Start,
                        range_end=comment.Scope.End
                    )
                    comments.append(comment_obj)
                    
                except Exception as e:
                    logger.warning(f"提取批注 {i} 失败: {str(e)}")
                    continue
                    
        except Exception as e:
            logger.error(f"提取批注失败: {str(e)}")
            
        return comments
    
    def extract_structure(self, document):
        """提取文档结构"""
        try:
            # 简化的结构提取
            return {
                "headings": [],
                "styles": [],
                "toc_entries": [],
                "hyperlinks": []
            }
        except Exception as e:
            logger.warning(f"提取文档结构失败: {str(e)}")
            return {
                "headings": [],
                "styles": [],
                "toc_entries": [],
                "hyperlinks": []
            }


# 保持向后兼容
DocInspector = DocumentInspector