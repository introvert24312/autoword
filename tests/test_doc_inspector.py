"""
Test AutoWord Document Inspector
测试文档检查器
"""

import pytest
from unittest.mock import Mock, patch, MagicMock
from datetime import datetime

from autoword.core.doc_inspector import DocInspector
from autoword.core.models import Comment, Heading, DocumentStructure
from autoword.core.exceptions import DocumentError


class TestDocInspector:
    """测试 DocInspector 类"""
    
    def setup_method(self):
        """测试前设置"""
        self.inspector = DocInspector()
    
    @patch('autoword.core.doc_inspector.win32_constants')
    def test_extract_comments_success(self, mock_constants):
        """测试成功提取批注"""
        mock_constants.wdActiveEndPageNumber = 1
        
        # 创建模拟批注
        mock_comment1 = Mock()
        mock_comment1.Author = "张三"
        mock_comment1.Range.Text = "这是第一个批注"
        mock_comment1.Scope.Text = "锚点文本1"
        mock_comment1.Scope.Start = 100
        mock_comment1.Scope.End = 150
        mock_comment1.Scope.Information.return_value = 1
        mock_comment1.Date = datetime(2024, 1, 1, 12, 0, 0)
        
        mock_comment2 = Mock()
        mock_comment2.Author = "李四"
        mock_comment2.Range.Text = "这是第二个批注"
        mock_comment2.Scope.Text = "锚点文本2"
        mock_comment2.Scope.Start = 200
        mock_comment2.Scope.End = 250
        mock_comment2.Scope.Information.return_value = 2
        mock_comment2.Date = datetime(2024, 1, 1, 13, 0, 0)
        
        # 创建模拟文档
        mock_document = Mock()
        mock_document.Comments = [mock_comment1, mock_comment2]
        
        # 执行测试
        comments = self.inspector.extract_comments(mock_document)
        
        # 验证结果
        assert len(comments) == 2
        
        assert comments[0].id == "comment_1"
        assert comments[0].author == "张三"
        assert comments[0].comment_text == "这是第一个批注"
        assert comments[0].anchor_text == "锚点文本1"
        assert comments[0].page == 1
        
        assert comments[1].id == "comment_2"
        assert comments[1].author == "李四"
        assert comments[1].comment_text == "这是第二个批注"
        assert comments[1].anchor_text == "锚点文本2"
        assert comments[1].page == 2
    
    def test_extract_comments_empty_document(self):
        """测试提取空文档的批注"""
        mock_document = Mock()
        mock_document.Comments = []
        
        comments = self.inspector.extract_comments(mock_document)
        
        assert len(comments) == 0
    
    @patch('autoword.core.doc_inspector.win32_constants')
    def test_extract_comments_with_errors(self, mock_constants):
        """测试提取批注时的错误处理"""
        mock_constants.wdActiveEndPageNumber = 1
        
        # 创建一个会抛出异常的批注
        mock_comment_bad = Mock()
        mock_comment_bad.Author = Mock(side_effect=Exception("Access denied"))
        
        # 创建一个正常的批注
        mock_comment_good = Mock()
        mock_comment_good.Author = "张三"
        mock_comment_good.Range.Text = "正常批注"
        mock_comment_good.Scope.Text = "锚点文本"
        mock_comment_good.Scope.Start = 100
        mock_comment_good.Scope.End = 150
        mock_comment_good.Scope.Information.return_value = 1
        mock_comment_good.Date = None  # 设置为 None 而不是 Mock
        
        mock_document = Mock()
        mock_document.Comments = [mock_comment_bad, mock_comment_good]
        
        # 执行测试
        comments = self.inspector.extract_comments(mock_document)
        
        # 验证只提取了正常的批注
        assert len(comments) == 1
        assert comments[0].author == "张三"
    
    def test_extract_headings_success(self):
        """测试成功提取标题"""
        # 创建模拟段落
        mock_para1 = Mock()
        mock_para1.Style.NameLocal = "标题 1"
        mock_para1.Range.Text = "第一章 概述\r"
        mock_para1.Range.Start = 0
        mock_para1.Range.End = 10
        
        mock_para2 = Mock()
        mock_para2.Style.NameLocal = "Heading 2"
        mock_para2.Range.Text = "1.1 背景\r"
        mock_para2.Range.Start = 20
        mock_para2.Range.End = 30
        
        mock_para3 = Mock()
        mock_para3.Style.NameLocal = "正文"
        mock_para3.Range.Text = "这是正文内容\r"
        mock_para3.Range.Start = 40
        mock_para3.Range.End = 50
        
        mock_document = Mock()
        mock_document.Paragraphs = [mock_para1, mock_para2, mock_para3]
        
        # 执行测试
        headings = self.inspector.extract_headings(mock_document)
        
        # 验证结果
        assert len(headings) == 2
        
        assert headings[0].level == 1
        assert headings[0].text == "第一章 概述"
        assert headings[0].style == "标题 1"
        
        assert headings[1].level == 2
        assert headings[1].text == "1.1 背景"
        assert headings[1].style == "Heading 2"
    
    def test_extract_styles_success(self):
        """测试成功提取样式"""
        # 创建模拟样式
        mock_style1 = Mock()
        mock_style1.NameLocal = "标题 1"
        mock_style1.Type = 1  # wdStyleTypeParagraph
        mock_style1.BuiltIn = True
        mock_style1.InUse = True
        
        mock_style2 = Mock()
        mock_style2.NameLocal = "正文"
        mock_style2.Type = 1
        mock_style2.BuiltIn = True
        mock_style2.InUse = True
        
        mock_document = Mock()
        mock_document.Styles = [mock_style1, mock_style2]
        
        # 模拟常量
        with patch('autoword.core.doc_inspector.win32_constants') as mock_constants:
            mock_constants.wdStyleTypeParagraph = 1
            
            # 执行测试
            styles = self.inspector.extract_styles(mock_document)
        
        # 验证结果
        assert len(styles) == 2
        assert styles[0].name == "标题 1"
        assert styles[0].type == "paragraph"
        assert styles[0].built_in is True
        assert styles[0].in_use is True
    
    def test_extract_hyperlinks_success(self):
        """测试成功提取超链接"""
        # 创建模拟超链接
        mock_link1 = Mock()
        mock_link1.TextToDisplay = "百度搜索"
        mock_link1.Address = "https://www.baidu.com"
        mock_link1.SubAddress = None
        mock_link1.Range.Start = 100
        mock_link1.Range.End = 110
        
        mock_link2 = Mock()
        mock_link2.TextToDisplay = "内部链接"
        mock_link2.Address = None
        mock_link2.SubAddress = "bookmark1"
        mock_link2.Range.Start = 200
        mock_link2.Range.End = 210
        
        mock_document = Mock()
        mock_document.Hyperlinks = [mock_link1, mock_link2]
        
        # 执行测试
        hyperlinks = self.inspector.extract_hyperlinks(mock_document)
        
        # 验证结果
        assert len(hyperlinks) == 2
        
        assert hyperlinks[0].text == "百度搜索"
        assert hyperlinks[0].address == "https://www.baidu.com"
        assert hyperlinks[0].type == "web"
        
        assert hyperlinks[1].text == "内部链接"
        assert hyperlinks[1].address == "bookmark1"
        assert hyperlinks[1].type == "internal"
    
    @patch('autoword.core.doc_inspector.win32_constants')
    def test_extract_structure_success(self, mock_constants):
        """测试成功提取文档结构"""
        mock_constants.wdStatisticPages = 2
        mock_constants.wdStatisticWords = 3
        
        # 创建模拟文档
        mock_document = Mock()
        mock_document.ComputeStatistics.side_effect = [5, 1000]  # 5页，1000字
        mock_document.Paragraphs = []
        mock_document.Styles = []
        mock_document.TablesOfContents = []
        mock_document.Hyperlinks = []
        mock_document.Bookmarks = []
        mock_document.Fields = []
        
        # 执行测试
        structure = self.inspector.extract_structure(mock_document)
        
        # 验证结果
        assert isinstance(structure, DocumentStructure)
        assert structure.page_count == 5
        assert structure.word_count == 1000
        assert len(structure.headings) == 0
        assert len(structure.styles) == 0
    
    @patch('autoword.core.doc_inspector.calculate_file_checksum')
    def test_create_snapshot_success(self, mock_checksum):
        """测试成功创建文档快照"""
        mock_checksum.return_value = "abc123"
        
        # 创建模拟文档
        mock_document = Mock()
        mock_document.ComputeStatistics.side_effect = [1, 100]
        mock_document.Paragraphs = []
        mock_document.Styles = []
        mock_document.TablesOfContents = []
        mock_document.Hyperlinks = []
        mock_document.Bookmarks = []
        mock_document.Fields = []
        mock_document.Comments = []
        
        with patch('autoword.core.doc_inspector.win32_constants'):
            # 执行测试
            snapshot = self.inspector.create_snapshot(mock_document, "test.docx")
        
        # 验证结果
        assert snapshot.document_path == "test.docx"
        assert snapshot.checksum == "abc123"
        assert isinstance(snapshot.timestamp, datetime)
        assert isinstance(snapshot.structure, DocumentStructure)
        assert len(snapshot.comments) == 0
    
    def test_is_heading_style(self):
        """测试标题样式识别"""
        assert self.inspector._is_heading_style("标题 1") is True
        assert self.inspector._is_heading_style("Heading 2") is True
        assert self.inspector._is_heading_style("一级标题") is True
        assert self.inspector._is_heading_style("正文") is False
        assert self.inspector._is_heading_style("Normal") is False
    
    def test_get_heading_level(self):
        """测试标题级别推断"""
        assert self.inspector._get_heading_level("标题 1") == 1
        assert self.inspector._get_heading_level("Heading 2") == 2
        assert self.inspector._get_heading_level("三级标题") == 3
        assert self.inspector._get_heading_level("一级标题") == 1
        assert self.inspector._get_heading_level("未知标题") == 1  # 默认级别
    
    @patch('autoword.core.doc_inspector.win32_constants')
    def test_get_style_type(self, mock_constants):
        """测试样式类型获取"""
        mock_constants.wdStyleTypeParagraph = 1
        mock_constants.wdStyleTypeCharacter = 2
        mock_constants.wdStyleTypeTable = 3
        mock_constants.wdStyleTypeList = 4
        
        assert self.inspector._get_style_type(1) == "paragraph"
        assert self.inspector._get_style_type(2) == "character"
        assert self.inspector._get_style_type(3) == "table"
        assert self.inspector._get_style_type(4) == "list"
        assert self.inspector._get_style_type(999) == "unknown"
    
    def test_extract_page_number_from_toc_text(self):
        """测试从目录文本提取页码"""
        assert self.inspector._extract_page_number_from_toc_text("第一章 概述 .......... 1") == 1
        assert self.inspector._extract_page_number_from_toc_text("1.1 背景介绍 5") == 5
        assert self.inspector._extract_page_number_from_toc_text("结论 123") == 123
        assert self.inspector._extract_page_number_from_toc_text("无页码标题") == 1  # 默认页码


if __name__ == "__main__":
    pytest.main([__file__])