"""
Test AutoWord Prompt Builder
测试提示词构建器
"""

import json
import pytest
from unittest.mock import Mock, patch, mock_open
from datetime import datetime

from autoword.core.prompt_builder import PromptBuilder, PromptContext
from autoword.core.models import (
    Comment, DocumentStructure, Heading, Style, TocEntry, Hyperlink, Reference
)
from autoword.core.constants import SYSTEM_PROMPT
from autoword.core.exceptions import ValidationError


class TestPromptContext:
    """测试 PromptContext 数据类"""
    
    def test_prompt_context_creation(self):
        """测试创建提示词上下文"""
        structure = DocumentStructure(
            headings=[],
            styles=[],
            toc_entries=[],
            hyperlinks=[],
            references=[],
            page_count=1,
            word_count=100
        )
        
        comments = [
            Comment(
                id="comment_1",
                author="张三",
                page=1,
                anchor_text="测试文本",
                comment_text="这是测试批注",
                range_start=0,
                range_end=10
            )
        ]
        
        context = PromptContext(
            document_structure=structure,
            comments=comments,
            document_path="test.docx"
        )
        
        assert context.document_structure == structure
        assert context.comments == comments
        assert context.document_path == "test.docx"
        assert context.max_context_length == 32000
        assert context.include_full_content is False


class TestPromptBuilder:
    """测试 PromptBuilder 类"""
    
    def setup_method(self):
        """测试前设置"""
        self.builder = PromptBuilder()
    
    def test_build_system_prompt(self):
        """测试构建系统提示词"""
        system_prompt = self.builder.build_system_prompt()
        assert system_prompt == SYSTEM_PROMPT
        assert "Word editing planner" in system_prompt
        assert "JSON" in system_prompt
    
    def test_build_document_summary_basic(self):
        """测试构建基本文档摘要"""
        structure = DocumentStructure(
            headings=[],
            styles=[],
            toc_entries=[],
            hyperlinks=[],
            references=[],
            page_count=5,
            word_count=1000
        )
        
        summary = self.builder.build_document_summary(structure)
        
        assert "5 页" in summary
        assert "1000 字" in summary
    
    def test_build_document_summary_with_headings(self):
        """测试包含标题的文档摘要"""
        headings = [
            Heading(level=1, text="第一章", style="标题 1", range_start=0, range_end=10),
            Heading(level=2, text="1.1 概述", style="标题 2", range_start=20, range_end=30),
            Heading(level=2, text="1.2 方法", style="标题 2", range_start=40, range_end=50),
            Heading(level=1, text="第二章", style="标题 1", range_start=60, range_end=70)
        ]
        
        structure = DocumentStructure(
            headings=headings,
            styles=[],
            toc_entries=[],
            hyperlinks=[],
            references=[],
            page_count=1,
            word_count=100
        )
        
        summary = self.builder.build_document_summary(structure)
        
        assert "标题结构 (4 个)" in summary
        assert "1级标题 (2 个)" in summary
        assert "2级标题 (2 个)" in summary
        assert "第一章" in summary
        assert "第二章" in summary
    
    def test_build_document_summary_with_styles(self):
        """测试包含样式的文档摘要"""
        styles = [
            Style(name="标题 1", type="paragraph", built_in=True, in_use=True),
            Style(name="正文", type="paragraph", built_in=True, in_use=True),
            Style(name="未使用样式", type="paragraph", built_in=False, in_use=False)
        ]
        
        structure = DocumentStructure(
            headings=[],
            styles=styles,
            toc_entries=[],
            hyperlinks=[],
            references=[],
            page_count=1,
            word_count=100
        )
        
        summary = self.builder.build_document_summary(structure)
        
        assert "使用中的样式 (2 个)" in summary
        assert "paragraph: 标题 1, 正文" in summary
        assert "未使用样式" not in summary
    
    def test_build_document_summary_with_toc(self):
        """测试包含目录的文档摘要"""
        toc_entries = [
            TocEntry(level=1, text="第一章 概述", page_number=1, range_start=0, range_end=10),
            TocEntry(level=2, text="1.1 背景", page_number=2, range_start=20, range_end=30)
        ]
        
        structure = DocumentStructure(
            headings=[],
            styles=[],
            toc_entries=toc_entries,
            hyperlinks=[],
            references=[],
            page_count=1,
            word_count=100
        )
        
        summary = self.builder.build_document_summary(structure)
        
        assert "目录结构 (2 个条目)" in summary
        assert "1级: 第一章 概述 (第1页)" in summary
        assert "2级: 1.1 背景 (第2页)" in summary
    
    def test_build_document_summary_with_hyperlinks(self):
        """测试包含超链接的文档摘要"""
        hyperlinks = [
            Hyperlink(text="百度", address="https://baidu.com", type="web", range_start=0, range_end=10),
            Hyperlink(text="内部链接", address="bookmark1", type="internal", range_start=20, range_end=30),
            Hyperlink(text="邮箱", address="mailto:test@example.com", type="email", range_start=40, range_end=50)
        ]
        
        structure = DocumentStructure(
            headings=[],
            styles=[],
            toc_entries=[],
            hyperlinks=hyperlinks,
            references=[],
            page_count=1,
            word_count=100
        )
        
        summary = self.builder.build_document_summary(structure)
        
        assert "超链接 (3 个)" in summary
        assert "web: 1 个" in summary
        assert "internal: 1 个" in summary
        assert "email: 1 个" in summary
    
    def test_build_comments_summary_empty(self):
        """测试空批注摘要"""
        summary = self.builder.build_comments_summary([])
        assert summary == "批注: 无"
    
    def test_build_comments_summary_single_author(self):
        """测试单作者批注摘要"""
        comments = [
            Comment(
                id="comment_1",
                author="张三",
                page=1,
                anchor_text="测试锚点",
                comment_text="这是第一个批注",
                range_start=0,
                range_end=10
            ),
            Comment(
                id="comment_2",
                author="张三",
                page=2,
                anchor_text="另一个锚点",
                comment_text="这是第二个批注",
                range_start=20,
                range_end=30
            )
        ]
        
        summary = self.builder.build_comments_summary(comments)
        
        assert "批注 (2 个)" in summary
        assert "1. ID: comment_1" in summary
        assert "作者: 张三" in summary
        assert "页码: 1" in summary
        assert "锚点: \"测试锚点\"" in summary
        assert "内容: \"这是第一个批注\"" in summary
        assert "位置: 0-10" in summary
    
    def test_build_comments_summary_multiple_authors(self):
        """测试多作者批注摘要"""
        comments = [
            Comment(
                id="comment_1",
                author="张三",
                page=1,
                anchor_text="锚点1",
                comment_text="张三的批注",
                range_start=0,
                range_end=10
            ),
            Comment(
                id="comment_2",
                author="李四",
                page=1,
                anchor_text="锚点2",
                comment_text="李四的批注",
                range_start=20,
                range_end=30
            ),
            Comment(
                id="comment_3",
                author="张三",
                page=2,
                anchor_text="锚点3",
                comment_text="张三的第二个批注",
                range_start=40,
                range_end=50
            )
        ]
        
        summary = self.builder.build_comments_summary(comments)
        
        assert "批注 (3 个)" in summary
        assert "作者统计:" in summary
        assert "张三: 2 个" in summary
        assert "李四: 1 个" in summary
    
    @patch('autoword.core.prompt_builder.load_json_schema')
    def test_get_schema_success(self, mock_load_schema):
        """测试成功获取 Schema"""
        mock_schema = {"type": "object", "properties": {}}
        mock_load_schema.return_value = mock_schema
        
        schema = self.builder.get_schema()
        
        assert schema == mock_schema
        mock_load_schema.assert_called_once_with(self.builder.schema_path)
    
    @patch('autoword.core.prompt_builder.load_json_schema')
    def test_get_schema_fallback_to_default(self, mock_load_schema):
        """测试 Schema 加载失败时使用默认值"""
        mock_load_schema.side_effect = Exception("File not found")
        
        schema = self.builder.get_schema()
        
        assert "type" in schema
        assert "properties" in schema
        assert "tasks" in schema["properties"]
    
    def test_build_user_prompt(self):
        """测试构建用户提示词"""
        structure = DocumentStructure(
            headings=[],
            styles=[],
            toc_entries=[],
            hyperlinks=[],
            references=[],
            page_count=1,
            word_count=100
        )
        
        comments = [
            Comment(
                id="comment_1",
                author="张三",
                page=1,
                anchor_text="测试",
                comment_text="测试批注",
                range_start=0,
                range_end=10
            )
        ]
        
        context = PromptContext(
            document_structure=structure,
            comments=comments
        )
        
        user_prompt = self.builder.build_user_prompt(context)
        
        assert "Document Structure Summary:" in user_prompt
        assert "Comments" in user_prompt
        assert "Output Schema" in user_prompt
        assert "1 页" in user_prompt
        assert "100 字" in user_prompt
        assert "张三" in user_prompt
        assert "测试批注" in user_prompt
    
    def test_estimate_token_count(self):
        """测试 token 数量估算"""
        # 纯英文
        english_text = "Hello world this is a test"
        english_tokens = self.builder.estimate_token_count(english_text)
        assert english_tokens > 0
        
        # 纯中文
        chinese_text = "这是一个测试文本"
        chinese_tokens = self.builder.estimate_token_count(chinese_text)
        assert chinese_tokens > 0
        
        # 混合文本
        mixed_text = "这是 mixed text 测试"
        mixed_tokens = self.builder.estimate_token_count(mixed_text)
        assert mixed_tokens > 0
        
        # 中文应该比英文估算更多 tokens
        assert chinese_tokens > len(chinese_text.split())
    
    def test_check_context_length_within_limit(self):
        """测试上下文长度检查（在限制内）"""
        structure = DocumentStructure(
            headings=[],
            styles=[],
            toc_entries=[],
            hyperlinks=[],
            references=[],
            page_count=1,
            word_count=10
        )
        
        context = PromptContext(
            document_structure=structure,
            comments=[],
            max_context_length=10000
        )
        
        result = self.builder.check_context_length(context)
        
        assert result["is_within_limit"] is True
        assert result["overflow_tokens"] == 0
        assert result["total_tokens"] > 0
        assert result["total_tokens"] <= 10000
    
    def test_check_context_length_overflow(self):
        """测试上下文长度检查（超出限制）"""
        # 创建大量内容
        long_comments = []
        for i in range(100):
            comment = Comment(
                id=f"comment_{i}",
                author="测试作者",
                page=1,
                anchor_text="很长的锚点文本" * 10,
                comment_text="很长的批注内容" * 20,
                range_start=i * 100,
                range_end=(i + 1) * 100
            )
            long_comments.append(comment)
        
        structure = DocumentStructure(
            headings=[],
            styles=[],
            toc_entries=[],
            hyperlinks=[],
            references=[],
            page_count=1,
            word_count=10000
        )
        
        context = PromptContext(
            document_structure=structure,
            comments=long_comments,
            max_context_length=1000  # 很小的限制
        )
        
        result = self.builder.check_context_length(context)
        
        assert result["is_within_limit"] is False
        assert result["overflow_tokens"] > 0
        assert result["total_tokens"] > 1000
    
    def test_handle_context_overflow_by_headings(self):
        """测试按标题处理上下文溢出"""
        headings = [
            Heading(level=1, text="第一章", style="标题 1", range_start=0, range_end=100),
            Heading(level=1, text="第二章", style="标题 1", range_start=1000, range_end=1100),
            Heading(level=1, text="第三章", style="标题 1", range_start=2000, range_end=2100)
        ]
        
        comments = [
            Comment(id="c1", author="张三", page=1, anchor_text="锚点1", 
                   comment_text="第一章的批注", range_start=50, range_end=60),
            Comment(id="c2", author="李四", page=2, anchor_text="锚点2", 
                   comment_text="第二章的批注", range_start=1050, range_end=1060),
            Comment(id="c3", author="王五", page=3, anchor_text="锚点3", 
                   comment_text="第三章的批注", range_start=2050, range_end=2060)
        ]
        
        structure = DocumentStructure(
            headings=headings,
            styles=[],
            toc_entries=[],
            hyperlinks=[],
            references=[],
            page_count=3,
            word_count=3000
        )
        
        context = PromptContext(
            document_structure=structure,
            comments=comments
        )
        
        chunks = self.builder.handle_context_overflow(context)
        
        assert len(chunks) == 3  # 按3个一级标题分块
        
        # 验证第一个块
        assert len(chunks[0].document_structure.headings) == 1
        assert chunks[0].document_structure.headings[0].text == "第一章"
        assert len(chunks[0].comments) == 1
        assert chunks[0].comments[0].id == "c1"
        
        # 验证第二个块
        assert len(chunks[1].document_structure.headings) == 1
        assert chunks[1].document_structure.headings[0].text == "第二章"
        assert len(chunks[1].comments) == 1
        assert chunks[1].comments[0].id == "c2"
    
    def test_handle_context_overflow_by_comments(self):
        """测试按批注处理上下文溢出"""
        # 没有标题的情况，应该按批注分块
        comments = []
        for i in range(10):
            comment = Comment(
                id=f"comment_{i}",
                author="测试作者",
                page=1,
                anchor_text=f"锚点{i}",
                comment_text=f"批注内容{i}",
                range_start=i * 100,
                range_end=(i + 1) * 100
            )
            comments.append(comment)
        
        structure = DocumentStructure(
            headings=[],  # 没有标题
            styles=[],
            toc_entries=[],
            hyperlinks=[],
            references=[],
            page_count=1,
            word_count=1000
        )
        
        context = PromptContext(
            document_structure=structure,
            comments=comments
        )
        
        chunks = self.builder.handle_context_overflow(context)
        
        # 应该分成大约3块
        assert len(chunks) >= 2
        
        # 验证所有批注都被包含
        total_comments = sum(len(chunk.comments) for chunk in chunks)
        assert total_comments == 10
    
    def test_get_default_schema(self):
        """测试获取默认 Schema"""
        schema = self.builder._get_default_schema()
        
        assert schema["type"] == "object"
        assert "tasks" in schema["properties"]
        assert "required" in schema
        assert "tasks" in schema["required"]
        
        # 验证任务 Schema 结构
        task_schema = schema["properties"]["tasks"]["items"]
        assert "id" in task_schema["required"]
        assert "type" in task_schema["required"]
        assert "locator" in task_schema["required"]
        assert "instruction" in task_schema["required"]


if __name__ == "__main__":
    pytest.main([__file__])