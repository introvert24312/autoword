"""
Test AutoWord Exporter
测试导出器
"""

import json
import pytest
from unittest.mock import Mock, patch, mock_open
from pathlib import Path
from datetime import datetime

from autoword.core.exporter import Exporter, export_execution_report
from autoword.core.models import (
    TaskPlan, Task, TaskType, RiskLevel, LocatorType, Locator,
    ExecutionResult, TaskResult, Comment, DocumentStructure, Heading, Style
)


class TestExporter:
    """测试导出器"""
    
    def setup_method(self):
        """测试前设置"""
        with patch('pathlib.Path.mkdir'):
            self.exporter = Exporter("test_output")
    
    @patch('builtins.open', new_callable=mock_open)
    @patch('json.dump')
    def test_export_plan(self, mock_json_dump, mock_file):
        """测试导出任务计划"""
        # 创建测试任务计划
        tasks = [
            Task(
                id="task_1",
                type=TaskType.REWRITE,
                source_comment_id="comment_1",
                locator=Locator(by=LocatorType.FIND, value="test"),
                instruction="重写内容",
                risk=RiskLevel.LOW
            )
        ]
        
        task_plan = TaskPlan(
            tasks=tasks,
            document_path="test.docx",
            total_tasks=1
        )
        
        # 执行导出
        result = self.exporter.export_plan(task_plan, "test_plan.json")
        
        # 验证结果
        assert "test_plan.json" in result
        mock_file.assert_called_once()
        mock_json_dump.assert_called_once()
        
        # 验证导出的数据结构
        call_args = mock_json_dump.call_args[0]
        exported_data = call_args[0]
        
        assert "metadata" in exported_data
        assert "tasks" in exported_data
        assert exported_data["metadata"]["total_tasks"] == 1
        assert len(exported_data["tasks"]) == 1
        assert exported_data["tasks"][0]["id"] == "task_1"
    
    @patch('builtins.open', new_callable=mock_open)
    @patch('json.dump')
    def test_export_execution_log(self, mock_json_dump, mock_file):
        """测试导出执行日志"""
        # 创建测试执行结果
        task_results = [
            TaskResult(
                task_id="task_1",
                success=True,
                message="执行成功",
                execution_time=1.5
            )
        ]
        
        execution_result = ExecutionResult(
            success=True,
            total_tasks=1,
            completed_tasks=1,
            failed_tasks=0,
            task_results=task_results,
            execution_time=2.0
        )
        
        # 执行导出
        result = self.exporter.export_execution_log(execution_result, "test_log.json")
        
        # 验证结果
        assert "test_log.json" in result
        mock_file.assert_called_once()
        mock_json_dump.assert_called_once()
        
        # 验证导出的数据结构
        call_args = mock_json_dump.call_args[0]
        exported_data = call_args[0]
        
        assert "metadata" in exported_data
        assert "summary" in exported_data
        assert "task_results" in exported_data
        assert exported_data["summary"]["success"] is True
        assert len(exported_data["task_results"]) == 1
    
    @patch('builtins.open', new_callable=mock_open)
    def test_export_diff_report(self, mock_file):
        """测试导出变更报告"""
        # 创建测试文档结构
        before_structure = DocumentStructure(
            headings=[
                Heading(level=1, text="标题1", style="标题 1", range_start=0, range_end=10)
            ],
            styles=[
                Style(name="正文", type="paragraph", built_in=True, in_use=True)
            ],
            toc_entries=[],
            hyperlinks=[],
            references=[],
            page_count=1,
            word_count=100
        )
        
        after_structure = DocumentStructure(
            headings=[
                Heading(level=1, text="标题1", style="标题 1", range_start=0, range_end=10),
                Heading(level=2, text="标题2", style="标题 2", range_start=20, range_end=30)
            ],
            styles=[
                Style(name="正文", type="paragraph", built_in=True, in_use=True),
                Style(name="标题 2", type="paragraph", built_in=True, in_use=True)
            ],
            toc_entries=[],
            hyperlinks=[],
            references=[],
            page_count=2,
            word_count=200
        )
        
        # 执行导出
        result = self.exporter.export_diff_report(before_structure, after_structure, "test_diff.md")
        
        # 验证结果
        assert "test_diff.md" in result
        mock_file.assert_called_once()
        
        # 验证写入的内容
        written_content = "".join(call.args[0] for call in mock_file().write.call_args_list)
        assert "# 文档变更报告" in written_content
        assert "页数变化: 1 → 2" in written_content
        assert "字数变化: 100 → 200" in written_content
    
    @patch('builtins.open', new_callable=mock_open)
    @patch('json.dump')
    def test_export_comments_data(self, mock_json_dump, mock_file):
        """测试导出批注数据"""
        # 创建测试批注
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
        
        # 执行导出
        result = self.exporter.export_comments_data(comments, "test_comments.json")
        
        # 验证结果
        assert "test_comments.json" in result
        mock_file.assert_called_once()
        mock_json_dump.assert_called_once()
        
        # 验证导出的数据结构
        call_args = mock_json_dump.call_args[0]
        exported_data = call_args[0]
        
        assert "metadata" in exported_data
        assert "comments" in exported_data
        assert exported_data["metadata"]["total_comments"] == 1
        assert len(exported_data["comments"]) == 1
        assert exported_data["comments"][0]["id"] == "comment_1"
    
    def test_compare_headings_added(self):
        """测试标题比较 - 新增"""
        before_headings = []
        after_headings = [
            Heading(level=1, text="新标题", style="标题 1", range_start=0, range_end=10)
        ]
        
        result = self.exporter._compare_headings(before_headings, after_headings)
        
        assert any("新增标题" in line for line in result)
        assert any("新标题" in line for line in result)
    
    def test_compare_headings_removed(self):
        """测试标题比较 - 删除"""
        before_headings = [
            Heading(level=1, text="旧标题", style="标题 1", range_start=0, range_end=10)
        ]
        after_headings = []
        
        result = self.exporter._compare_headings(before_headings, after_headings)
        
        assert any("删除标题" in line for line in result)
        assert any("旧标题" in line for line in result)
    
    def test_compare_headings_modified(self):
        """测试标题比较 - 修改"""
        before_headings = [
            Heading(level=1, text="标题", style="标题 1", range_start=0, range_end=10)
        ]
        after_headings = [
            Heading(level=2, text="标题", style="标题 2", range_start=0, range_end=10)
        ]
        
        result = self.exporter._compare_headings(before_headings, after_headings)
        
        assert any("修改标题" in line for line in result)
        assert any("级别: 1 → 2" in line for line in result)
    
    def test_compare_styles(self):
        """测试样式比较"""
        before_styles = [
            Style(name="正文", type="paragraph", built_in=True, in_use=True)
        ]
        after_styles = [
            Style(name="正文", type="paragraph", built_in=True, in_use=True),
            Style(name="标题 1", type="paragraph", built_in=True, in_use=True)
        ]
        
        result = self.exporter._compare_styles(before_styles, after_styles)
        
        assert any("新增样式" in line for line in result)
        assert any("标题 1" in line for line in result)


class TestConvenienceFunction:
    """测试便捷函数"""
    
    @patch('autoword.core.exporter.Exporter')
    def test_export_execution_report(self, mock_exporter_class):
        """测试导出执行报告便捷函数"""
        # 设置模拟
        mock_exporter = Mock()
        mock_exporter.export_plan.return_value = "plan.json"
        mock_exporter.export_execution_log.return_value = "log.json"
        mock_exporter.export_diff_report.return_value = "diff.md"
        mock_exporter_class.return_value = mock_exporter
        
        # 创建测试数据
        task_plan = Mock()
        execution_result = Mock()
        before_structure = Mock()
        after_structure = Mock()
        
        # 执行测试
        result = export_execution_report(
            task_plan, execution_result, before_structure, after_structure
        )
        
        # 验证结果
        assert "plan" in result
        assert "log" in result
        assert "diff" in result
        assert result["plan"] == "plan.json"
        assert result["log"] == "log.json"
        assert result["diff"] == "diff.md"
        
        # 验证调用
        mock_exporter.export_plan.assert_called_once_with(task_plan)
        mock_exporter.export_execution_log.assert_called_once_with(execution_result)
        mock_exporter.export_diff_report.assert_called_once_with(before_structure, after_structure)


if __name__ == "__main__":
    pytest.main([__file__])