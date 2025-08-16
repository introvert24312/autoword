"""
Test AutoWord Format Validator
测试格式验证器
"""

import pytest
from unittest.mock import Mock, patch
from datetime import datetime

from autoword.core.format_validator import (
    FormatValidator, DocumentComparator, AuthorizationChecker,
    FormatChange, ValidationReport
)
from autoword.core.models import (
    Task, TaskType, LocatorType, Locator, Comment,
    DocumentSnapshot, DocumentStructure, Heading, Style
)


class TestFormatChange:
    """测试格式变更记录"""
    
    def test_format_change_creation(self):
        """测试创建格式变更记录"""
        change = FormatChange(
            change_type="heading_level_change",
            element_type="heading",
            element_id="heading_1",
            old_value=1,
            new_value=2,
            authorized=False,
            source_comment_id="comment_1"
        )
        
        assert change.change_type == "heading_level_change"
        assert change.element_type == "heading"
        assert change.old_value == 1
        assert change.new_value == 2
        assert change.authorized is False
        assert change.source_comment_id == "comment_1"
        assert isinstance(change.timestamp, datetime)


class TestDocumentComparator:
    """测试文档比较器"""
    
    def setup_method(self):
        """测试前设置"""
        self.comparator = DocumentComparator()
    
    def test_compare_headings_level_change(self):
        """测试标题级别变更比较"""
        # 创建前后标题
        before_headings = [
            Heading(level=1, text="标题1", style="标题 1", range_start=0, range_end=10)
        ]
        
        after_headings = [
            Heading(level=2, text="标题1", style="标题 2", range_start=0, range_end=10)
        ]
        
        changes = self.comparator._compare_headings(before_headings, after_headings)
        
        assert len(changes) == 2  # 级别变更 + 样式变更
        
        level_change = next(c for c in changes if c.change_type == "heading_level_change")
        assert level_change.old_value == 1
        assert level_change.new_value == 2
        
        style_change = next(c for c in changes if c.change_type == "heading_style_change")
        assert style_change.old_value == "标题 1"
        assert style_change.new_value == "标题 2"
    
    def test_compare_headings_added(self):
        """测试新增标题比较"""
        before_headings = []
        
        after_headings = [
            Heading(level=1, text="新标题", style="标题 1", range_start=0, range_end=10)
        ]
        
        changes = self.comparator._compare_headings(before_headings, after_headings)
        
        assert len(changes) == 1
        assert changes[0].change_type == "heading_added"
        assert changes[0].old_value is None
        assert changes[0].new_value == 1
    
    def test_compare_headings_removed(self):
        """测试删除标题比较"""
        before_headings = [
            Heading(level=1, text="要删除的标题", style="标题 1", range_start=0, range_end=10)
        ]
        
        after_headings = []
        
        changes = self.comparator._compare_headings(before_headings, after_headings)
        
        assert len(changes) == 1
        assert changes[0].change_type == "heading_removed"
        assert changes[0].old_value == 1
        assert changes[0].new_value is None
    
    def test_compare_styles_usage_change(self):
        """测试样式使用状态变更比较"""
        before_styles = [
            Style(name="标题 1", type="paragraph", built_in=True, in_use=False)
        ]
        
        after_styles = [
            Style(name="标题 1", type="paragraph", built_in=True, in_use=True)
        ]
        
        changes = self.comparator._compare_styles(before_styles, after_styles)
        
        assert len(changes) == 1
        assert changes[0].change_type == "style_usage_change"
        assert changes[0].old_value is False
        assert changes[0].new_value is True
    
    def test_compare_snapshots(self):
        """测试完整快照比较"""
        # 创建前后快照
        before_structure = DocumentStructure(
            headings=[
                Heading(level=1, text="标题1", style="标题 1", range_start=0, range_end=10)
            ],
            styles=[
                Style(name="标题 1", type="paragraph", built_in=True, in_use=False)
            ],
            toc_entries=[],
            hyperlinks=[],
            references=[],
            page_count=1,
            word_count=100
        )
        
        after_structure = DocumentStructure(
            headings=[
                Heading(level=2, text="标题1", style="标题 2", range_start=0, range_end=10)
            ],
            styles=[
                Style(name="标题 1", type="paragraph", built_in=True, in_use=True)
            ],
            toc_entries=[],
            hyperlinks=[],
            references=[],
            page_count=1,
            word_count=100
        )
        
        before_snapshot = Mock()
        before_snapshot.structure = before_structure
        
        after_snapshot = Mock()
        after_snapshot.structure = after_structure
        
        changes = self.comparator.compare_snapshots(before_snapshot, after_snapshot)
        
        # 应该检测到标题级别变更、样式变更和样式使用状态变更
        assert len(changes) >= 2
        change_types = [c.change_type for c in changes]
        assert "heading_level_change" in change_types
        assert "style_usage_change" in change_types


class TestAuthorizationChecker:
    """测试授权检查器"""
    
    def setup_method(self):
        """测试前设置"""
        self.checker = AuthorizationChecker()
    
    def test_check_authorization_with_valid_task(self):
        """测试有效任务的授权检查"""
        # 创建格式变更
        changes = [
            FormatChange(
                change_type="heading_level_change",
                element_type="heading",
                element_id="heading_0_10",
                old_value=1,
                new_value=2,
                authorized=False
            )
        ]
        
        # 创建授权任务
        tasks = [
            Task(
                id="task_1",
                type=TaskType.SET_HEADING_LEVEL,
                source_comment_id="comment_1",
                locator=Locator(by=LocatorType.RANGE, value="0-10"),
                instruction="设置为2级标题"
            )
        ]
        
        # 创建批注
        comments = [
            Comment(
                id="comment_1",
                author="张三",
                page=1,
                anchor_text="标题",
                comment_text="改为2级标题",
                range_start=0,
                range_end=10
            )
        ]
        
        # 执行授权检查
        result_changes = self.checker.check_authorization(changes, tasks, comments)
        
        assert len(result_changes) == 1
        assert result_changes[0].authorized is True
        assert result_changes[0].source_comment_id == "comment_1"
    
    def test_check_authorization_without_valid_task(self):
        """测试无有效任务的授权检查"""
        # 创建格式变更
        changes = [
            FormatChange(
                change_type="heading_level_change",
                element_type="heading",
                element_id="heading_0_10",
                old_value=1,
                new_value=2,
                authorized=False
            )
        ]
        
        # 创建无授权的任务
        tasks = [
            Task(
                id="task_1",
                type=TaskType.SET_HEADING_LEVEL,
                source_comment_id=None,  # 无批注授权
                locator=Locator(by=LocatorType.RANGE, value="0-10"),
                instruction="设置为2级标题"
            )
        ]
        
        comments = []
        
        # 执行授权检查
        result_changes = self.checker.check_authorization(changes, tasks, comments)
        
        assert len(result_changes) == 1
        assert result_changes[0].authorized is False
        assert result_changes[0].source_comment_id is None
    
    def test_get_relevant_task_types(self):
        """测试获取相关任务类型"""
        relevant_types = self.checker._get_relevant_task_types("heading_level_change")
        assert TaskType.SET_HEADING_LEVEL in relevant_types
        
        relevant_types = self.checker._get_relevant_task_types("hyperlink_address_change")
        assert TaskType.REPLACE_HYPERLINK in relevant_types
        
        relevant_types = self.checker._get_relevant_task_types("unknown_change")
        assert len(relevant_types) == 0


class TestFormatValidator:
    """测试格式验证器"""
    
    def setup_method(self):
        """测试前设置"""
        self.validator = FormatValidator()
    
    def test_validate_execution_result_success(self):
        """测试成功的执行结果验证"""
        # 创建相同的前后快照（无变更）
        structure = DocumentStructure(
            headings=[],
            styles=[],
            toc_entries=[],
            hyperlinks=[],
            references=[],
            page_count=1,
            word_count=100
        )
        
        before_snapshot = Mock()
        before_snapshot.structure = structure
        
        after_snapshot = Mock()
        after_snapshot.structure = structure
        
        # 模拟比较器返回无变更
        with patch.object(self.validator.comparator, 'compare_snapshots', return_value=[]):
            report = self.validator.validate_execution_result(
                before_snapshot=before_snapshot,
                after_snapshot=after_snapshot,
                executed_tasks=[],
                comments=[]
            )
        
        assert report.is_valid is True
        assert len(report.authorized_changes) == 0
        assert len(report.unauthorized_changes) == 0
    
    def test_validate_execution_result_with_unauthorized_changes(self):
        """测试有未授权变更的执行结果验证"""
        before_snapshot = Mock()
        after_snapshot = Mock()
        
        # 模拟未授权的格式变更
        unauthorized_change = FormatChange(
            change_type="heading_level_change",
            element_type="heading",
            element_id="heading_1",
            old_value=1,
            new_value=2,
            authorized=False
        )
        
        with patch.object(self.validator.comparator, 'compare_snapshots', return_value=[unauthorized_change]):
            with patch.object(self.validator.auth_checker, 'check_authorization', return_value=[unauthorized_change]):
                report = self.validator.validate_execution_result(
                    before_snapshot=before_snapshot,
                    after_snapshot=after_snapshot,
                    executed_tasks=[],
                    comments=[]
                )
        
        assert report.is_valid is False
        assert len(report.unauthorized_changes) == 1
        assert len(report.warnings) > 0
        assert len(report.recommendations) > 0
    
    def test_should_rollback_with_unauthorized_changes(self):
        """测试有未授权变更时的回滚判断"""
        report = ValidationReport(
            is_valid=False,
            authorized_changes=[],
            unauthorized_changes=[Mock()],  # 有未授权变更
            warnings=[],
            recommendations=[],
            validation_time=datetime.now()
        )
        
        assert self.validator.should_rollback(report) is True
    
    def test_should_rollback_without_unauthorized_changes(self):
        """测试无未授权变更时的回滚判断"""
        report = ValidationReport(
            is_valid=True,
            authorized_changes=[Mock()],
            unauthorized_changes=[],  # 无未授权变更
            warnings=[],
            recommendations=[],
            validation_time=datetime.now()
        )
        
        assert self.validator.should_rollback(report) is False


class TestValidationReport:
    """测试验证报告"""
    
    def test_validation_report_properties(self):
        """测试验证报告属性"""
        authorized_changes = [Mock(), Mock()]
        unauthorized_changes = [Mock()]
        
        report = ValidationReport(
            is_valid=False,
            authorized_changes=authorized_changes,
            unauthorized_changes=unauthorized_changes,
            warnings=["警告1"],
            recommendations=["建议1"],
            validation_time=datetime.now()
        )
        
        assert report.total_changes == 3
        assert report.unauthorized_count == 1
        assert report.is_valid is False


if __name__ == "__main__":
    pytest.main([__file__])