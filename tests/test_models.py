"""
Test AutoWord Core Models
测试核心数据模型
"""

import pytest
from datetime import datetime
from pydantic import ValidationError

from autoword.core.models import (
    Comment, Task, Locator, TaskType, RiskLevel, LocatorType,
    DocumentStructure, TaskPlan, TaskResult, ExecutionResult
)


class TestComment:
    """测试 Comment 模型"""
    
    def test_valid_comment(self):
        """测试有效的批注"""
        comment = Comment(
            id="comment_1",
            author="张三",
            page=1,
            anchor_text="这是锚点文本",
            comment_text="这是批注内容",
            range_start=100,
            range_end=200
        )
        
        assert comment.id == "comment_1"
        assert comment.author == "张三"
        assert comment.page == 1
    
    def test_invalid_page_number(self):
        """测试无效的页码"""
        with pytest.raises(ValidationError):
            Comment(
                id="comment_1",
                author="张三",
                page=0,  # 无效页码
                anchor_text="锚点文本",
                comment_text="批注内容",
                range_start=100,
                range_end=200
            )


class TestLocator:
    """测试 Locator 模型"""
    
    def test_valid_locator(self):
        """测试有效的定位器"""
        locator = Locator(by=LocatorType.BOOKMARK, value="bookmark_1")
        assert locator.by == LocatorType.BOOKMARK
        assert locator.value == "bookmark_1"
    
    def test_empty_value(self):
        """测试空值"""
        with pytest.raises(ValidationError):
            Locator(by=LocatorType.BOOKMARK, value="   ")


class TestTask:
    """测试 Task 模型"""
    
    def test_valid_task(self):
        """测试有效的任务"""
        locator = Locator(by=LocatorType.BOOKMARK, value="bookmark_1")
        task = Task(
            id="task_1",
            type=TaskType.REWRITE,
            locator=locator,
            instruction="重写这段文本"
        )
        
        assert task.id == "task_1"
        assert task.type == TaskType.REWRITE
        assert task.risk == RiskLevel.LOW  # 默认值
        assert task.requires_user_review is False  # 默认值
    
    def test_empty_instruction(self):
        """测试空指令"""
        locator = Locator(by=LocatorType.BOOKMARK, value="bookmark_1")
        with pytest.raises(ValidationError):
            Task(
                id="task_1",
                type=TaskType.REWRITE,
                locator=locator,
                instruction="   "  # 空指令
            )


class TestTaskPlan:
    """测试 TaskPlan 模型"""
    
    def test_valid_task_plan(self):
        """测试有效的任务计划"""
        locator = Locator(by=LocatorType.BOOKMARK, value="bookmark_1")
        task = Task(
            id="task_1",
            type=TaskType.REWRITE,
            locator=locator,
            instruction="重写文本"
        )
        
        plan = TaskPlan(
            tasks=[task],
            total_tasks=1
        )
        
        assert len(plan.tasks) == 1
        assert plan.total_tasks == 1
        assert isinstance(plan.created_at, datetime)
    
    def test_mismatched_task_count(self):
        """测试任务数量不匹配"""
        locator = Locator(by=LocatorType.BOOKMARK, value="bookmark_1")
        task = Task(
            id="task_1",
            type=TaskType.REWRITE,
            locator=locator,
            instruction="重写文本"
        )
        
        with pytest.raises(ValidationError):
            TaskPlan(
                tasks=[task],
                total_tasks=2  # 数量不匹配
            )


class TestTaskResult:
    """测试 TaskResult 模型"""
    
    def test_successful_result(self):
        """测试成功结果"""
        result = TaskResult(
            task_id="task_1",
            success=True,
            message="任务执行成功",
            execution_time=1.5
        )
        
        assert result.success is True
        assert result.execution_time == 1.5
        assert result.error_details is None
    
    def test_failed_result(self):
        """测试失败结果"""
        result = TaskResult(
            task_id="task_1",
            success=False,
            message="任务执行失败",
            execution_time=0.5,
            error_details="详细错误信息"
        )
        
        assert result.success is False
        assert result.error_details == "详细错误信息"


if __name__ == "__main__":
    pytest.main([__file__])