"""
Test AutoWord Task Planner
测试任务规划器
"""

import json
import pytest
from unittest.mock import Mock, patch, MagicMock
from datetime import datetime

from autoword.core.planner import (
    TaskPlanner, FormatProtectionGuard, TaskDependencyResolver, 
    RiskAssessment, PlanningResult, create_task_plan
)
from autoword.core.models import (
    Comment, DocumentStructure, Task, TaskType, RiskLevel, 
    LocatorType, Locator, ValidationResult
)
from autoword.core.llm_client import ModelType
from autoword.core.exceptions import LLMError, ValidationError, FormatProtectionError


class TestFormatProtectionGuard:
    """测试格式保护守卫"""
    
    def setup_method(self):
        """测试前设置"""
        self.guard = FormatProtectionGuard()
    
    def test_validate_task_authorization_format_task_with_comment(self):
        """测试格式化任务有批注授权"""
        task_data = {
            "id": "task_1",
            "type": "set_paragraph_style",
            "source_comment_id": "comment_1",
            "instruction": "设置段落样式"
        }
        
        is_authorized, error_msg = self.guard.validate_task_authorization(task_data)
        
        assert is_authorized is True
        assert error_msg == ""
    
    def test_validate_task_authorization_format_task_without_comment(self):
        """测试格式化任务无批注授权"""
        task_data = {
            "id": "task_1",
            "type": "set_paragraph_style",
            "source_comment_id": None,
            "instruction": "设置段落样式"
        }
        
        is_authorized, error_msg = self.guard.validate_task_authorization(task_data)
        
        assert is_authorized is False
        assert "格式化任务" in error_msg
        assert "缺少批注授权" in error_msg
    
    def test_validate_task_authorization_content_task_without_comment(self):
        """测试内容任务无批注授权（应该通过）"""
        task_data = {
            "id": "task_1",
            "type": "rewrite",
            "source_comment_id": None,
            "instruction": "重写内容"
        }
        
        is_authorized, error_msg = self.guard.validate_task_authorization(task_data)
        
        assert is_authorized is True
        assert error_msg == ""
    
    def test_validate_task_authorization_unknown_task_without_comment(self):
        """测试未知任务类型无批注授权"""
        task_data = {
            "id": "task_1",
            "type": "unknown_task",
            "source_comment_id": None,
            "instruction": "未知任务"
        }
        
        is_authorized, error_msg = self.guard.validate_task_authorization(task_data)
        
        assert is_authorized is False
        assert "不在白名单中" in error_msg
    
    def test_filter_unauthorized_tasks(self):
        """测试过滤未授权任务"""
        tasks = [
            {
                "id": "task_1",
                "type": "rewrite",
                "source_comment_id": None,
                "instruction": "重写内容"
            },
            {
                "id": "task_2", 
                "type": "set_paragraph_style",
                "source_comment_id": "comment_1",
                "instruction": "设置样式"
            },
            {
                "id": "task_3",
                "type": "set_paragraph_style",
                "source_comment_id": None,
                "instruction": "设置样式（无授权）"
            }
        ]
        
        authorized, filtered = self.guard.filter_unauthorized_tasks(tasks)
        
        assert len(authorized) == 2
        assert len(filtered) == 1
        
        # 验证授权任务
        authorized_ids = [t["id"] for t in authorized]
        assert "task_1" in authorized_ids
        assert "task_2" in authorized_ids
        
        # 验证被过滤任务
        assert filtered[0]["id"] == "task_3"
        assert filtered[0]["filtered"] is True
        assert "格式化任务" in filtered[0]["filter_reason"]
    
    def test_validate_task_safety_format_task_with_valid_comment(self):
        """测试格式化任务的安全性验证（有效批注）"""
        task = Task(
            id="task_1",
            type=TaskType.SET_PARAGRAPH_STYLE,
            source_comment_id="comment_1",
            locator=Locator(by=LocatorType.FIND, value="test"),
            instruction="设置样式"
        )
        
        comments = [
            Comment(
                id="comment_1",
                author="张三",
                page=1,
                anchor_text="测试",
                comment_text="设置样式",
                range_start=0,
                range_end=10
            )
        ]
        
        result = self.guard.validate_task_safety(task, comments)
        
        assert result.is_valid is True
        assert len(result.errors) == 0
    
    def test_validate_task_safety_format_task_with_invalid_comment(self):
        """测试格式化任务的安全性验证（无效批注）"""
        task = Task(
            id="task_1",
            type=TaskType.SET_PARAGRAPH_STYLE,
            source_comment_id="nonexistent_comment",
            locator=Locator(by=LocatorType.FIND, value="test"),
            instruction="设置样式"
        )
        
        comments = [
            Comment(
                id="comment_1",
                author="张三",
                page=1,
                anchor_text="测试",
                comment_text="设置样式",
                range_start=0,
                range_end=10
            )
        ]
        
        result = self.guard.validate_task_safety(task, comments)
        
        assert result.is_valid is False
        assert len(result.errors) == 1
        assert "不存在" in result.errors[0]


class TestTaskDependencyResolver:
    """测试任务依赖解析器"""
    
    def setup_method(self):
        """测试前设置"""
        self.resolver = TaskDependencyResolver()
    
    def test_resolve_dependencies_no_dependencies(self):
        """测试无依赖任务的解析"""
        tasks = [
            Task(
                id="task_1",
                type=TaskType.REWRITE,
                locator=Locator(by=LocatorType.FIND, value="test1"),
                instruction="任务1"
            ),
            Task(
                id="task_2",
                type=TaskType.INSERT,
                locator=Locator(by=LocatorType.FIND, value="test2"),
                instruction="任务2"
            )
        ]
        
        sorted_tasks = self.resolver.resolve_dependencies(tasks)
        
        assert len(sorted_tasks) == 2
        # 应该按风险级别和类型排序
        assert sorted_tasks[0].type in [TaskType.REWRITE, TaskType.INSERT]
    
    def test_resolve_dependencies_with_dependencies(self):
        """测试有依赖关系的任务解析"""
        tasks = [
            Task(
                id="task_1",
                type=TaskType.REWRITE,
                locator=Locator(by=LocatorType.FIND, value="test1"),
                instruction="任务1",
                dependencies=["task_2"]  # 依赖 task_2
            ),
            Task(
                id="task_2",
                type=TaskType.INSERT,
                locator=Locator(by=LocatorType.FIND, value="test2"),
                instruction="任务2"
            )
        ]
        
        sorted_tasks = self.resolver.resolve_dependencies(tasks)
        
        assert len(sorted_tasks) == 2
        # task_2 应该在 task_1 之前
        task_ids = [t.id for t in sorted_tasks]
        assert task_ids.index("task_2") < task_ids.index("task_1")
    
    def test_resolve_dependencies_circular_dependency(self):
        """测试循环依赖的处理"""
        tasks = [
            Task(
                id="task_1",
                type=TaskType.REWRITE,
                locator=Locator(by=LocatorType.FIND, value="test1"),
                instruction="任务1",
                dependencies=["task_2"]
            ),
            Task(
                id="task_2",
                type=TaskType.INSERT,
                locator=Locator(by=LocatorType.FIND, value="test2"),
                instruction="任务2",
                dependencies=["task_1"]  # 循环依赖
            )
        ]
        
        # 应该能处理循环依赖而不崩溃
        sorted_tasks = self.resolver.resolve_dependencies(tasks)
        
        assert len(sorted_tasks) == 2


class TestRiskAssessment:
    """测试风险评估器"""
    
    def setup_method(self):
        """测试前设置"""
        self.assessor = RiskAssessment()
    
    def test_assess_task_risk_high_risk(self):
        """测试高风险任务评估"""
        task_data = {"type": "apply_template"}
        
        risk = self.assessor.assess_task_risk(task_data)
        
        assert risk == RiskLevel.HIGH
    
    def test_assess_task_risk_medium_risk(self):
        """测试中等风险任务评估"""
        task_data = {"type": "set_heading_level"}
        
        risk = self.assessor.assess_task_risk(task_data)
        
        assert risk == RiskLevel.MEDIUM
    
    def test_assess_task_risk_low_risk(self):
        """测试低风险任务评估"""
        task_data = {"type": "rewrite"}
        
        risk = self.assessor.assess_task_risk(task_data)
        
        assert risk == RiskLevel.LOW
    
    def test_assess_task_risk_unknown_type(self):
        """测试未知任务类型的风险评估"""
        task_data = {"type": "unknown_task"}
        
        risk = self.assessor.assess_task_risk(task_data)
        
        assert risk == RiskLevel.MEDIUM  # 默认中等风险
    
    def test_assess_batch_risk_high_risk_batch(self):
        """测试高风险批次评估"""
        tasks = [
            {"type": "apply_template"},
            {"type": "rebuild_toc"},
            {"type": "rewrite"}
        ]
        
        report = self.assessor.assess_batch_risk(tasks)
        
        assert report["overall_risk"] == RiskLevel.HIGH
        assert report["risk_distribution"]["high"] == 2
        assert report["risk_distribution"]["low"] == 1
        assert report["high_risk_ratio"] > 0.3
    
    def test_assess_batch_risk_low_risk_batch(self):
        """测试低风险批次评估"""
        tasks = [
            {"type": "rewrite"},
            {"type": "insert"},
            {"type": "delete"}
        ]
        
        report = self.assessor.assess_batch_risk(tasks)
        
        assert report["overall_risk"] == RiskLevel.LOW
        assert report["risk_distribution"]["low"] == 3
        assert report["high_risk_ratio"] == 0.0


class TestTaskPlanner:
    """测试任务规划器"""
    
    def setup_method(self):
        """测试前设置"""
        self.planner = TaskPlanner()
    
    def teardown_method(self):
        """测试后清理"""
        self.planner.close()
    
    @patch.object(TaskPlanner, '_call_llm')
    def test_generate_plan_success(self, mock_call_llm):
        """测试成功生成任务计划"""
        # 模拟 LLM 响应
        mock_llm_response = {
            "tasks": [
                {
                    "id": "task_1",
                    "type": "rewrite",
                    "source_comment_id": "comment_1",
                    "locator": {"by": "find", "value": "test text"},
                    "instruction": "重写这段文本",
                    "risk": "low"
                }
            ]
        }
        mock_call_llm.return_value = json.dumps(mock_llm_response)
        
        # 创建测试数据
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
                anchor_text="test text",
                comment_text="需要重写",
                range_start=0,
                range_end=10
            )
        ]
        
        # 执行测试
        result = self.planner.generate_plan(structure, comments)
        
        # 验证结果
        assert result.success is True
        assert result.task_plan is not None
        assert len(result.task_plan.tasks) == 1
        assert result.task_plan.tasks[0].id == "task_1"
        assert result.task_plan.tasks[0].type == TaskType.REWRITE
    
    @patch.object(TaskPlanner, '_call_llm')
    def test_generate_plan_with_format_protection(self, mock_call_llm):
        """测试格式保护功能"""
        # 模拟 LLM 响应（包含未授权的格式化任务）
        mock_llm_response = {
            "tasks": [
                {
                    "id": "task_1",
                    "type": "rewrite",
                    "source_comment_id": "comment_1",
                    "locator": {"by": "find", "value": "test"},
                    "instruction": "重写内容"
                },
                {
                    "id": "task_2",
                    "type": "set_paragraph_style",
                    "source_comment_id": None,  # 无批注授权
                    "locator": {"by": "find", "value": "test"},
                    "instruction": "设置样式"
                }
            ]
        }
        mock_call_llm.return_value = json.dumps(mock_llm_response)
        
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
                anchor_text="test",
                comment_text="重写内容",
                range_start=0,
                range_end=10
            )
        ]
        
        # 执行测试
        result = self.planner.generate_plan(structure, comments)
        
        # 验证结果
        assert result.success is True
        assert len(result.task_plan.tasks) == 1  # 只有授权任务
        assert result.task_plan.tasks[0].type == TaskType.REWRITE
        assert len(result.skipped_tasks) == 1  # 一个被过滤的任务
    
    @patch.object(TaskPlanner, '_call_llm')
    def test_generate_plan_llm_error(self, mock_call_llm):
        """测试 LLM 调用错误"""
        mock_call_llm.side_effect = LLMError("LLM 服务不可用")
        
        structure = DocumentStructure(
            headings=[],
            styles=[],
            toc_entries=[],
            hyperlinks=[],
            references=[],
            page_count=1,
            word_count=100
        )
        
        result = self.planner.generate_plan(structure, [])
        
        assert result.success is False
        assert "LLM 服务不可用" in result.error_message
    
    def test_parse_llm_response_valid(self):
        """测试解析有效的 LLM 响应"""
        response = json.dumps({
            "tasks": [
                {
                    "id": "task_1",
                    "type": "rewrite",
                    "instruction": "重写内容"
                }
            ]
        })
        
        tasks = self.planner._parse_llm_response(response)
        
        assert len(tasks) == 1
        assert tasks[0]["id"] == "task_1"
    
    def test_parse_llm_response_invalid_json(self):
        """测试解析无效的 JSON 响应"""
        response = "这不是有效的 JSON"
        
        with pytest.raises(ValidationError, match="不是有效的 JSON"):
            self.planner._parse_llm_response(response)
    
    def test_parse_llm_response_missing_tasks(self):
        """测试缺少 tasks 字段的响应"""
        response = json.dumps({"result": "success"})
        
        with pytest.raises(ValidationError, match="缺少 'tasks' 字段"):
            self.planner._parse_llm_response(response)
    
    def test_convert_to_task_objects(self):
        """测试转换任务对象"""
        tasks_data = [
            {
                "id": "task_1",
                "type": "rewrite",
                "source_comment_id": "comment_1",
                "locator": {"by": "find", "value": "test"},
                "instruction": "重写内容",
                "risk": "low"
            }
        ]
        
        comments = [
            Comment(
                id="comment_1",
                author="张三",
                page=1,
                anchor_text="test",
                comment_text="重写",
                range_start=0,
                range_end=10
            )
        ]
        
        task_objects = self.planner._convert_to_task_objects(tasks_data, comments)
        
        assert len(task_objects) == 1
        assert isinstance(task_objects[0], Task)
        assert task_objects[0].id == "task_1"
        assert task_objects[0].type == TaskType.REWRITE
    
    def test_validate_plan(self):
        """测试验证任务计划"""
        tasks = [
            Task(
                id="task_1",
                type=TaskType.REWRITE,
                source_comment_id="comment_1",
                locator=Locator(by=LocatorType.FIND, value="test"),
                instruction="重写内容"
            )
        ]
        
        from autoword.core.models import TaskPlan
        task_plan = TaskPlan(tasks=tasks, total_tasks=1)
        
        comments = [
            Comment(
                id="comment_1",
                author="张三",
                page=1,
                anchor_text="test",
                comment_text="重写",
                range_start=0,
                range_end=10
            )
        ]
        
        result = self.planner.validate_plan(task_plan, comments)
        
        assert result.is_valid is True
        assert len(result.errors) == 0


class TestConvenienceFunction:
    """测试便捷函数"""
    
    @patch('autoword.core.planner.TaskPlanner')
    def test_create_task_plan(self, mock_planner_class):
        """测试创建任务计划便捷函数"""
        mock_planner = Mock()
        mock_result = PlanningResult(success=True)
        mock_planner.generate_plan.return_value = mock_result
        mock_planner_class.return_value = mock_planner
        
        structure = DocumentStructure(
            headings=[],
            styles=[],
            toc_entries=[],
            hyperlinks=[],
            references=[],
            page_count=1,
            word_count=100
        )
        
        result = create_task_plan(structure, [])
        
        assert result.success is True
        mock_planner.generate_plan.assert_called_once()
        mock_planner.close.assert_called_once()


if __name__ == "__main__":
    pytest.main([__file__])