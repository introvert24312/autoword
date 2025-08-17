"""
Test AutoWord vNext Document Planner
测试 vNext 文档规划器
"""

import json
import pytest
from unittest.mock import Mock, patch, MagicMock
from pathlib import Path
from datetime import datetime

from autoword.vnext.planner import DocumentPlanner
from autoword.vnext.models import (
    StructureV1, PlanV1, ValidationResult, DocumentMetadata,
    StyleDefinition, ParagraphSkeleton, HeadingReference,
    DeleteSectionByHeading, UpdateToc, SetStyleRule, FontSpec, ParagraphSpec,
    LineSpacingMode, StyleType, ClearDirectFormatting
)
from autoword.vnext.exceptions import (
    PlanningError, SchemaValidationError, WhitelistViolationError
)
from autoword.core.llm_client import LLMClient, LLMResponse, ModelType


class TestDocumentPlanner:
    """测试文档规划器"""
    
    def setup_method(self):
        """测试前设置"""
        # Create mock LLM client
        self.mock_llm_client = Mock(spec=LLMClient)
        
        # Create test schema path
        self.schema_path = Path(__file__).parent.parent / "autoword" / "vnext" / "schemas" / "plan.v1.json"
        
        # Initialize planner with mock client
        self.planner = DocumentPlanner(llm_client=self.mock_llm_client)
        
        # Create test structure
        self.test_structure = StructureV1(
            metadata=DocumentMetadata(
                title="测试文档",
                author="测试作者",
                page_count=5,
                paragraph_count=20,
                word_count=1000
            ),
            styles=[
                StyleDefinition(
                    name="标题 1",
                    type=StyleType.PARAGRAPH,
                    font=FontSpec(east_asian="楷体", latin="Times New Roman", size_pt=14, bold=True),
                    paragraph=ParagraphSpec(line_spacing_mode=LineSpacingMode.MULTIPLE, line_spacing_value=2.0)
                ),
                StyleDefinition(
                    name="正文",
                    type=StyleType.PARAGRAPH,
                    font=FontSpec(east_asian="宋体", latin="Times New Roman", size_pt=12),
                    paragraph=ParagraphSpec(line_spacing_mode=LineSpacingMode.MULTIPLE, line_spacing_value=2.0)
                )
            ],
            paragraphs=[
                ParagraphSkeleton(index=0, style_name="标题 1", preview_text="摘要", is_heading=True, heading_level=1),
                ParagraphSkeleton(index=1, style_name="正文", preview_text="这是摘要内容...", is_heading=False),
                ParagraphSkeleton(index=2, style_name="标题 1", preview_text="1. 引言", is_heading=True, heading_level=1),
                ParagraphSkeleton(index=3, style_name="正文", preview_text="这是引言内容...", is_heading=False)
            ],
            headings=[
                HeadingReference(paragraph_index=0, level=1, text="摘要", style_name="标题 1"),
                HeadingReference(paragraph_index=2, level=1, text="1. 引言", style_name="标题 1")
            ]
        )
    
    def test_init_default_llm_client(self):
        """测试默认LLM客户端初始化"""
        planner = DocumentPlanner()
        assert planner.llm_client is not None
        assert isinstance(planner.llm_client, LLMClient)
    
    def test_init_custom_llm_client(self):
        """测试自定义LLM客户端初始化"""
        custom_client = Mock(spec=LLMClient)
        planner = DocumentPlanner(llm_client=custom_client)
        assert planner.llm_client is custom_client
    
    def test_init_invalid_schema_path(self):
        """测试无效schema路径"""
        with pytest.raises(PlanningError, match="Failed to load plan schema"):
            DocumentPlanner(schema_path="/nonexistent/path.json")
    
    def test_generate_plan_success(self):
        """测试成功生成计划"""
        # Mock successful LLM response
        mock_plan = {
            "schema_version": "plan.v1",
            "ops": [
                {
                    "operation_type": "delete_section_by_heading",
                    "heading_text": "摘要",
                    "level": 1,
                    "match": "EXACT",
                    "case_sensitive": False
                },
                {
                    "operation_type": "update_toc"
                }
            ]
        }
        
        mock_response = LLMResponse(
            success=True,
            content=json.dumps(mock_plan),
            model="gpt-4o"
        )
        self.mock_llm_client.call_with_json_retry.return_value = mock_response
        
        # Execute test
        result = self.planner.generate_plan(self.test_structure, "删除摘要部分并更新目录")
        
        # Verify result
        assert isinstance(result, PlanV1)
        assert result.schema_version == "plan.v1"
        assert len(result.ops) == 2
        
        # Verify first operation
        first_op = result.ops[0]
        assert isinstance(first_op, DeleteSectionByHeading)
        assert first_op.heading_text == "摘要"
        assert first_op.level == 1
        
        # Verify second operation
        second_op = result.ops[1]
        assert isinstance(second_op, UpdateToc)
    
    def test_generate_plan_llm_failure(self):
        """测试LLM调用失败"""
        mock_response = LLMResponse(
            success=False,
            content="",
            model="gpt-4o",
            error="API调用失败"
        )
        self.mock_llm_client.call_with_json_retry.return_value = mock_response
        
        with pytest.raises(PlanningError, match="LLM call failed"):
            self.planner.generate_plan(self.test_structure, "测试意图")
    
    def test_generate_plan_invalid_json(self):
        """测试无效JSON响应"""
        mock_response = LLMResponse(
            success=True,
            content="这不是有效的JSON",
            model="gpt-4o"
        )
        self.mock_llm_client.call_with_json_retry.return_value = mock_response
        
        with pytest.raises(PlanningError, match="Invalid JSON response"):
            self.planner.generate_plan(self.test_structure, "测试意图")
    
    def test_generate_plan_schema_validation_failure(self):
        """测试schema验证失败"""
        # Invalid plan missing required fields
        invalid_plan = {
            "schema_version": "plan.v1",
            "ops": [
                {
                    "operation_type": "delete_section_by_heading"
                    # Missing required fields: heading_text, level
                }
            ]
        }
        
        mock_response = LLMResponse(
            success=True,
            content=json.dumps(invalid_plan),
            model="gpt-4o"
        )
        self.mock_llm_client.call_with_json_retry.return_value = mock_response
        
        with pytest.raises(SchemaValidationError, match="Plan failed schema validation"):
            self.planner.generate_plan(self.test_structure, "测试意图")
    
    def test_generate_plan_whitelist_violation(self):
        """测试白名单违规"""
        # Plan with non-whitelisted operation (but valid schema structure)
        invalid_plan = {
            "schema_version": "plan.v1",
            "ops": [
                {
                    "operation_type": "unauthorized_operation",
                    "some_param": "value"
                }
            ]
        }
        
        mock_response = LLMResponse(
            success=True,
            content=json.dumps(invalid_plan),
            model="gpt-4o"
        )
        self.mock_llm_client.call_with_json_retry.return_value = mock_response
        
        # This should fail at schema validation stage since unauthorized_operation is not in the schema
        with pytest.raises(SchemaValidationError, match="Plan failed schema validation"):
            self.planner.generate_plan(self.test_structure, "测试意图")
    
    def test_validate_plan_schema_valid(self):
        """测试有效计划的schema验证"""
        valid_plan = {
            "schema_version": "plan.v1",
            "ops": [
                {
                    "operation_type": "delete_section_by_heading",
                    "heading_text": "摘要",
                    "level": 1,
                    "match": "EXACT",
                    "case_sensitive": False
                }
            ]
        }
        
        result = self.planner.validate_plan_schema(valid_plan)
        
        assert result.is_valid is True
        assert len(result.errors) == 0
    
    def test_validate_plan_schema_invalid(self):
        """测试无效计划的schema验证"""
        invalid_plan = {
            "schema_version": "plan.v1",
            "ops": [
                {
                    "operation_type": "delete_section_by_heading",
                    "heading_text": "",  # Empty string not allowed
                    "level": 10,  # Level > 9 not allowed
                    "match": "INVALID_MATCH"  # Invalid match mode
                }
            ]
        }
        
        result = self.planner.validate_plan_schema(invalid_plan)
        
        assert result.is_valid is False
        assert len(result.errors) > 0
    
    def test_validate_plan_schema_missing_operation_type(self):
        """测试缺少operation_type的验证"""
        invalid_plan = {
            "schema_version": "plan.v1",
            "ops": [
                {
                    "heading_text": "摘要",
                    "level": 1
                    # Missing operation_type
                }
            ]
        }
        
        result = self.planner.validate_plan_schema(invalid_plan)
        
        assert result.is_valid is False
        # Check that there are validation errors (the specific message may vary)
        assert len(result.errors) > 0
    
    def test_check_whitelist_compliance_valid(self):
        """测试白名单合规性检查 - 有效操作"""
        valid_plan = PlanV1(ops=[
            DeleteSectionByHeading(heading_text="摘要", level=1),
            UpdateToc()
        ])
        
        result = self.planner.check_whitelist_compliance(valid_plan)
        
        assert result.is_valid is True
        assert len(result.errors) == 0
    
    def test_check_whitelist_compliance_invalid(self):
        """测试白名单合规性检查 - 无效操作"""
        # Test the whitelist compliance directly with a plan that has valid operations
        # but we'll test the method with a modified operation type
        valid_plan = PlanV1(ops=[
            DeleteSectionByHeading(heading_text="摘要", level=1)
        ])
        
        # Temporarily modify the operation type to test whitelist checking
        original_op_type = valid_plan.ops[0].operation_type
        valid_plan.ops[0].operation_type = "unauthorized_operation"
        
        result = self.planner.check_whitelist_compliance(valid_plan)
        
        # Restore original operation type
        valid_plan.ops[0].operation_type = original_op_type
        
        assert result.is_valid is False
        assert any("not whitelisted" in error for error in result.errors)
    
    def test_check_whitelist_compliance_clear_formatting_without_auth(self):
        """测试clear_direct_formatting操作缺少授权"""
        # Create a valid clear_direct_formatting operation first
        clear_op = ClearDirectFormatting(scope="document", authorization_required=True)
        valid_plan = PlanV1(ops=[clear_op])
        
        # Temporarily modify the authorization to test the check
        original_auth = valid_plan.ops[0].authorization_required
        valid_plan.ops[0].authorization_required = False
        
        result = self.planner.check_whitelist_compliance(valid_plan)
        
        # Restore original authorization
        valid_plan.ops[0].authorization_required = original_auth
        
        assert result.is_valid is False
        assert any("requires authorization" in error for error in result.errors)
    
    def test_build_system_prompt(self):
        """测试系统提示词构建"""
        system_prompt = self.planner._build_system_prompt()
        
        assert "plan.v1 schema" in system_prompt
        assert "whitelisted operations" in system_prompt
        assert "delete_section_by_heading" in system_prompt
        assert "update_toc" in system_prompt
        assert "JSON" in system_prompt
    
    def test_build_user_prompt(self):
        """测试用户提示词构建"""
        user_intent = "删除摘要部分"
        user_prompt = self.planner._build_user_prompt(self.test_structure, user_intent)
        
        assert "Document Structure:" in user_prompt
        assert user_intent in user_prompt
        assert "schema_version" in user_prompt  # Structure JSON should be included
        assert "plan.v1 schema" in user_prompt
    
    def test_validate_operation_constraints_delete_section_valid(self):
        """测试delete_section_by_heading操作约束验证 - 有效"""
        op = {
            "operation_type": "delete_section_by_heading",
            "heading_text": "摘要",
            "level": 1,
            "match": "EXACT"
        }
        
        errors = self.planner._validate_operation_constraints(op, "delete_section_by_heading")
        
        assert len(errors) == 0
    
    def test_validate_operation_constraints_delete_section_invalid(self):
        """测试delete_section_by_heading操作约束验证 - 无效"""
        op = {
            "operation_type": "delete_section_by_heading",
            "heading_text": "",  # Empty text
            "level": 10,  # Invalid level
            "match": "INVALID"  # Invalid match mode
        }
        
        errors = self.planner._validate_operation_constraints(op, "delete_section_by_heading")
        
        assert len(errors) >= 3
        assert any("heading_text" in error for error in errors)
        assert any("level must be" in error for error in errors)
        assert any("match must be" in error for error in errors)
    
    def test_validate_operation_constraints_set_style_rule_valid(self):
        """测试set_style_rule操作约束验证 - 有效"""
        op = {
            "operation_type": "set_style_rule",
            "target_style_name": "标题 1",
            "font": {
                "size_pt": 14,
                "bold": True
            }
        }
        
        errors = self.planner._validate_operation_constraints(op, "set_style_rule")
        
        assert len(errors) == 0
    
    def test_validate_operation_constraints_set_style_rule_invalid(self):
        """测试set_style_rule操作约束验证 - 无效"""
        op = {
            "operation_type": "set_style_rule",
            "target_style_name": "",  # Empty name
            "font": {
                "size_pt": 100  # Invalid size
            }
        }
        
        errors = self.planner._validate_operation_constraints(op, "set_style_rule")
        
        assert len(errors) >= 2
        assert any("target_style_name" in error for error in errors)
        assert any("size_pt must be" in error for error in errors)
    
    def test_validate_operation_constraints_reassign_paragraphs_valid(self):
        """测试reassign_paragraphs_to_style操作约束验证 - 有效"""
        op = {
            "operation_type": "reassign_paragraphs_to_style",
            "selector": {"style_name": "旧样式"},
            "target_style_name": "新样式"
        }
        
        errors = self.planner._validate_operation_constraints(op, "reassign_paragraphs_to_style")
        
        assert len(errors) == 0
    
    def test_validate_operation_constraints_reassign_paragraphs_invalid(self):
        """测试reassign_paragraphs_to_style操作约束验证 - 无效"""
        op = {
            "operation_type": "reassign_paragraphs_to_style",
            "selector": {},  # Empty selector
            "target_style_name": ""  # Empty target
        }
        
        errors = self.planner._validate_operation_constraints(op, "reassign_paragraphs_to_style")
        
        assert len(errors) >= 2
        assert any("selector" in error for error in errors)
        assert any("target_style_name" in error for error in errors)
    
    def test_validate_operation_constraints_clear_formatting_valid(self):
        """测试clear_direct_formatting操作约束验证 - 有效"""
        op = {
            "operation_type": "clear_direct_formatting",
            "scope": "document",
            "authorization_required": True
        }
        
        errors = self.planner._validate_operation_constraints(op, "clear_direct_formatting")
        
        assert len(errors) == 0
    
    def test_validate_operation_constraints_clear_formatting_invalid(self):
        """测试clear_direct_formatting操作约束验证 - 无效"""
        op = {
            "operation_type": "clear_direct_formatting",
            "scope": "invalid_scope",  # Invalid scope
            "authorization_required": False  # Should be True
        }
        
        errors = self.planner._validate_operation_constraints(op, "clear_direct_formatting")
        
        assert len(errors) >= 2
        assert any("scope must be" in error for error in errors)
        assert any("authorization_required must be true" in error for error in errors)
    
    def test_whitelisted_operations_constant(self):
        """测试白名单操作常量"""
        expected_operations = {
            "delete_section_by_heading",
            "update_toc",
            "delete_toc",
            "set_style_rule",
            "reassign_paragraphs_to_style",
            "clear_direct_formatting"
        }
        
        assert self.planner.WHITELISTED_OPERATIONS == expected_operations


class TestDocumentPlannerIntegration:
    """测试文档规划器集成"""
    
    def setup_method(self):
        """测试前设置"""
        # Use real LLM client for integration tests
        self.planner = DocumentPlanner()
        
        # Create comprehensive test structure
        self.test_structure = StructureV1(
            metadata=DocumentMetadata(
                title="学术论文",
                author="研究者",
                page_count=10,
                paragraph_count=50,
                word_count=5000
            ),
            styles=[
                StyleDefinition(
                    name="标题 1",
                    type=StyleType.PARAGRAPH,
                    font=FontSpec(east_asian="楷体", size_pt=14, bold=True),
                    paragraph=ParagraphSpec(line_spacing_mode=LineSpacingMode.MULTIPLE, line_spacing_value=2.0)
                ),
                StyleDefinition(
                    name="标题 2", 
                    type=StyleType.PARAGRAPH,
                    font=FontSpec(east_asian="宋体", size_pt=12, bold=True),
                    paragraph=ParagraphSpec(line_spacing_mode=LineSpacingMode.MULTIPLE, line_spacing_value=2.0)
                ),
                StyleDefinition(
                    name="正文",
                    type=StyleType.PARAGRAPH,
                    font=FontSpec(east_asian="宋体", size_pt=12),
                    paragraph=ParagraphSpec(line_spacing_mode=LineSpacingMode.MULTIPLE, line_spacing_value=2.0)
                )
            ],
            headings=[
                HeadingReference(paragraph_index=0, level=1, text="摘要", style_name="标题 1"),
                HeadingReference(paragraph_index=5, level=1, text="1. 引言", style_name="标题 1"),
                HeadingReference(paragraph_index=10, level=2, text="1.1 研究背景", style_name="标题 2"),
                HeadingReference(paragraph_index=20, level=1, text="参考文献", style_name="标题 1")
            ]
        )
    
    @pytest.mark.integration
    @patch.object(LLMClient, 'call_with_json_retry')
    def test_generate_plan_realistic_scenario(self, mock_llm_call):
        """测试真实场景的计划生成"""
        # Mock realistic LLM response
        realistic_plan = {
            "schema_version": "plan.v1",
            "ops": [
                {
                    "operation_type": "delete_section_by_heading",
                    "heading_text": "摘要",
                    "level": 1,
                    "match": "EXACT",
                    "case_sensitive": False
                },
                {
                    "operation_type": "delete_section_by_heading",
                    "heading_text": "参考文献",
                    "level": 1,
                    "match": "EXACT",
                    "case_sensitive": False
                },
                {
                    "operation_type": "set_style_rule",
                    "target_style_name": "标题 1",
                    "font": {
                        "east_asian": "楷体",
                        "size_pt": 12,
                        "bold": True
                    },
                    "paragraph": {
                        "line_spacing_mode": "MULTIPLE",
                        "line_spacing_value": 2.0
                    }
                },
                {
                    "operation_type": "set_style_rule",
                    "target_style_name": "标题 2",
                    "font": {
                        "east_asian": "宋体",
                        "size_pt": 12,
                        "bold": True
                    },
                    "paragraph": {
                        "line_spacing_mode": "MULTIPLE",
                        "line_spacing_value": 2.0
                    }
                },
                {
                    "operation_type": "set_style_rule",
                    "target_style_name": "正文",
                    "font": {
                        "east_asian": "宋体",
                        "size_pt": 12,
                        "bold": False
                    },
                    "paragraph": {
                        "line_spacing_mode": "MULTIPLE",
                        "line_spacing_value": 2.0
                    }
                },
                {
                    "operation_type": "update_toc"
                }
            ]
        }
        
        mock_response = LLMResponse(
            success=True,
            content=json.dumps(realistic_plan),
            model="gpt-4o"
        )
        mock_llm_call.return_value = mock_response
        
        # Execute test
        user_intent = "删除摘要和参考文献部分，设置标准的学术论文格式（H1楷体12pt粗体，H2宋体12pt粗体，正文宋体12pt，行距2倍），最后更新目录"
        
        result = self.planner.generate_plan(self.test_structure, user_intent)
        
        # Verify comprehensive result
        assert isinstance(result, PlanV1)
        assert len(result.ops) == 6
        
        # Verify delete operations
        delete_ops = [op for op in result.ops if op.operation_type == "delete_section_by_heading"]
        assert len(delete_ops) == 2
        
        # Verify style operations
        style_ops = [op for op in result.ops if op.operation_type == "set_style_rule"]
        assert len(style_ops) == 3
        
        # Verify TOC update
        toc_ops = [op for op in result.ops if op.operation_type == "update_toc"]
        assert len(toc_ops) == 1


if __name__ == "__main__":
    pytest.main([__file__, "-v"])