"""
Tests for runtime constraint enforcement system.

This module tests all aspects of the RuntimeConstraintEnforcer:
- Whitelist operation validation
- String replacement prevention
- Word object layer enforcement
- LLM output validation with JSON schema gateway
- User input sanitization and parameter validation
"""

import json
import pytest
from unittest.mock import Mock, patch

from autoword.vnext.constraints import (
    RuntimeConstraintEnforcer, validate_plan_constraints,
    validate_llm_output_constraints, sanitize_user_input_safe
)
from autoword.vnext.models import (
    PlanV1, DeleteSectionByHeading, UpdateToc, SetStyleRule,
    ReassignParagraphsToStyle, ClearDirectFormatting,
    FontSpec, ParagraphSpec, MatchMode, LineSpacingMode
)
from autoword.vnext.exceptions import SecurityViolationError


class TestRuntimeConstraintEnforcer:
    """Test runtime constraint enforcement."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.enforcer = RuntimeConstraintEnforcer()
    
    def test_init(self):
        """Test enforcer initialization."""
        assert self.enforcer is not None
        assert len(self.enforcer.WHITELISTED_OPERATIONS) == 6
        assert "delete_section_by_heading" in self.enforcer.WHITELISTED_OPERATIONS
        assert len(self.enforcer._violation_log) == 0
    
    def test_validate_whitelist_operations_valid(self):
        """Test whitelist validation with valid operations."""
        plan = PlanV1(ops=[
            DeleteSectionByHeading(heading_text="Abstract", level=1),
            UpdateToc(),
            SetStyleRule(target_style_name="Heading 1")
        ])
        
        result = self.enforcer._validate_whitelist_operations(plan)
        
        assert result.is_valid
        assert len(result.errors) == 0
    
    def test_validate_whitelist_operations_invalid(self):
        """Test whitelist validation with invalid operations."""
        # Create a plan with invalid operation type
        plan_data = {
            "schema_version": "plan.v1",
            "ops": [
                {
                    "operation_type": "invalid_operation",
                    "some_param": "value"
                }
            ]
        }
        
        # Mock the plan to bypass Pydantic validation for testing
        mock_plan = Mock()
        mock_plan.ops = [Mock()]
        mock_plan.ops[0].operation_type = "invalid_operation"
        
        result = self.enforcer._validate_whitelist_operations(mock_plan)
        
        assert not result.is_valid
        assert len(result.errors) == 1
        assert "invalid_operation" in result.errors[0]
        assert "not in whitelist" in result.errors[0]
    
    def test_validate_no_string_replacement_valid(self):
        """Test string replacement validation with valid plan."""
        plan = PlanV1(ops=[
            DeleteSectionByHeading(heading_text="Abstract", level=1),
            SetStyleRule(target_style_name="Heading 1")
        ])
        
        result = self.enforcer._validate_no_string_replacement(plan)
        
        assert result.is_valid
        assert len(result.errors) == 0
    
    def test_validate_no_string_replacement_invalid(self):
        """Test string replacement validation with forbidden patterns."""
        # Create plan data with forbidden patterns
        plan_data = {
            "schema_version": "plan.v1",
            "ops": [
                {
                    "operation_type": "set_style_rule",
                    "target_style_name": "text.replace('old', 'new')",  # Forbidden pattern
                    "font": {"east_asian": "宋体"}
                }
            ]
        }
        
        mock_plan = Mock()
        mock_plan.model_dump.return_value = plan_data
        
        result = self.enforcer._validate_no_string_replacement(mock_plan)
        
        assert not result.is_valid
        assert len(result.errors) >= 1
        assert any("string replacement pattern" in error for error in result.errors)
    
    def test_validate_word_object_layer_valid(self):
        """Test Word object layer validation with valid operations."""
        plan = PlanV1(ops=[
            SetStyleRule(
                target_style_name="Heading 1",
                font=FontSpec(east_asian="宋体", size_pt=12)
            )
        ])
        
        result = self.enforcer._validate_word_object_layer_only(plan)
        
        assert result.is_valid
        assert len(result.errors) == 0
    
    def test_validate_word_object_layer_invalid(self):
        """Test Word object layer validation with forbidden operations."""
        plan_data = {
            "schema_version": "plan.v1",
            "ops": [
                {
                    "operation_type": "set_style_rule",
                    "target_style_name": "Heading 1",
                    "Text": "Direct text manipulation"  # Forbidden
                }
            ]
        }
        
        mock_plan = Mock()
        mock_plan.model_dump.return_value = plan_data
        
        result = self.enforcer._validate_word_object_layer_only(mock_plan)
        
        assert not result.is_valid
        assert len(result.errors) >= 1
        assert any("Forbidden Word COM operation" in error for error in result.errors)
    
    def test_validate_authorization_requirements_valid(self):
        """Test authorization validation with properly authorized operations."""
        plan = PlanV1(ops=[
            ClearDirectFormatting(
                scope="document",
                authorization_required=True
            )
        ])
        
        result = self.enforcer._validate_authorization_requirements(plan)
        
        assert result.is_valid
        assert len(result.errors) == 0
    
    def test_validate_authorization_requirements_invalid(self):
        """Test authorization validation with missing authorization."""
        # Create mock operation without authorization
        mock_plan = Mock()
        mock_op = Mock()
        mock_op.operation_type = "clear_direct_formatting"
        mock_op.authorization_required = False
        mock_plan.ops = [mock_op]
        
        result = self.enforcer._validate_authorization_requirements(mock_plan)
        
        assert not result.is_valid
        assert len(result.errors) == 1
        assert "requires explicit authorization" in result.errors[0]
    
    def test_validate_and_sanitize_parameters_valid(self):
        """Test parameter validation with valid parameters."""
        plan = PlanV1(ops=[
            DeleteSectionByHeading(
                heading_text="Abstract",
                level=1,
                match=MatchMode.EXACT
            ),
            SetStyleRule(
                target_style_name="Heading 1",
                font=FontSpec(color_hex="#000000")
            )
        ])
        
        result = self.enforcer._validate_and_sanitize_parameters(plan)
        
        assert result.is_valid
        assert len(result.errors) == 0
    
    def test_validate_and_sanitize_parameters_invalid(self):
        """Test parameter validation with invalid parameters."""
        # Create operations with invalid parameters
        mock_plan = Mock()
        
        # Mock operation with empty heading text
        mock_op1 = Mock()
        mock_op1.operation_type = "delete_section_by_heading"
        mock_op1.heading_text = "   "  # Empty after strip
        mock_op1.match = Mock()
        mock_op1.match.value = "EXACT"
        
        # Mock operation with invalid color
        mock_op2 = Mock()
        mock_op2.operation_type = "set_style_rule"
        mock_op2.target_style_name = "Heading 1"
        mock_op2.font = Mock()
        mock_op2.font.color_hex = "invalid_color"
        
        mock_plan.ops = [mock_op1, mock_op2]
        
        result = self.enforcer._validate_and_sanitize_parameters(mock_plan)
        
        assert not result.is_valid
        assert len(result.errors) >= 2
        assert any("heading_text cannot be empty" in error for error in result.errors)
        assert any("invalid color_hex format" in error for error in result.errors)
    
    def test_validate_plan_constraints_comprehensive(self):
        """Test comprehensive plan constraint validation."""
        plan = PlanV1(ops=[
            DeleteSectionByHeading(heading_text="Abstract", level=1),
            UpdateToc(),
            SetStyleRule(
                target_style_name="Heading 1",
                font=FontSpec(east_asian="宋体", size_pt=12)
            )
        ])
        
        result = self.enforcer.validate_plan_constraints(plan)
        
        assert result.is_valid
        assert len(result.errors) == 0
    
    def test_validate_llm_output_valid_json(self):
        """Test LLM output validation with valid JSON."""
        llm_output = {
            "schema_version": "plan.v1",
            "ops": [
                {
                    "operation_type": "delete_section_by_heading",
                    "heading_text": "Abstract",
                    "level": 1
                }
            ]
        }
        
        result = self.enforcer.validate_llm_output(llm_output)
        
        assert result.is_valid
        assert len(result.errors) == 0
    
    def test_validate_llm_output_invalid_json(self):
        """Test LLM output validation with invalid JSON."""
        llm_output = '{"invalid": json,}'  # Invalid JSON
        
        result = self.enforcer.validate_llm_output(llm_output)
        
        assert not result.is_valid
        assert len(result.errors) >= 1
        assert any("Invalid JSON" in error for error in result.errors)
    
    def test_validate_llm_output_suspicious_content(self):
        """Test LLM output validation with suspicious content."""
        llm_output = {
            "schema_version": "plan.v1",
            "ops": [],
            "suspicious_content": "<w:document>OOXML content</w:document>"
        }
        
        result = self.enforcer.validate_llm_output(llm_output)
        
        # Should detect suspicious content
        assert not result.is_valid or len(result.warnings) > 0
    
    def test_sanitize_user_input_valid(self):
        """Test user input sanitization with valid input."""
        user_input = {
            "heading_text": "Abstract",
            "level": 1,
            "font_name": "宋体"
        }
        
        sanitized, warnings = self.enforcer.sanitize_user_input(user_input)
        
        assert sanitized == user_input
        assert len(warnings) == 0
    
    def test_sanitize_user_input_with_sanitization(self):
        """Test user input sanitization with content that needs sanitization."""
        user_input = {
            "heading_text": "Abstract\x00\x01",  # Contains null and control chars
            "long_text": "a" * 1500,  # Too long
            "nested": {
                "value": "test\x00"
            }
        }
        
        sanitized, warnings = self.enforcer.sanitize_user_input(user_input)
        
        assert sanitized["heading_text"] == "Abstract"
        assert len(sanitized["long_text"]) == 1000
        assert sanitized["nested"]["value"] == "test"
        assert len(warnings) >= 2  # Should have warnings for sanitization
    
    def test_sanitize_user_input_security_violation(self):
        """Test user input sanitization with security violations."""
        user_input = {
            "malicious_field": "<script>alert('xss')</script>"
        }
        
        with pytest.raises(SecurityViolationError):
            self.enforcer.sanitize_user_input(user_input)
    
    def test_validate_operation_execution_valid(self):
        """Test operation execution validation with valid operation."""
        operation = DeleteSectionByHeading(
            heading_text="Abstract",
            level=1,
            match=MatchMode.EXACT
        )
        
        result = self.enforcer.validate_operation_execution(operation)
        
        assert result.is_valid
        assert len(result.errors) == 0
    
    def test_validate_operation_execution_invalid_type(self):
        """Test operation execution validation with invalid operation type."""
        mock_operation = Mock()
        mock_operation.operation_type = "invalid_operation"
        
        result = self.enforcer.validate_operation_execution(mock_operation)
        
        assert not result.is_valid
        assert len(result.errors) >= 1
        assert any("not in whitelist" in error for error in result.errors)
    
    def test_validate_delete_section_execution_complex_regex(self):
        """Test delete section execution validation with complex regex."""
        operation = DeleteSectionByHeading(
            heading_text=".*+",  # Dangerous regex pattern
            level=1,
            match=MatchMode.REGEX
        )
        
        result = self.enforcer._validate_delete_section_execution(operation)
        
        assert result.is_valid  # Should be valid but with warnings
        assert len(result.warnings) >= 1
        assert any("Complex regex pattern" in warning for warning in result.warnings)
    
    def test_validate_set_style_execution_suspicious_font(self):
        """Test set style execution validation with suspicious font name."""
        operation = SetStyleRule(
            target_style_name="Heading 1",
            font=FontSpec(east_asian="<script>alert('xss')</script>")
        )
        
        result = self.enforcer._validate_set_style_execution(operation)
        
        assert result.is_valid  # Should be valid but with warnings
        assert len(result.warnings) >= 1
        assert any("suspicious characters" in warning for warning in result.warnings)
    
    def test_validate_clear_formatting_execution_no_auth(self):
        """Test clear formatting execution validation without authorization."""
        mock_operation = Mock()
        mock_operation.operation_type = "clear_direct_formatting"
        mock_operation.authorization_required = False
        mock_operation.range_spec = None
        
        result = self.enforcer._validate_clear_formatting_execution(mock_operation)
        
        assert not result.is_valid
        assert len(result.errors) >= 1
        assert any("requires explicit authorization" in error for error in result.errors)
    
    def test_violation_log(self):
        """Test security violation logging."""
        # Trigger a violation
        mock_plan = Mock()
        mock_plan.ops = [Mock()]
        mock_plan.ops[0].operation_type = "invalid_operation"
        
        self.enforcer._validate_whitelist_operations(mock_plan)
        
        # Check violation log
        violations = self.enforcer.get_violation_log()
        assert len(violations) >= 1
        
        # Clear log
        self.enforcer.clear_violation_log()
        assert len(self.enforcer.get_violation_log()) == 0
    
    def test_detect_suspicious_llm_content_clean(self):
        """Test suspicious content detection with clean content."""
        clean_output = {
            "schema_version": "plan.v1",
            "ops": [
                {
                    "operation_type": "delete_section_by_heading",
                    "heading_text": "Abstract",
                    "level": 1
                }
            ]
        }
        
        result = self.enforcer._detect_suspicious_llm_content(clean_output)
        
        assert result.is_valid
        assert len(result.errors) == 0
    
    def test_detect_suspicious_llm_content_suspicious(self):
        """Test suspicious content detection with suspicious content."""
        suspicious_output = {
            "schema_version": "plan.v1",
            "ops": [],
            "embedded_docx": "<w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>"
        }
        
        result = self.enforcer._detect_suspicious_llm_content(suspicious_output)
        
        assert not result.is_valid
        assert len(result.errors) >= 1
        assert any("Suspicious content detected" in error for error in result.errors)


class TestConvenienceFunctions:
    """Test convenience functions for constraint validation."""
    
    def test_validate_plan_constraints_function(self):
        """Test validate_plan_constraints convenience function."""
        plan = PlanV1(ops=[
            DeleteSectionByHeading(heading_text="Abstract", level=1)
        ])
        
        result = validate_plan_constraints(plan)
        
        assert result.is_valid
        assert len(result.errors) == 0
    
    def test_validate_llm_output_constraints_function(self):
        """Test validate_llm_output_constraints convenience function."""
        llm_output = {
            "schema_version": "plan.v1",
            "ops": [
                {
                    "operation_type": "update_toc"
                }
            ]
        }
        
        result = validate_llm_output_constraints(llm_output)
        
        assert result.is_valid
        assert len(result.errors) == 0
    
    def test_sanitize_user_input_safe_function(self):
        """Test sanitize_user_input_safe convenience function."""
        user_input = {
            "heading_text": "Abstract",
            "level": 1
        }
        
        sanitized, warnings = sanitize_user_input_safe(user_input)
        
        assert sanitized == user_input
        assert len(warnings) == 0
    
    def test_sanitize_user_input_safe_function_violation(self):
        """Test sanitize_user_input_safe convenience function with violation."""
        user_input = {
            "malicious": "<script>alert('xss')</script>"
        }
        
        with pytest.raises(SecurityViolationError):
            sanitize_user_input_safe(user_input)


class TestConstraintIntegration:
    """Test constraint enforcement integration scenarios."""
    
    def test_full_plan_validation_pipeline(self):
        """Test complete plan validation pipeline."""
        # Create a comprehensive plan
        plan = PlanV1(ops=[
            DeleteSectionByHeading(
                heading_text="摘要",
                level=1,
                match=MatchMode.EXACT,
                case_sensitive=False
            ),
            DeleteSectionByHeading(
                heading_text="参考文献",
                level=1,
                match=MatchMode.EXACT,
                case_sensitive=False
            ),
            UpdateToc(),
            SetStyleRule(
                target_style_name="标题 1",
                font=FontSpec(
                    east_asian="楷体",
                    size_pt=12,
                    bold=True
                ),
                paragraph=ParagraphSpec(
                    line_spacing_mode=LineSpacingMode.MULTIPLE,
                    line_spacing_value=2.0
                )
            ),
            SetStyleRule(
                target_style_name="标题 2",
                font=FontSpec(
                    east_asian="宋体",
                    size_pt=12,
                    bold=True
                ),
                paragraph=ParagraphSpec(
                    line_spacing_mode=LineSpacingMode.MULTIPLE,
                    line_spacing_value=2.0
                )
            ),
            SetStyleRule(
                target_style_name="正文",
                font=FontSpec(
                    east_asian="宋体",
                    size_pt=12,
                    bold=False
                ),
                paragraph=ParagraphSpec(
                    line_spacing_mode=LineSpacingMode.MULTIPLE,
                    line_spacing_value=2.0
                )
            )
        ])
        
        enforcer = RuntimeConstraintEnforcer()
        result = enforcer.validate_plan_constraints(plan)
        
        assert result.is_valid
        assert len(result.errors) == 0
        
        # Validate each operation individually
        for operation in plan.ops:
            op_result = enforcer.validate_operation_execution(operation)
            assert op_result.is_valid, f"Operation {operation.operation_type} failed validation"
    
    def test_llm_output_to_plan_validation(self):
        """Test LLM output validation and conversion to plan."""
        llm_json_output = '''
        {
            "schema_version": "plan.v1",
            "ops": [
                {
                    "operation_type": "delete_section_by_heading",
                    "heading_text": "摘要",
                    "level": 1,
                    "match": "EXACT",
                    "case_sensitive": false
                },
                {
                    "operation_type": "update_toc"
                }
            ]
        }
        '''
        
        enforcer = RuntimeConstraintEnforcer()
        
        # Validate LLM output
        llm_result = enforcer.validate_llm_output(llm_json_output)
        assert llm_result.is_valid
        
        # Convert to plan and validate constraints
        plan_data = json.loads(llm_json_output)
        plan = PlanV1(**plan_data)
        
        plan_result = enforcer.validate_plan_constraints(plan)
        assert plan_result.is_valid
    
    def test_user_input_sanitization_pipeline(self):
        """Test user input sanitization pipeline."""
        raw_user_input = {
            "user_intent": "删除摘要和参考文献部分，更新目录",
            "document_path": "/path/to/document.docx",
            "style_preferences": {
                "heading_font": "楷体",
                "body_font": "宋体",
                "font_size": 12
            },
            "special_instructions": "请保持格式一致性\x00"  # Contains null byte
        }
        
        enforcer = RuntimeConstraintEnforcer()
        sanitized, warnings = enforcer.sanitize_user_input(raw_user_input)
        
        # Should sanitize null byte
        assert "\x00" not in sanitized["special_instructions"]
        assert len(warnings) >= 1
        
        # Other fields should remain intact
        assert sanitized["user_intent"] == raw_user_input["user_intent"]
        assert sanitized["style_preferences"]["heading_font"] == "楷体"


if __name__ == "__main__":
    pytest.main([__file__])