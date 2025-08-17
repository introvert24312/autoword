"""
Runtime Constraint Enforcement System Demo

This demo shows how the RuntimeConstraintEnforcer validates plans, LLM output,
and user input to ensure security and safety in the AutoWord vNext pipeline.
"""

import json
import logging
from pathlib import Path

# Add the parent directory to the path so we can import autoword modules
import sys
sys.path.insert(0, str(Path(__file__).parent.parent))

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


# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


def demo_whitelist_validation():
    """Demonstrate whitelist operation validation."""
    print("\n" + "="*60)
    print("DEMO: Whitelist Operation Validation")
    print("="*60)
    
    enforcer = RuntimeConstraintEnforcer()
    
    # Valid plan with whitelisted operations
    print("\n1. Valid Plan with Whitelisted Operations:")
    valid_plan = PlanV1(ops=[
        DeleteSectionByHeading(heading_text="摘要", level=1),
        DeleteSectionByHeading(heading_text="参考文献", level=1),
        UpdateToc(),
        SetStyleRule(
            target_style_name="标题 1",
            font=FontSpec(east_asian="楷体", size_pt=12, bold=True)
        )
    ])
    
    result = enforcer.validate_plan_constraints(valid_plan)
    print(f"✓ Valid plan: {result.is_valid}")
    print(f"  Errors: {len(result.errors)}")
    print(f"  Warnings: {len(result.warnings)}")
    
    # Invalid plan with non-whitelisted operation (simulated)
    print("\n2. Invalid Plan with Non-Whitelisted Operation:")
    try:
        # This would normally be caught by Pydantic validation,
        # but we'll simulate it for demo purposes
        from unittest.mock import Mock
        
        mock_plan = Mock()
        mock_op = Mock()
        mock_op.operation_type = "execute_macro"  # Not whitelisted
        mock_plan.ops = [mock_op]
        
        result = enforcer._validate_whitelist_operations(mock_plan)
        print(f"✗ Invalid plan: {result.is_valid}")
        print(f"  Errors: {result.errors}")
        
    except Exception as e:
        print(f"✗ Validation failed as expected: {str(e)}")


def demo_string_replacement_prevention():
    """Demonstrate string replacement prevention."""
    print("\n" + "="*60)
    print("DEMO: String Replacement Prevention")
    print("="*60)
    
    enforcer = RuntimeConstraintEnforcer()
    
    # Valid plan without string replacement
    print("\n1. Valid Plan (No String Replacement):")
    valid_plan = PlanV1(ops=[
        SetStyleRule(
            target_style_name="Heading 1",
            font=FontSpec(east_asian="宋体", size_pt=12)
        )
    ])
    
    result = enforcer._validate_no_string_replacement(valid_plan)
    print(f"✓ No string replacement detected: {result.is_valid}")
    
    # Simulate plan with string replacement patterns
    print("\n2. Plan with Forbidden String Replacement Patterns:")
    from unittest.mock import Mock
    
    mock_plan = Mock()
    mock_plan.model_dump.return_value = {
        "schema_version": "plan.v1",
        "ops": [
            {
                "operation_type": "set_style_rule",
                "target_style_name": "text.replace('old', 'new')",  # Forbidden
                "malicious_code": "doc.Range.Text = 'hacked'"  # Forbidden
            }
        ]
    }
    
    result = enforcer._validate_no_string_replacement(mock_plan)
    print(f"✗ String replacement detected: {result.is_valid}")
    print(f"  Errors: {result.errors}")


def demo_llm_output_validation():
    """Demonstrate LLM output validation with JSON schema gateway."""
    print("\n" + "="*60)
    print("DEMO: LLM Output Validation")
    print("="*60)
    
    enforcer = RuntimeConstraintEnforcer()
    
    # Valid LLM output
    print("\n1. Valid LLM JSON Output:")
    valid_llm_output = {
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
    
    result = enforcer.validate_llm_output(valid_llm_output)
    print(f"✓ Valid LLM output: {result.is_valid}")
    print(f"  Errors: {len(result.errors)}")
    print(f"  Warnings: {len(result.warnings)}")
    
    # Invalid JSON from LLM
    print("\n2. Invalid JSON from LLM:")
    invalid_json = '{"schema_version": "plan.v1", "ops": [invalid json,]}'
    
    result = enforcer.validate_llm_output(invalid_json)
    print(f"✗ Invalid JSON: {result.is_valid}")
    print(f"  Errors: {result.errors}")
    
    # LLM output with suspicious content
    print("\n3. LLM Output with Suspicious Content:")
    suspicious_output = {
        "schema_version": "plan.v1",
        "ops": [],
        "embedded_docx": "<w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>",
        "ooxml_content": "<?xml version='1.0'?><document>"
    }
    
    result = enforcer.validate_llm_output(suspicious_output)
    print(f"✗ Suspicious content detected: {result.is_valid}")
    if result.errors:
        print(f"  Errors: {result.errors}")
    if result.warnings:
        print(f"  Warnings: {result.warnings}")


def demo_user_input_sanitization():
    """Demonstrate user input sanitization and validation."""
    print("\n" + "="*60)
    print("DEMO: User Input Sanitization")
    print("="*60)
    
    enforcer = RuntimeConstraintEnforcer()
    
    # Clean user input
    print("\n1. Clean User Input:")
    clean_input = {
        "user_intent": "删除摘要和参考文献，更新目录",
        "document_path": "/path/to/document.docx",
        "style_preferences": {
            "heading_font": "楷体",
            "body_font": "宋体"
        }
    }
    
    sanitized, warnings = enforcer.sanitize_user_input(clean_input)
    print(f"✓ Clean input processed successfully")
    print(f"  Warnings: {len(warnings)}")
    print(f"  Input unchanged: {sanitized == clean_input}")
    
    # Input requiring sanitization
    print("\n2. Input Requiring Sanitization:")
    dirty_input = {
        "heading_text": "Abstract\x00\x01\x02",  # Contains control characters
        "long_description": "a" * 1500,  # Too long
        "nested_data": {
            "value": "test\x00data"
        }
    }
    
    sanitized, warnings = enforcer.sanitize_user_input(dirty_input)
    print(f"✓ Input sanitized successfully")
    print(f"  Warnings: {warnings}")
    print(f"  Original heading_text length: {len(dirty_input['heading_text'])}")
    print(f"  Sanitized heading_text: '{sanitized['heading_text']}'")
    print(f"  Original description length: {len(dirty_input['long_description'])}")
    print(f"  Sanitized description length: {len(sanitized['long_description'])}")
    
    # Malicious input
    print("\n3. Malicious Input (Security Violation):")
    malicious_input = {
        "user_intent": "删除文档",
        "malicious_script": "<script>alert('XSS attack')</script>",
        "injection_attempt": "javascript:void(0)"
    }
    
    try:
        sanitized, warnings = enforcer.sanitize_user_input(malicious_input)
        print("✗ Malicious input should have been rejected!")
    except SecurityViolationError as e:
        print(f"✓ Security violation detected and blocked: {str(e)}")


def demo_operation_execution_validation():
    """Demonstrate operation execution validation."""
    print("\n" + "="*60)
    print("DEMO: Operation Execution Validation")
    print("="*60)
    
    enforcer = RuntimeConstraintEnforcer()
    
    # Valid operations
    print("\n1. Valid Operations:")
    operations = [
        DeleteSectionByHeading(heading_text="摘要", level=1),
        SetStyleRule(
            target_style_name="标题 1",
            font=FontSpec(east_asian="楷体", size_pt=12)
        ),
        ClearDirectFormatting(
            scope="document",
            authorization_required=True
        )
    ]
    
    for i, operation in enumerate(operations, 1):
        result = enforcer.validate_operation_execution(operation)
        print(f"  Operation {i} ({operation.operation_type}): {'✓' if result.is_valid else '✗'}")
        if result.warnings:
            print(f"    Warnings: {result.warnings}")
    
    # Operation with complex regex
    print("\n2. Operation with Complex Regex:")
    complex_regex_op = DeleteSectionByHeading(
        heading_text=".*+",  # Potentially dangerous regex
        level=1,
        match=MatchMode.REGEX
    )
    
    result = enforcer.validate_operation_execution(complex_regex_op)
    print(f"  Complex regex operation: {'✓' if result.is_valid else '✗'}")
    print(f"  Warnings: {result.warnings}")
    
    # Operation with suspicious content
    print("\n3. Operation with Suspicious Content:")
    suspicious_op = SetStyleRule(
        target_style_name="Heading 1",
        font=FontSpec(east_asian="<script>alert('xss')</script>")
    )
    
    result = enforcer.validate_operation_execution(suspicious_op)
    print(f"  Suspicious content operation: {'✓' if result.is_valid else '✗'}")
    print(f"  Warnings: {result.warnings}")


def demo_comprehensive_validation():
    """Demonstrate comprehensive validation pipeline."""
    print("\n" + "="*60)
    print("DEMO: Comprehensive Validation Pipeline")
    print("="*60)
    
    # Simulate a complete workflow
    print("\n1. User Input → Sanitization → LLM → Plan Validation")
    
    # Step 1: User input
    user_input = {
        "user_intent": "删除摘要和参考文献部分，更新目录，设置标题样式",
        "document_path": "/path/to/paper.docx",
        "style_requirements": {
            "h1_font": "楷体",
            "h1_size": 12,
            "h2_font": "宋体",
            "h2_size": 12
        }
    }
    
    print("  Step 1: User Input Sanitization")
    try:
        sanitized_input, warnings = sanitize_user_input_safe(user_input)
        print(f"    ✓ Input sanitized, warnings: {len(warnings)}")
    except SecurityViolationError as e:
        print(f"    ✗ Security violation: {str(e)}")
        return
    
    # Step 2: Simulate LLM output
    print("  Step 2: LLM Output Validation")
    llm_output = {
        "schema_version": "plan.v1",
        "ops": [
            {
                "operation_type": "delete_section_by_heading",
                "heading_text": "摘要",
                "level": 1,
                "match": "EXACT"
            },
            {
                "operation_type": "delete_section_by_heading", 
                "heading_text": "参考文献",
                "level": 1,
                "match": "EXACT"
            },
            {
                "operation_type": "update_toc"
            },
            {
                "operation_type": "set_style_rule",
                "target_style_name": "标题 1",
                "font": {
                    "east_asian": "楷体",
                    "size_pt": 12,
                    "bold": True
                }
            },
            {
                "operation_type": "set_style_rule",
                "target_style_name": "标题 2", 
                "font": {
                    "east_asian": "宋体",
                    "size_pt": 12,
                    "bold": True
                }
            }
        ]
    }
    
    llm_result = validate_llm_output_constraints(llm_output)
    print(f"    {'✓' if llm_result.is_valid else '✗'} LLM output validation: {llm_result.is_valid}")
    if llm_result.errors:
        print(f"      Errors: {llm_result.errors}")
    
    # Step 3: Plan constraint validation
    if llm_result.is_valid:
        print("  Step 3: Plan Constraint Validation")
        plan = PlanV1(**llm_output)
        plan_result = validate_plan_constraints(plan)
        print(f"    {'✓' if plan_result.is_valid else '✗'} Plan constraints: {plan_result.is_valid}")
        if plan_result.errors:
            print(f"      Errors: {plan_result.errors}")
        
        # Step 4: Individual operation validation
        print("  Step 4: Individual Operation Validation")
        enforcer = RuntimeConstraintEnforcer()
        all_ops_valid = True
        for i, operation in enumerate(plan.ops):
            op_result = enforcer.validate_operation_execution(operation)
            status = "✓" if op_result.is_valid else "✗"
            print(f"    {status} Operation {i+1} ({operation.operation_type}): {op_result.is_valid}")
            if not op_result.is_valid:
                all_ops_valid = False
        
        print(f"\n  Final Result: {'✓ All validations passed' if all_ops_valid and plan_result.is_valid else '✗ Validation failed'}")


def demo_security_violation_logging():
    """Demonstrate security violation logging."""
    print("\n" + "="*60)
    print("DEMO: Security Violation Logging")
    print("="*60)
    
    enforcer = RuntimeConstraintEnforcer()
    
    print("\n1. Triggering Security Violations:")
    
    # Trigger whitelist violation
    from unittest.mock import Mock
    mock_plan = Mock()
    mock_op = Mock()
    mock_op.operation_type = "malicious_operation"
    mock_plan.ops = [mock_op]
    
    enforcer._validate_whitelist_operations(mock_plan)
    
    # Trigger LLM output violation
    malicious_llm_output = {
        "schema_version": "plan.v1",
        "ops": [],
        "embedded_docx": "<w:document>malicious content</w:document>"
    }
    
    enforcer.validate_llm_output(malicious_llm_output)
    
    print("\n2. Security Violation Log:")
    violations = enforcer.get_violation_log()
    for i, violation in enumerate(violations, 1):
        print(f"  Violation {i}:")
        print(f"    Type: {violation['violation_type']}")
        print(f"    Context: {violation['context']}")
    
    print(f"\n  Total violations logged: {len(violations)}")
    
    print("\n3. Clearing Violation Log:")
    enforcer.clear_violation_log()
    print(f"  Violations after clear: {len(enforcer.get_violation_log())}")


def main():
    """Run all constraint enforcement demos."""
    print("AutoWord vNext - Runtime Constraint Enforcement System Demo")
    print("=" * 80)
    
    try:
        demo_whitelist_validation()
        demo_string_replacement_prevention()
        demo_llm_output_validation()
        demo_user_input_sanitization()
        demo_operation_execution_validation()
        demo_comprehensive_validation()
        demo_security_violation_logging()
        
        print("\n" + "="*80)
        print("✓ All constraint enforcement demos completed successfully!")
        print("="*80)
        
    except Exception as e:
        print(f"\n✗ Demo failed with error: {str(e)}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()