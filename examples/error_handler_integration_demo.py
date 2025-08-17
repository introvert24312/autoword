"""
Integration demo showing how to use the error handling system with other vNext modules.

This script demonstrates how the error handler integrates with the executor,
validator, and other pipeline components for comprehensive error recovery.
"""

import os
import tempfile
from pathlib import Path

from autoword.vnext.error_handler import (
    PipelineErrorHandler, ErrorContext, RevisionHandlingStrategy,
    ProcessingStatus
)
from autoword.vnext.exceptions import ExecutionError, ValidationError
from autoword.vnext.models import (
    PlanV1, DeleteSectionByHeading, MatchMode, OperationResult
)


def demo_executor_integration():
    """Demonstrate error handler integration with executor."""
    print("\n=== Demo: Executor Integration ===")
    
    with tempfile.TemporaryDirectory() as temp_dir:
        audit_dir = os.path.join(temp_dir, "audit")
        handler = PipelineErrorHandler(audit_dir, RevisionHandlingStrategy.BYPASS)
        
        # Create mock document files
        original_docx = os.path.join(temp_dir, "original.docx")
        modified_docx = os.path.join(temp_dir, "modified.docx")
        
        with open(original_docx, 'w', encoding='utf-8') as f:
            f.write("Original document content")
        with open(modified_docx, 'w', encoding='utf-8') as f:
            f.write("Modified document content")
        
        # Simulate executor operations with error handling
        operations = [
            {
                "type": "delete_section_by_heading",
                "data": {"heading_text": "摘要", "level": 1, "match": "EXACT"}
            },
            {
                "type": "execute_macro",  # This will trigger security violation
                "data": {"macro_name": "AutoOpen"}
            },
            {
                "type": "update_toc",
                "data": {}
            }
        ]
        
        for i, operation in enumerate(operations, 1):
            print(f"\n{i}. Processing operation: {operation['type']}")
            
            try:
                # Validate operation security
                handler.validate_operation_security(
                    operation["type"], 
                    operation["data"]
                )
                print(f"   ✓ Security validation passed")
                
                # Simulate operation execution
                if operation["type"] == "delete_section_by_heading":
                    # Simulate NOOP operation (heading not found)
                    handler.log_noop_operation(
                        operation["type"],
                        "Heading not found in document",
                        operation["data"]
                    )
                    print(f"   → NOOP: Heading not found")
                elif operation["type"] == "update_toc":
                    print(f"   ✓ Operation executed successfully")
                
            except Exception as e:
                print(f"   ✗ Operation failed: {e}")
                
                # Handle the error
                context = ErrorContext(
                    pipeline_stage="executor",
                    operation_type=operation["type"],
                    docx_path=modified_docx,
                    original_docx_path=original_docx
                )
                
                result = handler.handle_pipeline_error(e, context)
                print(f"   → Recovery status: {result.status}")
                print(f"   → Rollback performed: {result.rollback_performed}")


def demo_validator_integration():
    """Demonstrate error handler integration with validator."""
    print("\n=== Demo: Validator Integration ===")
    
    with tempfile.TemporaryDirectory() as temp_dir:
        audit_dir = os.path.join(temp_dir, "audit")
        handler = PipelineErrorHandler(audit_dir, RevisionHandlingStrategy.BYPASS)
        
        # Create mock document files
        original_docx = os.path.join(temp_dir, "original.docx")
        modified_docx = os.path.join(temp_dir, "modified.docx")
        
        with open(original_docx, 'w', encoding='utf-8') as f:
            f.write("Original document content")
        with open(modified_docx, 'w', encoding='utf-8') as f:
            f.write("Modified document with validation issues")
        
        # Simulate validation scenarios
        validation_scenarios = [
            {
                "name": "Chapter Assertion Failure",
                "error": ValidationError(
                    "Found '摘要' at level 1",
                    assertion_failures=["Chapter assertion: no '摘要/参考文献' at level 1"],
                    rollback_path=original_docx
                )
            },
            {
                "name": "Style Assertion Failure", 
                "error": ValidationError(
                    "H1 style font mismatch",
                    assertion_failures=["Style assertion: H1 font should be 楷体, found 宋体"],
                    rollback_path=original_docx
                )
            },
            {
                "name": "TOC Assertion Failure",
                "error": ValidationError(
                    "TOC page numbers don't match headings",
                    assertion_failures=["TOC assertion: page number mismatch for heading 'Introduction'"],
                    rollback_path=original_docx
                )
            }
        ]
        
        for i, scenario in enumerate(validation_scenarios, 1):
            print(f"\n{i}. Testing {scenario['name']}...")
            
            context = ErrorContext(
                pipeline_stage="validator",
                docx_path=modified_docx,
                original_docx_path=original_docx
            )
            
            result = handler.handle_pipeline_error(scenario["error"], context)
            
            print(f"   Error: {str(scenario['error'])[:60]}...")
            print(f"   Status: {result.status}")
            print(f"   Rollback: {result.rollback_performed}")
            print(f"   Warnings: {len(result.warnings)} warning(s)")
            
            # Check if status file was created
            status_file = os.path.join(audit_dir, "result.status.txt")
            if os.path.exists(status_file):
                print(f"   ✓ Status file created")


def demo_pipeline_orchestration():
    """Demonstrate complete pipeline error handling orchestration."""
    print("\n=== Demo: Pipeline Orchestration ===")
    
    with tempfile.TemporaryDirectory() as temp_dir:
        audit_dir = os.path.join(temp_dir, "audit")
        handler = PipelineErrorHandler(audit_dir, RevisionHandlingStrategy.ACCEPT_ALL)
        
        # Create mock document files
        original_docx = os.path.join(temp_dir, "original.docx")
        current_docx = os.path.join(temp_dir, "current.docx")
        
        with open(original_docx, 'w', encoding='utf-8') as f:
            f.write("Original document content")
        with open(current_docx, 'w', encoding='utf-8') as f:
            f.write("Original document content")
        
        # Simulate complete pipeline execution with error handling
        pipeline_stages = [
            ("extractor", "Document structure extraction"),
            ("planner", "LLM plan generation"),
            ("executor", "Atomic operations execution"),
            ("validator", "Document validation"),
            ("auditor", "Audit trail generation")
        ]
        
        print("Simulating complete pipeline execution...")
        
        for stage_name, stage_description in pipeline_stages:
            print(f"\n→ {stage_name.upper()}: {stage_description}")
            
            try:
                # Simulate different types of errors at different stages
                if stage_name == "planner":
                    # Simulate planning error
                    raise ExecutionError(
                        "LLM returned invalid JSON",
                        operation_type="plan_generation",
                        operation_data={"user_intent": "Remove abstract section"}
                    )
                elif stage_name == "executor":
                    # Simulate successful execution with NOOP operations
                    handler.log_noop_operation(
                        "delete_section_by_heading",
                        "Heading '参考文献' not found",
                        {"heading_text": "参考文献", "level": 1}
                    )
                    print("  ✓ Operations executed (with NOOPs)")
                elif stage_name == "validator":
                    # Simulate validation success
                    print("  ✓ All assertions passed")
                elif stage_name == "auditor":
                    # Simulate audit trail creation
                    print("  ✓ Audit trail created")
                else:
                    print("  ✓ Stage completed successfully")
                    
            except Exception as e:
                print(f"  ✗ Stage failed: {str(e)[:50]}...")
                
                context = ErrorContext(
                    pipeline_stage=stage_name,
                    docx_path=current_docx,
                    original_docx_path=original_docx
                )
                
                result = handler.handle_pipeline_error(e, context)
                
                print(f"  → Recovery: {result.status}")
                print(f"  → Rollback: {result.rollback_performed}")
                
                # In real pipeline, we would stop here on critical errors
                if result.status in [ProcessingStatus.SECURITY_VIOLATION, ProcessingStatus.ROLLBACK]:
                    print("  → Pipeline terminated due to critical error")
                    break
        
        # Show final audit summary
        print(f"\nFinal audit summary:")
        warnings_log = os.path.join(audit_dir, "warnings.log")
        if os.path.exists(warnings_log):
            with open(warnings_log, 'r', encoding='utf-8') as f:
                lines = f.readlines()
                print(f"  Warnings logged: {len([l for l in lines if not l.startswith('#')])} entries")
        
        status_file = os.path.join(audit_dir, "result.status.txt")
        if os.path.exists(status_file):
            with open(status_file, 'r', encoding='utf-8') as f:
                status_content = f.read()
                status_line = [l for l in status_content.split('\n') if l.startswith('Status:')][0]
                print(f"  Final status: {status_line}")


def main():
    """Run all integration demos."""
    print("AutoWord vNext Error Handler Integration Demo")
    print("=" * 60)
    
    try:
        demo_executor_integration()
        demo_validator_integration()
        demo_pipeline_orchestration()
        
        print("\n" + "=" * 60)
        print("Integration demo completed successfully!")
        print("\nKey integration features demonstrated:")
        print("✓ Security validation in executor pipeline")
        print("✓ NOOP operation logging during execution")
        print("✓ Automatic rollback on validation failures")
        print("✓ Complete pipeline error orchestration")
        print("✓ Audit trail generation across all stages")
        print("✓ Status tracking and recovery reporting")
        
    except Exception as e:
        print(f"\nIntegration demo failed with error: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()