"""
Demo script for AutoWord vNext error handling and recovery system.

This script demonstrates comprehensive error handling, automatic rollback,
NOOP operation logging, security violation detection, and revision handling.
"""

import os
import tempfile
import shutil
from pathlib import Path

from autoword.vnext.error_handler import (
    PipelineErrorHandler, ErrorContext, RevisionHandlingStrategy,
    ProcessingStatus
)
from autoword.vnext.exceptions import (
    SecurityViolationError, ValidationError, ExecutionError
)


def create_demo_files(temp_dir: str) -> tuple[str, str]:
    """Create demo DOCX files for testing."""
    original_docx = os.path.join(temp_dir, "original_document.docx")
    modified_docx = os.path.join(temp_dir, "modified_document.docx")
    
    # Create mock DOCX files (in real scenario these would be actual Word documents)
    with open(original_docx, 'w', encoding='utf-8') as f:
        f.write("Original document content with proper formatting")
    
    with open(modified_docx, 'w', encoding='utf-8') as f:
        f.write("Modified document content with changes applied")
    
    return original_docx, modified_docx


def demo_warnings_logging():
    """Demonstrate warnings logging functionality."""
    print("\n=== Demo: Warnings Logging ===")
    
    with tempfile.TemporaryDirectory() as temp_dir:
        audit_dir = os.path.join(temp_dir, "audit")
        handler = PipelineErrorHandler(audit_dir, RevisionHandlingStrategy.BYPASS)
        
        # Demo NOOP operation logging
        print("1. Logging NOOP operation...")
        handler.log_noop_operation(
            "delete_section_by_heading",
            "Heading '摘要' not found in document",
            {"heading_text": "摘要", "level": 1, "match": "EXACT"}
        )
        
        # Demo font fallback logging
        print("2. Logging font fallback...")
        handler.log_font_fallback(
            "楷体", "楷体_GB2312", ["楷体", "楷体_GB2312", "STKaiti"]
        )
        
        # Demo localization fallback logging
        print("3. Logging localization fallback...")
        handler.log_localization_fallback("Heading 1", "标题 1")
        
        # Show warnings.log content
        warnings_log = os.path.join(audit_dir, "warnings.log")
        print(f"\nWarnings log content ({warnings_log}):")
        with open(warnings_log, 'r', encoding='utf-8') as f:
            print(f.read())


def demo_security_validation():
    """Demonstrate security validation functionality."""
    print("\n=== Demo: Security Validation ===")
    
    with tempfile.TemporaryDirectory() as temp_dir:
        audit_dir = os.path.join(temp_dir, "audit")
        handler = PipelineErrorHandler(audit_dir, RevisionHandlingStrategy.BYPASS)
        
        # Demo valid operation
        print("1. Validating authorized operation...")
        try:
            result = handler.validate_operation_security(
                "delete_section_by_heading",
                {"heading_text": "摘要", "level": 1, "match": "EXACT"}
            )
            print(f"   ✓ Operation authorized: {result}")
        except SecurityViolationError as e:
            print(f"   ✗ Security violation: {e}")
        
        # Demo invalid operation type
        print("2. Validating unauthorized operation type...")
        try:
            result = handler.validate_operation_security(
                "execute_macro",  # Not in whitelist
                {"macro_name": "AutoOpen"}
            )
            print(f"   ✓ Operation authorized: {result}")
        except SecurityViolationError as e:
            print(f"   ✗ Security violation: {e}")
        
        # Demo forbidden pattern
        print("3. Validating operation with forbidden pattern...")
        try:
            result = handler.validate_operation_security(
                "set_style_rule",
                {
                    "target_style": "Normal",
                    "action": "string_replacement",  # Forbidden pattern
                    "content": "Replace this text"
                }
            )
            print(f"   ✓ Operation authorized: {result}")
        except SecurityViolationError as e:
            print(f"   ✗ Security violation: {e}")


def demo_rollback_functionality():
    """Demonstrate automatic rollback functionality."""
    print("\n=== Demo: Rollback Functionality ===")
    
    with tempfile.TemporaryDirectory() as temp_dir:
        audit_dir = os.path.join(temp_dir, "audit")
        handler = PipelineErrorHandler(audit_dir, RevisionHandlingStrategy.BYPASS)
        
        # Create demo files
        original_docx, modified_docx = create_demo_files(temp_dir)
        
        print(f"Original file: {original_docx}")
        print(f"Modified file: {modified_docx}")
        
        # Show initial content
        with open(modified_docx, 'r', encoding='utf-8') as f:
            initial_content = f.read()
            print(f"Initial modified content: {initial_content[:50]}...")
        
        # Simulate validation error that triggers rollback
        print("\n1. Simulating validation error...")
        validation_error = ValidationError(
            "Chapter assertion failed: Found '摘要' at level 1",
            assertion_failures=["Chapter assertion: no '摘要/参考文献' at level 1"],
            rollback_path=original_docx
        )
        
        context = ErrorContext(
            pipeline_stage="validator",
            docx_path=modified_docx,
            original_docx_path=original_docx
        )
        
        result = handler.handle_pipeline_error(validation_error, context)
        
        print(f"Recovery result:")
        print(f"  Success: {result.success}")
        print(f"  Status: {result.status}")
        print(f"  Rollback performed: {result.rollback_performed}")
        print(f"  Warnings: {result.warnings}")
        print(f"  Recovery path: {result.recovery_path}")
        
        # Show content after rollback
        with open(modified_docx, 'r', encoding='utf-8') as f:
            rollback_content = f.read()
            print(f"Content after rollback: {rollback_content[:50]}...")
        
        # Check if backup was created
        backup_path = f"{modified_docx}.rollback_backup"
        if os.path.exists(backup_path):
            print(f"Backup created: {backup_path}")
            with open(backup_path, 'r', encoding='utf-8') as f:
                backup_content = f.read()
                print(f"Backup content: {backup_content[:50]}...")
        
        # Check status file
        status_file = os.path.join(audit_dir, "result.status.txt")
        if os.path.exists(status_file):
            print(f"\nStatus file content ({status_file}):")
            with open(status_file, 'r', encoding='utf-8') as f:
                print(f.read())


def demo_revision_handling():
    """Demonstrate revision handling strategies."""
    print("\n=== Demo: Revision Handling Strategies ===")
    
    strategies = [
        RevisionHandlingStrategy.ACCEPT_ALL,
        RevisionHandlingStrategy.REJECT_ALL,
        RevisionHandlingStrategy.BYPASS,
        RevisionHandlingStrategy.FAIL_ON_REVISIONS
    ]
    
    for strategy in strategies:
        print(f"\n{strategy.value.upper()} Strategy:")
        
        with tempfile.TemporaryDirectory() as temp_dir:
            audit_dir = os.path.join(temp_dir, "audit")
            handler = PipelineErrorHandler(audit_dir, strategy)
            
            # Create a mock document path
            docx_path = os.path.join(temp_dir, "document_with_revisions.docx")
            with open(docx_path, 'w') as f:
                f.write("Document with tracked changes")
            
            print(f"  Strategy: {strategy.value}")
            print(f"  Document: {docx_path}")
            
            # Note: In real scenario, this would interact with Word COM
            # For demo, we'll show what would happen
            if strategy == RevisionHandlingStrategy.FAIL_ON_REVISIONS:
                print("  → Would fail if revisions are found")
            elif strategy == RevisionHandlingStrategy.ACCEPT_ALL:
                print("  → Would accept all tracked changes")
            elif strategy == RevisionHandlingStrategy.REJECT_ALL:
                print("  → Would reject all tracked changes")
            elif strategy == RevisionHandlingStrategy.BYPASS:
                print("  → Would leave tracked changes unchanged")


def demo_comprehensive_error_handling():
    """Demonstrate comprehensive error handling scenarios."""
    print("\n=== Demo: Comprehensive Error Handling ===")
    
    with tempfile.TemporaryDirectory() as temp_dir:
        audit_dir = os.path.join(temp_dir, "audit")
        handler = PipelineErrorHandler(audit_dir, RevisionHandlingStrategy.BYPASS)
        
        # Create demo files
        original_docx, modified_docx = create_demo_files(temp_dir)
        
        error_scenarios = [
            ("Security Violation", SecurityViolationError(
                "Unauthorized macro execution",
                operation_type="execute_macro",
                security_context={"macro": "AutoOpen", "source": "external"}
            )),
            ("Execution Error", ExecutionError(
                "COM object access failed",
                operation_type="delete_section_by_heading",
                operation_data={"heading_text": "Test", "level": 1}
            )),
            ("Unexpected Error", ValueError("Unexpected system error occurred"))
        ]
        
        for scenario_name, error in error_scenarios:
            print(f"\n{scenario_name}:")
            
            context = ErrorContext(
                pipeline_stage="executor",
                docx_path=modified_docx,
                original_docx_path=original_docx
            )
            
            result = handler.handle_pipeline_error(error, context)
            
            print(f"  Error: {str(error)[:60]}...")
            print(f"  Status: {result.status}")
            print(f"  Rollback: {result.rollback_performed}")
            print(f"  Warnings: {len(result.warnings)} warning(s)")
            print(f"  Errors: {len(result.errors)} error(s)")
        
        # Show final warnings.log
        warnings_log = os.path.join(audit_dir, "warnings.log")
        print(f"\nFinal warnings.log content:")
        with open(warnings_log, 'r', encoding='utf-8') as f:
            lines = f.readlines()
            for line in lines[-10:]:  # Show last 10 lines
                print(f"  {line.strip()}")


def main():
    """Run all error handling demos."""
    print("AutoWord vNext Error Handling and Recovery System Demo")
    print("=" * 60)
    
    try:
        demo_warnings_logging()
        demo_security_validation()
        demo_rollback_functionality()
        demo_revision_handling()
        demo_comprehensive_error_handling()
        
        print("\n" + "=" * 60)
        print("Demo completed successfully!")
        print("\nKey features demonstrated:")
        print("✓ NOOP operation logging")
        print("✓ Font and localization fallback logging")
        print("✓ Security violation detection")
        print("✓ Automatic rollback on errors")
        print("✓ Revision handling strategies")
        print("✓ Comprehensive error recovery")
        print("✓ Audit trail generation")
        
    except Exception as e:
        print(f"\nDemo failed with error: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()