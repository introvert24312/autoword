#!/usr/bin/env python3
"""
AutoWord vNext - Basic Usage Examples

This script demonstrates basic usage patterns for AutoWord vNext,
including simple document processing, plan generation, and result handling.
"""

import os
import sys
from pathlib import Path

# Add autoword to path for examples
sys.path.insert(0, str(Path(__file__).parent.parent.parent))

from autoword.vnext import VNextPipeline
from autoword.vnext.core import VNextConfig, LLMConfig


def basic_document_processing():
    """Basic document processing example"""
    print("=== Basic Document Processing ===")
    
    # Initialize pipeline with default configuration
    pipeline = VNextPipeline()
    
    # Process a document
    document_path = "sample_document.docx"
    user_intent = "Delete the abstract section and update the table of contents"
    
    try:
        result = pipeline.process_document(document_path, user_intent)
        
        if result.status == "SUCCESS":
            print(f"✅ Processing successful!")
            print(f"   Output: {result.output_path}")
            print(f"   Audit: {result.audit_directory}")
            print(f"   Time: {result.processing_time:.2f}s")
        else:
            print(f"❌ Processing failed: {result.status}")
            if result.error:
                print(f"   Error: {result.error}")
            if result.validation_errors:
                print(f"   Validation errors: {result.validation_errors}")
                
    except Exception as e:
        print(f"❌ Exception occurred: {e}")


def dry_run_example():
    """Generate plan without execution (dry run)"""
    print("\n=== Dry Run (Plan Generation) ===")
    
    pipeline = VNextPipeline()
    
    document_path = "sample_document.docx"
    user_intent = "Set all Heading 1 styles to use 楷体 font, 12pt, bold"
    
    try:
        result = pipeline.generate_plan(document_path, user_intent)
        
        if result.is_valid:
            print(f"✅ Plan generated successfully!")
            print(f"   Operations: {len(result.plan.ops)}")
            
            for i, op in enumerate(result.plan.ops, 1):
                print(f"   {i}. {op.op_type}")
                if hasattr(op, 'heading_text'):
                    print(f"      - Heading: {op.heading_text}")
                if hasattr(op, 'target_style_name'):
                    print(f"      - Style: {op.target_style_name}")
        else:
            print(f"❌ Invalid plan generated")
            print(f"   Errors: {result.errors}")
            
    except Exception as e:
        print(f"❌ Exception occurred: {e}")


def document_validation_example():
    """Validate document structure"""
    print("\n=== Document Validation ===")
    
    pipeline = VNextPipeline()
    
    document_path = "sample_document.docx"
    
    try:
        result = pipeline.validate_document(document_path)
        
        if result.is_valid:
            print("✅ Document validation passed")
        else:
            print("❌ Document validation failed")
            for error in result.errors:
                print(f"   - {error}")
                
        if result.warnings:
            print("⚠️  Warnings:")
            for warning in result.warnings:
                print(f"   - {warning}")
                
    except Exception as e:
        print(f"❌ Exception occurred: {e}")


def multiple_operations_example():
    """Process document with multiple operations"""
    print("\n=== Multiple Operations ===")
    
    pipeline = VNextPipeline()
    
    document_path = "academic_paper.docx"
    user_intent = """
    1. Delete the abstract section
    2. Delete the references section  
    3. Update the table of contents
    4. Set Heading 1 style to use 楷体 font, 12pt, bold
    5. Set Normal style to use 宋体 font, 12pt
    """
    
    try:
        result = pipeline.process_document(document_path, user_intent)
        
        print(f"Status: {result.status}")
        
        if result.status == "SUCCESS":
            print("✅ All operations completed successfully")
            
            # Read audit information
            audit_dir = Path(result.audit_directory)
            if audit_dir.exists():
                status_file = audit_dir / "result.status.txt"
                if status_file.exists():
                    print(f"   Final status: {status_file.read_text().strip()}")
                    
                warnings_file = audit_dir / "warnings.log"
                if warnings_file.exists():
                    warnings = warnings_file.read_text().strip()
                    if warnings:
                        print(f"   Warnings: {warnings}")
        else:
            print(f"❌ Processing failed: {result.status}")
            
    except Exception as e:
        print(f"❌ Exception occurred: {e}")


def check_processing_result(result):
    """Helper function to check and display processing results"""
    print(f"\nProcessing Result:")
    print(f"  Status: {result.status}")
    
    if result.status == "SUCCESS":
        print(f"  ✅ Success!")
        print(f"  Output: {result.output_path}")
        print(f"  Audit: {result.audit_directory}")
        if result.processing_time:
            print(f"  Time: {result.processing_time:.2f}s")
            
    elif result.status == "FAILED_VALIDATION":
        print(f"  ❌ Validation failed")
        if result.validation_errors:
            for error in result.validation_errors:
                print(f"    - {error}")
                
    elif result.status == "INVALID_PLAN":
        print(f"  ❌ Invalid plan")
        if result.plan_errors:
            for error in result.plan_errors:
                print(f"    - {error}")
                
    elif result.status == "ROLLBACK":
        print(f"  ⚠️  Rolled back due to error")
        if result.error:
            print(f"    Error: {result.error}")
    
    return result.status == "SUCCESS"


def main():
    """Run all basic usage examples"""
    print("AutoWord vNext - Basic Usage Examples")
    print("=" * 50)
    
    # Check if sample documents exist
    sample_docs = ["sample_document.docx", "academic_paper.docx"]
    missing_docs = [doc for doc in sample_docs if not os.path.exists(doc)]
    
    if missing_docs:
        print("⚠️  Sample documents not found:")
        for doc in missing_docs:
            print(f"   - {doc}")
        print("\nPlease create sample documents or update paths in the examples.")
        print("The examples will show the expected behavior patterns.\n")
    
    # Run examples
    try:
        basic_document_processing()
        dry_run_example()
        document_validation_example()
        multiple_operations_example()
        
    except ImportError as e:
        print(f"❌ Import error: {e}")
        print("Please ensure AutoWord vNext is properly installed.")
        
    except Exception as e:
        print(f"❌ Unexpected error: {e}")
        
    print("\n" + "=" * 50)
    print("Examples completed. Check the User Guide for more details.")


if __name__ == "__main__":
    main()