#!/usr/bin/env python3
"""
AutoWord vNext Basic Usage Examples

This module demonstrates basic usage patterns for the AutoWord vNext pipeline.
These examples show the most common operations and usage scenarios.
"""

import os
from pathlib import Path
from autoword.vnext.pipeline import VNextPipeline
from autoword.vnext.exceptions import VNextException


def basic_document_processing():
    """
    Example 1: Basic document processing
    
    This is the most common usage pattern - process a single document
    with a natural language intent.
    """
    print("=== Basic Document Processing ===")
    
    # Initialize pipeline with default configuration
    pipeline = VNextPipeline()
    
    # Process document with simple intent
    try:
        result = pipeline.process_document(
            docx_path="sample_document.docx",
            user_intent="Remove abstract and references sections"
        )
        
        if result.status == "SUCCESS":
            print(f"✓ Document processed successfully")
            print(f"  Output: {result.output_path}")
            print(f"  Audit: {result.audit_directory}")
            print(f"  Time: {result.execution_time:.2f}s")
        else:
            print(f"✗ Processing failed: {result.status}")
            if result.error:
                print(f"  Error: {result.error}")
                
    except VNextException as e:
        print(f"✗ Pipeline error: {e}")
    except Exception as e:
        print(f"✗ Unexpected error: {e}")


def processing_with_configuration():
    """
    Example 2: Processing with custom configuration
    
    Shows how to customize pipeline behavior through configuration.
    """
    print("\n=== Processing with Configuration ===")
    
    # Custom configuration
    config = {
        "model": "gpt4",           # Use GPT-4 for better results
        "temperature": 0.05,       # Lower temperature for consistency
        "verbose": True,           # Enable verbose output
        "audit_dir": "./my_audits" # Custom audit directory
    }
    
    # Initialize pipeline with configuration
    pipeline = VNextPipeline(config)
    
    try:
        result = pipeline.process_document(
            docx_path="research_paper.docx",
            user_intent="Apply standard academic formatting and remove abstract section"
        )
        
        if result.status == "SUCCESS":
            print(f"✓ Document formatted successfully")
            print(f"  Operations executed: {result.operations_executed}")
        else:
            print(f"✗ Formatting failed: {result.status}")
            
    except Exception as e:
        print(f"✗ Error: {e}")


def dry_run_example():
    """
    Example 3: Dry run to preview changes
    
    Use dry run mode to see what changes would be made without
    actually modifying the document.
    """
    print("\n=== Dry Run Example ===")
    
    pipeline = VNextPipeline()
    
    try:
        # Generate plan without executing
        dry_result = pipeline.dry_run(
            docx_path="document.docx",
            user_intent="Update table of contents and apply heading styles"
        )
        
        if dry_result.plan_generated:
            print("✓ Plan generated successfully")
            print(f"  Operations planned: {len(dry_result.plan.ops)}")
            
            # Show planned operations
            for i, op in enumerate(dry_result.plan.ops, 1):
                print(f"  {i}. {op.operation}")
                if hasattr(op, 'heading_text'):
                    print(f"     Target: {op.heading_text}")
                if hasattr(op, 'target_style'):
                    print(f"     Style: {op.target_style}")
        else:
            print(f"✗ Plan generation failed: {dry_result.error}")
            
    except Exception as e:
        print(f"✗ Error: {e}")


def batch_processing_example():
    """
    Example 4: Batch processing multiple documents
    
    Process all DOCX files in a directory with the same intent.
    """
    print("\n=== Batch Processing Example ===")
    
    # Configuration for batch processing
    config = {
        "model": "gpt35",  # Faster model for batch operations
        "audit_dir": "./batch_audits"
    }
    
    pipeline = VNextPipeline(config)
    
    try:
        # Process all documents in directory
        batch_result = pipeline.process_batch(
            input_dir="./documents",
            user_intent="Apply standard formatting and update TOC"
        )
        
        print(f"✓ Batch processing completed")
        print(f"  Total processed: {batch_result.total_processed}")
        print(f"  Successful: {batch_result.successful}")
        print(f"  Failed: {batch_result.failed}")
        print(f"  Total time: {batch_result.execution_time:.2f}s")
        
        # Show individual results
        for i, result in enumerate(batch_result.results, 1):
            status_icon = "✓" if result.status == "SUCCESS" else "✗"
            print(f"  {i}. {status_icon} {result.status}")
            
    except Exception as e:
        print(f"✗ Batch processing error: {e}")


def error_handling_example():
    """
    Example 5: Comprehensive error handling
    
    Shows how to handle different types of errors and failures.
    """
    print("\n=== Error Handling Example ===")
    
    pipeline = VNextPipeline()
    
    # Example with intentionally problematic input
    try:
        result = pipeline.process_document(
            docx_path="nonexistent.docx",  # File doesn't exist
            user_intent="Process document"
        )
        
    except FileNotFoundError:
        print("✗ File not found - check document path")
        
    except VNextException as e:
        print(f"✗ Pipeline error: {e}")
        print("  Check system status and configuration")
        
    except Exception as e:
        print(f"✗ Unexpected error: {e}")
        print("  This may indicate a system issue")
    
    # Example with validation failure
    try:
        result = pipeline.process_document(
            docx_path="complex_document.docx",
            user_intent="Make document perfect"  # Vague intent
        )
        
        # Handle different result statuses
        if result.status == "SUCCESS":
            print("✓ Processing succeeded")
            
        elif result.status == "INVALID_PLAN":
            print("✗ Could not generate valid plan")
            print("  Try making your intent more specific")
            print("  Example: 'Remove abstract section and update TOC'")
            
        elif result.status == "FAILED_VALIDATION":
            print("✗ Changes failed validation")
            print("  Document was automatically restored")
            print(f"  Check audit trail: {result.audit_directory}")
            
        elif result.status == "ROLLBACK":
            print("✗ Error occurred, changes rolled back")
            print("  Document is safe and unchanged")
            print(f"  Error details: {result.error}")
            
    except Exception as e:
        print(f"✗ Error: {e}")


def working_with_audit_trails():
    """
    Example 6: Working with audit trails
    
    Shows how to examine and use audit trail information.
    """
    print("\n=== Working with Audit Trails ===")
    
    pipeline = VNextPipeline({
        "audit_dir": "./detailed_audits",
        "verbose": True
    })
    
    try:
        result = pipeline.process_document(
            docx_path="sample.docx",
            user_intent="Remove introduction section"
        )
        
        if result.audit_directory:
            audit_path = Path(result.audit_directory)
            print(f"✓ Audit trail created: {audit_path}")
            
            # List audit files
            print("  Audit files:")
            for file in audit_path.glob("*"):
                print(f"    - {file.name}")
            
            # Read status file
            status_file = audit_path / "result.status.txt"
            if status_file.exists():
                status_content = status_file.read_text()
                print(f"  Final status: {status_content.strip()}")
            
            # Check for warnings
            warnings_file = audit_path / "warnings.log"
            if warnings_file.exists():
                warnings_content = warnings_file.read_text()
                if warnings_content.strip():
                    print(f"  Warnings: {warnings_content.strip()}")
                else:
                    print("  No warnings")
                    
    except Exception as e:
        print(f"✗ Error: {e}")


def common_document_operations():
    """
    Example 7: Common document operations
    
    Shows typical document processing operations with specific intents.
    """
    print("\n=== Common Document Operations ===")
    
    pipeline = VNextPipeline({"model": "gpt4"})
    
    # Common operations with example intents
    operations = [
        ("Remove sections", "Remove abstract and references sections"),
        ("Update TOC", "Update table of contents and page numbers"),
        ("Apply formatting", "Apply Heading 1 style with 楷体 font to chapter titles"),
        ("Standardize styles", "Set all headings to use standard fonts and spacing"),
        ("Clean formatting", "Remove direct formatting and apply consistent styles")
    ]
    
    for operation_name, intent in operations:
        print(f"\n--- {operation_name} ---")
        
        try:
            # Use dry run to show what would happen
            dry_result = pipeline.dry_run("sample.docx", intent)
            
            if dry_result.plan_generated:
                print(f"✓ Plan for '{operation_name}':")
                for op in dry_result.plan.ops:
                    print(f"  - {op.operation}")
            else:
                print(f"✗ Could not plan '{operation_name}': {dry_result.error}")
                
        except Exception as e:
            print(f"✗ Error planning '{operation_name}': {e}")


def main():
    """
    Run all basic usage examples
    """
    print("AutoWord vNext Basic Usage Examples")
    print("=" * 50)
    
    # Check if we have sample documents
    if not Path("sample_document.docx").exists():
        print("Note: Some examples require sample documents to be present")
        print("Create sample DOCX files or modify paths in examples")
        print()
    
    # Run examples
    basic_document_processing()
    processing_with_configuration()
    dry_run_example()
    batch_processing_example()
    error_handling_example()
    working_with_audit_trails()
    common_document_operations()
    
    print("\n" + "=" * 50)
    print("Basic usage examples completed!")
    print("\nNext steps:")
    print("- Try these examples with your own documents")
    print("- Explore advanced examples in other files")
    print("- Read the User Guide for more detailed information")


if __name__ == "__main__":
    main()