"""
Demo script for AutoWord vNext DocumentExecutor.

This script demonstrates how to use the DocumentExecutor to execute
atomic operations on Word documents through COM automation.
"""

import os
import sys
import tempfile
from pathlib import Path

# Add the project root to the Python path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from autoword.vnext.executor import DocumentExecutor
from autoword.vnext.models import (
    PlanV1, DeleteSectionByHeading, UpdateToc, SetStyleRule,
    FontSpec, ParagraphSpec, LineSpacingMode, MatchMode
)
from autoword.vnext.exceptions import ExecutionError


def demo_basic_operations():
    """Demonstrate basic atomic operations."""
    print("=== AutoWord vNext DocumentExecutor Demo ===\n")
    
    try:
        # Initialize executor
        print("1. Initializing DocumentExecutor...")
        executor = DocumentExecutor()
        print("   ✓ DocumentExecutor initialized successfully")
        
        # Create sample operations
        print("\n2. Creating sample atomic operations...")
        
        # Delete section operation
        delete_op = DeleteSectionByHeading(
            heading_text="摘要",
            level=1,
            match=MatchMode.EXACT,
            case_sensitive=False
        )
        print(f"   ✓ Created delete_section_by_heading: {delete_op.heading_text}")
        
        # Update TOC operation
        update_toc_op = UpdateToc()
        print("   ✓ Created update_toc operation")
        
        # Set style rule operation
        style_op = SetStyleRule(
            target_style_name="Heading 1",
            font=FontSpec(
                east_asian="楷体",
                latin="Times New Roman",
                size_pt=12,
                bold=True
            ),
            paragraph=ParagraphSpec(
                line_spacing_mode=LineSpacingMode.MULTIPLE,
                line_spacing_value=2.0
            )
        )
        print(f"   ✓ Created set_style_rule: {style_op.target_style_name}")
        
        # Create execution plan
        print("\n3. Creating execution plan...")
        plan = PlanV1(ops=[delete_op, style_op, update_toc_op])
        print(f"   ✓ Created plan with {len(plan.ops)} operations")
        
        # Display plan details
        print("\n4. Plan details:")
        for i, op in enumerate(plan.ops, 1):
            print(f"   {i}. {op.operation_type}")
            if hasattr(op, 'heading_text'):
                print(f"      - Heading: {op.heading_text}")
            if hasattr(op, 'target_style_name'):
                print(f"      - Style: {op.target_style_name}")
        
        print("\n5. Execution simulation (no actual Word document):")
        print("   Note: This demo shows the structure without executing on a real document")
        print("   In real usage, you would call:")
        print("   result_path = executor.execute_plan(plan, 'path/to/document.docx')")
        
        # Show localization features
        print("\n6. Localization features:")
        print("   Style aliases:")
        for english, chinese in executor.localization_manager.STYLE_ALIASES.items():
            print(f"   - {english} ↔ {chinese}")
        
        print("\n   Font fallbacks:")
        for font, fallbacks in executor.localization_manager.FONT_FALLBACKS.items():
            print(f"   - {font}: {' → '.join(fallbacks)}")
        
        print("\n=== Demo completed successfully! ===")
        
    except ExecutionError as e:
        print(f"❌ Execution error: {e}")
        if e.details:
            print(f"   Details: {e.details}")
    except Exception as e:
        print(f"❌ Unexpected error: {e}")


def demo_operation_validation():
    """Demonstrate operation validation and error handling."""
    print("\n=== Operation Validation Demo ===\n")
    
    try:
        executor = DocumentExecutor()
        
        # Test valid operation
        print("1. Testing valid operation...")
        valid_op = UpdateToc()
        print(f"   ✓ Valid operation: {valid_op.operation_type}")
        
        # Test operation with validation
        print("\n2. Testing operation with parameters...")
        delete_op = DeleteSectionByHeading(
            heading_text="参考文献",
            level=1,
            match=MatchMode.CONTAINS,
            case_sensitive=False,
            occurrence_index=1
        )
        print(f"   ✓ Delete operation: '{delete_op.heading_text}' (level {delete_op.level})")
        print(f"   ✓ Match mode: {delete_op.match}")
        print(f"   ✓ Case sensitive: {delete_op.case_sensitive}")
        print(f"   ✓ Occurrence index: {delete_op.occurrence_index}")
        
        # Test style operation with complex formatting
        print("\n3. Testing complex style operation...")
        complex_style = SetStyleRule(
            target_style_name="标题 1",
            font=FontSpec(
                east_asian="楷体",
                latin="Times New Roman",
                size_pt=14,
                bold=True,
                italic=False,
                color_hex="#000000"
            ),
            paragraph=ParagraphSpec(
                line_spacing_mode=LineSpacingMode.EXACTLY,
                line_spacing_value=1.8,  # Fixed: must be <= 10
                space_before_pt=12.0,
                space_after_pt=6.0,
                indent_left_pt=0.0,
                indent_first_line_pt=0.0
            )
        )
        print(f"   ✓ Complex style rule for: {complex_style.target_style_name}")
        print(f"   ✓ Font: {complex_style.font.east_asian} / {complex_style.font.latin}")
        print(f"   ✓ Size: {complex_style.font.size_pt}pt, Bold: {complex_style.font.bold}")
        print(f"   ✓ Line spacing: {complex_style.paragraph.line_spacing_mode} ({complex_style.paragraph.line_spacing_value})")
        
        print("\n=== Validation demo completed! ===")
        
    except Exception as e:
        print(f"❌ Error during validation demo: {e}")


if __name__ == "__main__":
    demo_basic_operations()
    demo_operation_validation()