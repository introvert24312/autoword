"""
Demo script for DocumentAuditor functionality.

This script demonstrates the complete audit trail generation process
including timestamped directories, snapshots, diff reports, and status logging.
"""

import os
import tempfile
import shutil
from datetime import datetime
from pathlib import Path

from autoword.vnext.auditor.document_auditor import DocumentAuditor
from autoword.vnext.models import (
    StructureV1, PlanV1, DocumentMetadata, 
    ParagraphSkeleton, HeadingReference, StyleDefinition,
    FontSpec, ParagraphSpec, DeleteSectionByHeading, UpdateToc, SetStyleRule
)
from autoword.vnext.exceptions import AuditError


def create_sample_structures():
    """Create sample before/after structures for demonstration."""
    
    # Before structure - original document
    before_structure = StructureV1(
        metadata=DocumentMetadata(
            title="Research Paper Draft",
            author="Dr. Jane Smith",
            creation_time=datetime(2024, 1, 15, 10, 30),
            modified_time=datetime(2024, 1, 15, 14, 45),
            page_count=8,
            paragraph_count=45,
            word_count=2500
        ),
        paragraphs=[
            ParagraphSkeleton(index=0, style_name="Title", preview_text="Advanced Machine Learning Techniques", is_heading=False),
            ParagraphSkeleton(index=1, style_name="Normal", preview_text="This paper presents novel approaches to...", is_heading=False),
            ParagraphSkeleton(index=2, style_name="Heading 1", preview_text="摘要", is_heading=True, heading_level=1),
            ParagraphSkeleton(index=3, style_name="Normal", preview_text="本文提出了一种新的机器学习方法...", is_heading=False),
            ParagraphSkeleton(index=4, style_name="Heading 1", preview_text="Introduction", is_heading=True, heading_level=1),
            ParagraphSkeleton(index=5, style_name="Normal", preview_text="Machine learning has revolutionized...", is_heading=False),
            ParagraphSkeleton(index=6, style_name="Heading 2", preview_text="Background", is_heading=True, heading_level=2),
            ParagraphSkeleton(index=7, style_name="Normal", preview_text="Previous research in this area...", is_heading=False),
            ParagraphSkeleton(index=8, style_name="Heading 1", preview_text="Methodology", is_heading=True, heading_level=1),
            ParagraphSkeleton(index=9, style_name="Normal", preview_text="Our approach consists of three phases...", is_heading=False),
            ParagraphSkeleton(index=10, style_name="Heading 1", preview_text="Results", is_heading=True, heading_level=1),
            ParagraphSkeleton(index=11, style_name="Normal", preview_text="The experimental results show...", is_heading=False),
            ParagraphSkeleton(index=12, style_name="Heading 1", preview_text="参考文献", is_heading=True, heading_level=1),
            ParagraphSkeleton(index=13, style_name="Normal", preview_text="[1] Smith, J. et al. (2023)...", is_heading=False),
        ],
        headings=[
            HeadingReference(paragraph_index=2, level=1, text="摘要", style_name="Heading 1"),
            HeadingReference(paragraph_index=4, level=1, text="Introduction", style_name="Heading 1"),
            HeadingReference(paragraph_index=6, level=2, text="Background", style_name="Heading 2"),
            HeadingReference(paragraph_index=8, level=1, text="Methodology", style_name="Heading 1"),
            HeadingReference(paragraph_index=10, level=1, text="Results", style_name="Heading 1"),
            HeadingReference(paragraph_index=12, level=1, text="参考文献", style_name="Heading 1"),
        ],
        styles=[
            StyleDefinition(
                name="Normal",
                type="paragraph",
                font=FontSpec(latin="Times New Roman", east_asian="宋体", size_pt=12),
                paragraph=ParagraphSpec(line_spacing_mode="MULTIPLE", line_spacing_value=2.0)
            ),
            StyleDefinition(
                name="Heading 1",
                type="paragraph",
                font=FontSpec(latin="Times New Roman", east_asian="黑体", size_pt=14, bold=True),
                paragraph=ParagraphSpec(line_spacing_mode="MULTIPLE", line_spacing_value=2.0)
            ),
            StyleDefinition(
                name="Heading 2",
                type="paragraph",
                font=FontSpec(latin="Times New Roman", east_asian="黑体", size_pt=12, bold=True),
                paragraph=ParagraphSpec(line_spacing_mode="MULTIPLE", line_spacing_value=2.0)
            ),
            StyleDefinition(
                name="Title",
                type="paragraph",
                font=FontSpec(latin="Times New Roman", east_asian="黑体", size_pt=16, bold=True),
                paragraph=ParagraphSpec(line_spacing_mode="MULTIPLE", line_spacing_value=1.5)
            )
        ]
    )
    
    # After structure - processed document (摘要 and 参考文献 sections removed, styles updated)
    after_structure = StructureV1(
        metadata=DocumentMetadata(
            title="Research Paper Draft",
            author="Dr. Jane Smith",
            creation_time=datetime(2024, 1, 15, 10, 30),
            modified_time=datetime(2024, 1, 15, 15, 20),  # Updated modification time
            page_count=6,  # Reduced page count
            paragraph_count=35,  # Reduced paragraph count
            word_count=2000  # Reduced word count
        ),
        paragraphs=[
            ParagraphSkeleton(index=0, style_name="Title", preview_text="Advanced Machine Learning Techniques", is_heading=False),
            ParagraphSkeleton(index=1, style_name="Normal", preview_text="This paper presents novel approaches to...", is_heading=False),
            # Note: 摘要 section removed (paragraphs 2-3)
            ParagraphSkeleton(index=4, style_name="Heading 1", preview_text="Introduction", is_heading=True, heading_level=1),
            ParagraphSkeleton(index=5, style_name="Normal", preview_text="Machine learning has revolutionized...", is_heading=False),
            ParagraphSkeleton(index=6, style_name="Heading 2", preview_text="Background", is_heading=True, heading_level=2),
            ParagraphSkeleton(index=7, style_name="Normal", preview_text="Previous research in this area...", is_heading=False),
            ParagraphSkeleton(index=8, style_name="Heading 1", preview_text="Methodology", is_heading=True, heading_level=1),
            ParagraphSkeleton(index=9, style_name="Normal", preview_text="Our approach consists of three phases...", is_heading=False),
            ParagraphSkeleton(index=10, style_name="Heading 1", preview_text="Results", is_heading=True, heading_level=1),
            ParagraphSkeleton(index=11, style_name="Normal", preview_text="The experimental results show...", is_heading=False),
            # Note: 参考文献 section removed (paragraphs 12-13)
        ],
        headings=[
            # Note: 摘要 heading removed
            HeadingReference(paragraph_index=4, level=1, text="Introduction", style_name="Heading 1"),
            HeadingReference(paragraph_index=6, level=2, text="Background", style_name="Heading 2"),
            HeadingReference(paragraph_index=8, level=1, text="Methodology", style_name="Heading 1"),
            HeadingReference(paragraph_index=10, level=1, text="Results", style_name="Heading 1"),
            # Note: 参考文献 heading removed
        ],
        styles=[
            StyleDefinition(
                name="Normal",
                type="paragraph",
                font=FontSpec(latin="Times New Roman", east_asian="宋体", size_pt=12),
                paragraph=ParagraphSpec(line_spacing_mode="MULTIPLE", line_spacing_value=2.0)
            ),
            StyleDefinition(
                name="Heading 1",
                type="paragraph",
                font=FontSpec(latin="Times New Roman", east_asian="楷体", size_pt=12, bold=True),  # Font changed
                paragraph=ParagraphSpec(line_spacing_mode="MULTIPLE", line_spacing_value=2.0)
            ),
            StyleDefinition(
                name="Heading 2",
                type="paragraph",
                font=FontSpec(latin="Times New Roman", east_asian="宋体", size_pt=12, bold=True),  # Font changed
                paragraph=ParagraphSpec(line_spacing_mode="MULTIPLE", line_spacing_value=2.0)
            ),
            StyleDefinition(
                name="Title",
                type="paragraph",
                font=FontSpec(latin="Times New Roman", east_asian="黑体", size_pt=16, bold=True),
                paragraph=ParagraphSpec(line_spacing_mode="MULTIPLE", line_spacing_value=1.5)
            )
        ]
    )
    
    # Execution plan that was applied
    plan = PlanV1(
        ops=[
            DeleteSectionByHeading(
                heading_text="摘要",
                level=1,
                match="EXACT",
                case_sensitive=False
            ),
            DeleteSectionByHeading(
                heading_text="参考文献",
                level=1,
                match="EXACT",
                case_sensitive=False
            ),
            SetStyleRule(
                target_style_name="Heading 1",
                font=FontSpec(east_asian="楷体", size_pt=12, bold=True),
                paragraph=ParagraphSpec(line_spacing_mode="MULTIPLE", line_spacing_value=2.0)
            ),
            SetStyleRule(
                target_style_name="Heading 2",
                font=FontSpec(east_asian="宋体", size_pt=12, bold=True),
                paragraph=ParagraphSpec(line_spacing_mode="MULTIPLE", line_spacing_value=2.0)
            ),
            UpdateToc()
        ]
    )
    
    return before_structure, after_structure, plan


def create_mock_docx_files(temp_dir):
    """Create mock DOCX files for demonstration."""
    before_docx = os.path.join(temp_dir, "research_paper_before.docx")
    after_docx = os.path.join(temp_dir, "research_paper_after.docx")
    
    # Create mock DOCX content (in reality these would be actual Word documents)
    with open(before_docx, 'wb') as f:
        f.write(b"PK\x03\x04" + b"Mock DOCX content - original document with abstract and references")
    
    with open(after_docx, 'wb') as f:
        f.write(b"PK\x03\x04" + b"Mock DOCX content - processed document without abstract and references")
    
    return before_docx, after_docx


def demonstrate_audit_trail():
    """Demonstrate complete audit trail generation."""
    print("=== DocumentAuditor Demo ===\n")
    
    # Create temporary directory for demo
    temp_dir = tempfile.mkdtemp()
    print(f"Demo directory: {temp_dir}")
    
    try:
        # Initialize auditor
        auditor = DocumentAuditor(base_audit_dir=os.path.join(temp_dir, "audit_trails"))
        print("✓ DocumentAuditor initialized")
        
        # Create sample data
        before_structure, after_structure, plan = create_sample_structures()
        before_docx, after_docx = create_mock_docx_files(temp_dir)
        print("✓ Sample data created")
        
        # Step 1: Create audit directory
        print("\n--- Step 1: Creating Audit Directory ---")
        audit_dir = auditor.create_audit_directory()
        print(f"✓ Audit directory created: {audit_dir}")
        print(f"  - Subdirectories: snapshots/, structures/, reports/")
        
        # Step 2: Add some warnings during processing
        print("\n--- Step 2: Adding Warnings ---")
        auditor.add_warning("Font fallback applied: 楷体 -> 楷体_GB2312 (original font not available)")
        auditor.add_warning("NOOP operation: update_toc (no TOC found in document)")
        auditor.add_warning("Style alias used: Heading 1 -> 标题 1")
        print(f"✓ Added {len(auditor.warnings)} warnings")
        
        # Step 3: Save snapshots
        print("\n--- Step 3: Saving Snapshots ---")
        auditor.save_snapshots(
            before_docx=before_docx,
            after_docx=after_docx,
            before_structure=before_structure,
            after_structure=after_structure,
            plan=plan
        )
        print("✓ Snapshots saved:")
        print("  - before.docx and after.docx")
        print("  - structure.before.v1.json and structure.after.v1.json")
        print("  - plan.v1.json")
        
        # Step 4: Generate diff report
        print("\n--- Step 4: Generating Diff Report ---")
        diff_report = auditor.generate_diff_report(before_structure, after_structure)
        print("✓ Diff report generated:")
        print(f"  - Summary: {diff_report.summary}")
        print(f"  - Removed paragraphs: {diff_report.removed_paragraphs}")
        print(f"  - Style changes: {list(diff_report.style_changes.keys())}")
        print(f"  - Heading changes: {len(diff_report.heading_changes)} changes")
        
        # Step 5: Finalize audit (write warnings and status)
        print("\n--- Step 5: Finalizing Audit ---")
        auditor.finalize_audit("SUCCESS", "Document processed successfully with 2 sections removed and styles updated")
        print("✓ Audit finalized:")
        print("  - warnings.log written")
        print("  - result.status.txt written")
        
        # Step 6: Show audit directory contents
        print("\n--- Step 6: Audit Directory Contents ---")
        audit_path = Path(audit_dir)
        print(f"Audit directory: {audit_path.name}")
        
        for root, dirs, files in os.walk(audit_path):
            level = root.replace(str(audit_path), '').count(os.sep)
            indent = ' ' * 2 * level
            print(f"{indent}{os.path.basename(root)}/")
            subindent = ' ' * 2 * (level + 1)
            for file in files:
                file_path = os.path.join(root, file)
                file_size = os.path.getsize(file_path)
                print(f"{subindent}{file} ({file_size} bytes)")
        
        # Step 7: Show sample file contents
        print("\n--- Step 7: Sample File Contents ---")
        
        # Show warnings.log
        warnings_file = audit_path / "warnings.log"
        if warnings_file.exists():
            print("\nwarnings.log:")
            with open(warnings_file, 'r', encoding='utf-8') as f:
                for line in f:
                    print(f"  {line.strip()}")
        
        # Show result.status.txt
        status_file = audit_path / "result.status.txt"
        if status_file.exists():
            print("\nresult.status.txt:")
            with open(status_file, 'r', encoding='utf-8') as f:
                for line in f:
                    print(f"  {line.strip()}")
        
        # Show diff report summary
        diff_file = audit_path / "reports" / "diff.report.json"
        if diff_file.exists():
            print("\ndiff.report.json (summary):")
            import json
            with open(diff_file, 'r', encoding='utf-8') as f:
                diff_data = json.load(f)
                print(f"  Summary: {diff_data['summary']}")
                print(f"  Removed paragraphs: {diff_data['removed_paragraphs']}")
                print(f"  Style changes: {len(diff_data['style_changes'])} styles")
        
        print(f"\n✓ Demo completed successfully!")
        print(f"✓ Complete audit trail available at: {audit_dir}")
        
    except AuditError as e:
        print(f"❌ Audit error: {e}")
        print(f"   Details: {e.details}")
        
    except Exception as e:
        print(f"❌ Unexpected error: {e}")
        
    finally:
        # Clean up
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)
            print(f"\n🧹 Cleaned up demo directory: {temp_dir}")


def demonstrate_error_handling():
    """Demonstrate error handling scenarios."""
    print("\n=== Error Handling Demo ===\n")
    
    # Test 1: No audit directory
    print("--- Test 1: Operations without audit directory ---")
    auditor = DocumentAuditor()
    
    try:
        auditor.write_status("SUCCESS", "Test")
    except AuditError as e:
        print(f"✓ Expected error caught: {e.message}")
        print(f"  Stage: {e.details.get('audit_stage')}")
    
    # Test 2: Invalid status
    print("\n--- Test 2: Invalid status value ---")
    temp_dir = tempfile.mkdtemp()
    try:
        auditor = DocumentAuditor(base_audit_dir=temp_dir)
        auditor.create_audit_directory()
        
        try:
            auditor.write_status("INVALID_STATUS", "Test")
        except AuditError as e:
            print(f"✓ Expected error caught: {e.message}")
            print(f"  Stage: {e.details.get('audit_stage')}")
    
    finally:
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)
    
    # Test 3: Directory creation failure
    print("\n--- Test 3: Directory creation failure ---")
    try:
        auditor = DocumentAuditor(base_audit_dir="/invalid/path/that/cannot/exist")
        auditor.create_audit_directory()
    except AuditError as e:
        print(f"✓ Expected error caught: {e.message}")
        print(f"  Stage: {e.details.get('audit_stage')}")
    
    print("\n✓ Error handling demo completed")


if __name__ == "__main__":
    demonstrate_audit_trail()
    demonstrate_error_handling()