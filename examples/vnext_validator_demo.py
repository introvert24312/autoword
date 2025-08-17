"""
Demo script for AutoWord vNext DocumentValidator.

This script demonstrates the comprehensive validation and rollback capabilities
of the DocumentValidator module, including:
- Chapter assertions (no 摘要/参考文献 at level 1)
- Style assertions (H1/H2/Normal specifications)
- TOC assertions (consistency with heading tree)
- Pagination assertions (fields updated, repagination)
- Rollback functionality
"""

import os
import tempfile
import logging
from datetime import datetime
from pathlib import Path

from autoword.vnext.validator import DocumentValidator
from autoword.vnext.models import (
    StructureV1, DocumentMetadata, StyleDefinition, ParagraphSkeleton,
    HeadingReference, FieldReference, ValidationResult, FontSpec, 
    ParagraphSpec, LineSpacingMode, StyleType
)
from autoword.vnext.exceptions import ValidationError, RollbackError


# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


def create_sample_structure(valid=True):
    """Create a sample document structure for testing."""
    
    # Create metadata
    metadata = DocumentMetadata(
        title="Sample Document",
        author="Demo Author",
        creation_time=datetime(2024, 1, 1, 10, 0, 0),
        modified_time=datetime(2024, 1, 1, 11, 0, 0),
        word_version="16.0",
        page_count=3,
        paragraph_count=10,
        word_count=200
    )
    
    # Create styles with correct specifications
    styles = [
        StyleDefinition(
            name="Heading 1",
            type=StyleType.PARAGRAPH,
            font=FontSpec(
                east_asian="楷体" if valid else "宋体",  # Wrong font for invalid case
                latin="Times New Roman",
                size_pt=12,
                bold=True
            ),
            paragraph=ParagraphSpec(
                line_spacing_mode=LineSpacingMode.MULTIPLE,
                line_spacing_value=2.0
            )
        ),
        StyleDefinition(
            name="Heading 2",
            type=StyleType.PARAGRAPH,
            font=FontSpec(
                east_asian="宋体",
                latin="Times New Roman",
                size_pt=12,
                bold=True
            ),
            paragraph=ParagraphSpec(
                line_spacing_mode=LineSpacingMode.MULTIPLE,
                line_spacing_value=2.0
            )
        ),
        StyleDefinition(
            name="Normal",
            type=StyleType.PARAGRAPH,
            font=FontSpec(
                east_asian="宋体",
                latin="Times New Roman",
                size_pt=12,
                bold=False
            ),
            paragraph=ParagraphSpec(
                line_spacing_mode=LineSpacingMode.MULTIPLE,
                line_spacing_value=2.0
            )
        )
    ]
    
    # Create headings (valid or invalid based on parameter)
    if valid:
        headings_data = [
            ("Introduction", 1),
            ("Background", 2),
            ("Methodology", 1),
            ("Results", 2),
            ("Conclusion", 1)
        ]
    else:
        headings_data = [
            ("摘要", 1),  # Invalid - forbidden at level 1
            ("Introduction", 1),
            ("Methodology", 1)
        ]
    
    paragraphs = []
    headings = []
    
    for i, (heading_text, level) in enumerate(headings_data):
        paragraph = ParagraphSkeleton(
            index=i,
            style_name=f"Heading {level}",
            preview_text=heading_text,
            is_heading=True,
            heading_level=level
        )
        paragraphs.append(paragraph)
        
        heading = HeadingReference(
            paragraph_index=i,
            level=level,
            text=heading_text,
            style_name=f"Heading {level}"
        )
        headings.append(heading)
    
    # Add some normal paragraphs
    for i in range(len(headings_data), len(headings_data) + 3):
        paragraph = ParagraphSkeleton(
            index=i,
            style_name="Normal",
            preview_text=f"This is paragraph {i} with some sample content for demonstration purposes.",
            is_heading=False,
            heading_level=None
        )
        paragraphs.append(paragraph)
    
    # Create TOC field
    fields = [
        FieldReference(
            paragraph_index=0,
            field_type="TOC",
            field_code="TOC \\o \"1-3\" \\h \\z \\u",
            result_text="Introduction\t1\nBackground\t2\nMethodology\t3\nResults\t4\nConclusion\t5"
        )
    ]
    
    return StructureV1(
        metadata=metadata,
        styles=styles,
        paragraphs=paragraphs,
        headings=headings,
        fields=fields,
        tables=[]
    )


def demo_chapter_assertions():
    """Demonstrate chapter assertion validation."""
    print("\n" + "="*60)
    print("DEMO: Chapter Assertions")
    print("="*60)
    
    validator = DocumentValidator(visible=False)
    
    # Test valid structure
    print("\n1. Testing valid structure (no forbidden headings at level 1)...")
    valid_structure = create_sample_structure(valid=True)
    errors = validator.check_chapter_assertions(valid_structure)
    print(f"   Errors found: {len(errors)}")
    for error in errors:
        print(f"   - {error}")
    
    # Test invalid structure
    print("\n2. Testing invalid structure (摘要 at level 1)...")
    invalid_structure = create_sample_structure(valid=False)
    errors = validator.check_chapter_assertions(invalid_structure)
    print(f"   Errors found: {len(errors)}")
    for error in errors:
        print(f"   - {error}")


def demo_style_assertions():
    """Demonstrate style assertion validation."""
    print("\n" + "="*60)
    print("DEMO: Style Assertions")
    print("="*60)
    
    validator = DocumentValidator(visible=False)
    
    # Test valid styles
    print("\n1. Testing valid style specifications...")
    valid_structure = create_sample_structure(valid=True)
    errors = validator.check_style_assertions(valid_structure)
    print(f"   Errors found: {len(errors)}")
    for error in errors:
        print(f"   - {error}")
    
    # Test invalid styles
    print("\n2. Testing invalid style specifications...")
    invalid_structure = create_sample_structure(valid=False)
    errors = validator.check_style_assertions(invalid_structure)
    print(f"   Errors found: {len(errors)}")
    for error in errors:
        print(f"   - {error}")


def demo_toc_assertions():
    """Demonstrate TOC assertion validation."""
    print("\n" + "="*60)
    print("DEMO: TOC Assertions")
    print("="*60)
    
    validator = DocumentValidator(visible=False)
    
    # Test with TOC
    print("\n1. Testing TOC consistency...")
    structure = create_sample_structure(valid=True)
    errors = validator.check_toc_assertions(structure)
    print(f"   Errors found: {len(errors)}")
    for error in errors:
        print(f"   - {error}")
    
    # Test without TOC
    print("\n2. Testing document without TOC...")
    structure.fields = []  # Remove TOC
    errors = validator.check_toc_assertions(structure)
    print(f"   Errors found: {len(errors)}")
    for error in errors:
        print(f"   - {error}")


def demo_pagination_assertions():
    """Demonstrate pagination assertion validation."""
    print("\n" + "="*60)
    print("DEMO: Pagination Assertions")
    print("="*60)
    
    validator = DocumentValidator(visible=False)
    
    # Test valid pagination
    print("\n1. Testing valid pagination (modification time changed)...")
    original_structure = create_sample_structure(valid=True)
    original_structure.metadata.modified_time = datetime(2024, 1, 1, 10, 0, 0)
    
    modified_structure = create_sample_structure(valid=True)
    modified_structure.metadata.modified_time = datetime(2024, 1, 1, 11, 0, 0)  # Later
    
    errors = validator.check_pagination_assertions(original_structure, modified_structure)
    print(f"   Errors found: {len(errors)}")
    for error in errors:
        print(f"   - {error}")
    
    # Test invalid pagination
    print("\n2. Testing invalid pagination (no time change)...")
    same_time = datetime(2024, 1, 1, 10, 0, 0)
    original_structure.metadata.modified_time = same_time
    modified_structure.metadata.modified_time = same_time
    
    errors = validator.check_pagination_assertions(original_structure, modified_structure)
    print(f"   Errors found: {len(errors)}")
    for error in errors:
        print(f"   - {error}")


def demo_rollback_functionality():
    """Demonstrate rollback functionality."""
    print("\n" + "="*60)
    print("DEMO: Rollback Functionality")
    print("="*60)
    
    validator = DocumentValidator(visible=False)
    
    with tempfile.TemporaryDirectory() as temp_dir:
        # Create test files
        original_file = os.path.join(temp_dir, "original.docx")
        modified_file = os.path.join(temp_dir, "modified.docx")
        
        print(f"\n1. Creating test files in: {temp_dir}")
        
        # Write test content
        with open(original_file, 'w') as f:
            f.write("Original document content")
        with open(modified_file, 'w') as f:
            f.write("Modified document content")
        
        print(f"   Original file: {original_file}")
        print(f"   Modified file: {modified_file}")
        
        # Show initial content
        with open(modified_file, 'r') as f:
            initial_content = f.read()
        print(f"   Initial modified content: '{initial_content}'")
        
        # Perform rollback
        print("\n2. Performing rollback...")
        try:
            validator.rollback_document(original_file, modified_file)
            print("   Rollback successful!")
            
            # Show content after rollback
            with open(modified_file, 'r') as f:
                rollback_content = f.read()
            print(f"   Content after rollback: '{rollback_content}'")
            
            # Check if backup was created
            backup_file = f"{modified_file}.rollback_backup"
            if os.path.exists(backup_file):
                print(f"   Backup created: {backup_file}")
            
        except RollbackError as e:
            print(f"   Rollback failed: {e}")


def demo_text_matching():
    """Demonstrate text matching functionality."""
    print("\n" + "="*60)
    print("DEMO: Text Matching")
    print("="*60)
    
    validator = DocumentValidator(visible=False)
    
    test_cases = [
        ("Introduction", "Introduction", True),
        ("Introduction", "INTRODUCTION", True),
        ("Intro", "Introduction", True),
        ("Introduction", "Introductio", True),
        ("Introduction", "Conclusion", False),
        ("", "Introduction", False),
        ("Introduction", "", False)
    ]
    
    print("\nTesting approximate text matching:")
    for text1, text2, expected in test_cases:
        result = validator._text_matches_approximately(text1, text2)
        status = "✓" if result == expected else "✗"
        print(f"   {status} '{text1}' vs '{text2}' -> {result} (expected {expected})")


def demo_toc_parsing():
    """Demonstrate TOC parsing functionality."""
    print("\n" + "="*60)
    print("DEMO: TOC Parsing")
    print("="*60)
    
    validator = DocumentValidator(visible=False)
    
    # Test TOC parsing
    toc_text = "Introduction\t1\n  Background\t2\nMethodology\t3\n    Results\t4\nConclusion\t5"
    
    print(f"\nParsing TOC text:")
    print(f"'{toc_text}'")
    
    entries = validator._parse_toc_entries(toc_text)
    
    print(f"\nParsed entries ({len(entries)} total):")
    for i, (text, level) in enumerate(entries):
        print(f"   {i+1}. '{text}' (level {level})")


def main():
    """Run all validator demos."""
    print("AutoWord vNext DocumentValidator Demo")
    print("====================================")
    
    try:
        demo_chapter_assertions()
        demo_style_assertions()
        demo_toc_assertions()
        demo_pagination_assertions()
        demo_rollback_functionality()
        demo_text_matching()
        demo_toc_parsing()
        
        print("\n" + "="*60)
        print("DEMO COMPLETED SUCCESSFULLY")
        print("="*60)
        print("\nThe DocumentValidator module provides comprehensive validation")
        print("capabilities for the AutoWord vNext pipeline, including:")
        print("- Chapter assertions (forbidden headings at level 1)")
        print("- Style assertions (font, size, spacing specifications)")
        print("- TOC assertions (consistency with heading tree)")
        print("- Pagination assertions (field updates, modification time)")
        print("- Rollback functionality (restore original on failure)")
        print("\nAll validation methods return detailed error messages")
        print("and support the complete vNext pipeline workflow.")
        
    except Exception as e:
        logger.error(f"Demo failed: {e}")
        raise


if __name__ == "__main__":
    main()