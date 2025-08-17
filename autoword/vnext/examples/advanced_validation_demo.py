"""
Advanced Validation Demo

This example demonstrates the comprehensive validation and quality assurance
capabilities of the AdvancedValidator, including document integrity verification,
style consistency checking, cross-reference validation, accessibility compliance,
and quality metrics generation.
"""

import logging
import sys
from pathlib import Path

# Add parent directory to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent))

from validator.advanced_validator import AdvancedValidator, QualityMetrics
from extractor.document_extractor import DocumentExtractor
from models import StructureV1

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


def demonstrate_advanced_validation():
    """Demonstrate advanced validation capabilities."""
    
    print("=== AutoWord vNext Advanced Validation Demo ===\n")
    
    # Sample document path (would be a real DOCX file in practice)
    sample_docx = "sample_document.docx"
    
    # Create a sample structure for demonstration
    sample_structure = create_sample_structure()
    
    print("1. Document Integrity Validation")
    print("-" * 40)
    
    with AdvancedValidator(visible=False) as validator:
        # 1. Document integrity validation
        integrity_result = validator.validate_document_integrity(sample_structure, sample_docx)
        
        print(f"Document integrity: {'✓ VALID' if integrity_result.is_valid else '✗ INVALID'}")
        if integrity_result.errors:
            print("Errors found:")
            for error in integrity_result.errors[:3]:  # Show first 3 errors
                print(f"  - {error}")
        if integrity_result.warnings:
            print("Warnings:")
            for warning in integrity_result.warnings[:3]:  # Show first 3 warnings
                print(f"  - {warning}")
        print()
        
        # 2. Style consistency validation
        print("2. Style Consistency Validation")
        print("-" * 40)
        
        style_result = validator.check_style_consistency(sample_structure)
        
        print(f"Style consistency: {'✓ VALID' if style_result.is_valid else '✗ INVALID'}")
        if style_result.errors:
            print("Style issues found:")
            for error in style_result.errors[:3]:
                print(f"  - {error}")
        if style_result.warnings:
            print("Style warnings:")
            for warning in style_result.warnings[:3]:
                print(f"  - {warning}")
        print()
        
        # 3. Cross-reference validation
        print("3. Cross-Reference Validation")
        print("-" * 40)
        
        cross_ref_result = validator.validate_cross_references(sample_structure, sample_docx)
        
        print(f"Cross-references: {'✓ VALID' if cross_ref_result.is_valid else '✗ INVALID'}")
        if cross_ref_result.errors:
            print("Cross-reference issues:")
            for error in cross_ref_result.errors[:3]:
                print(f"  - {error}")
        if cross_ref_result.warnings:
            print("Recommendations:")
            for warning in cross_ref_result.warnings[:3]:
                print(f"  - {warning}")
        print()
        
        # 4. Accessibility compliance
        print("4. Accessibility Compliance")
        print("-" * 40)
        
        accessibility_result = validator.check_accessibility_compliance(sample_structure)
        
        print(f"Accessibility: {'✓ COMPLIANT' if accessibility_result.is_valid else '✗ ISSUES FOUND'}")
        if accessibility_result.errors:
            print("Accessibility issues:")
            for error in accessibility_result.errors[:3]:
                print(f"  - {error}")
        if accessibility_result.warnings:
            print("Accessibility recommendations:")
            for warning in accessibility_result.warnings[:3]:
                print(f"  - {warning}")
        print()
        
        # 5. Quality metrics
        print("5. Quality Metrics")
        print("-" * 40)
        
        metrics = validator.generate_quality_metrics(sample_structure, sample_docx)
        
        print(f"Overall Quality Score: {metrics.overall_score:.2f}/1.00")
        print(f"Style Consistency: {metrics.style_consistency_score:.2f}/1.00")
        print(f"Cross-Reference Integrity: {metrics.cross_reference_integrity_score:.2f}/1.00")
        print(f"Accessibility Score: {metrics.accessibility_score:.2f}/1.00")
        print(f"Formatting Quality: {metrics.formatting_quality_score:.2f}/1.00")
        print()
        
        print("Quality Assessment:")
        if metrics.overall_score >= 0.9:
            print("  ✓ Excellent document quality")
        elif metrics.overall_score >= 0.7:
            print("  ⚠ Good document quality with minor issues")
        elif metrics.overall_score >= 0.5:
            print("  ⚠ Moderate document quality - improvements recommended")
        else:
            print("  ✗ Poor document quality - significant improvements needed")
        
        print()
        print("Detailed Issues Summary:")
        print(f"  - Style inconsistencies: {len(metrics.inconsistent_styles)}")
        print(f"  - Broken cross-references: {len(metrics.broken_cross_references)}")
        print(f"  - Accessibility issues: {len(metrics.accessibility_issues)}")
        print(f"  - Formatting issues: {len(metrics.formatting_issues)}")
        
        if metrics.inconsistent_styles:
            print("\nTop style issues:")
            for issue in metrics.inconsistent_styles[:3]:
                print(f"  - {issue}")
        
        if metrics.accessibility_issues:
            print("\nTop accessibility issues:")
            for issue in metrics.accessibility_issues[:3]:
                print(f"  - {issue}")


def create_sample_structure():
    """Create a sample document structure for demonstration."""
    from models import (
        StructureV1, DocumentMetadata, StyleDefinition, FontSpec, ParagraphSpec,
        ParagraphSkeleton, HeadingReference, FieldReference, TableSkeleton,
        LineSpacingMode, StyleType
    )
    from datetime import datetime
    
    return StructureV1(
        metadata=DocumentMetadata(
            title="Sample Document for Advanced Validation",
            author="AutoWord Demo",
            creation_time=datetime.now(),
            modified_time=datetime.now(),
            page_count=3,
            paragraph_count=15,
            word_count=250
        ),
        styles=[
            StyleDefinition(
                name="Heading 1",
                type=StyleType.PARAGRAPH,
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
            ),
            # Add a style with potential issues for demonstration
            StyleDefinition(
                name="Small Text",
                type=StyleType.PARAGRAPH,
                font=FontSpec(
                    east_asian="宋体",
                    size_pt=8,  # Too small for accessibility
                    bold=False
                ),
                paragraph=ParagraphSpec(
                    line_spacing_mode=LineSpacingMode.MULTIPLE,
                    line_spacing_value=1.0
                )
            )
        ],
        paragraphs=[
            ParagraphSkeleton(
                index=0,
                style_name="Heading 1",
                preview_text="Introduction",
                is_heading=True,
                heading_level=1
            ),
            ParagraphSkeleton(
                index=1,
                style_name="Normal",
                preview_text="This document demonstrates advanced validation capabilities...",
                is_heading=False
            ),
            ParagraphSkeleton(
                index=2,
                style_name="Heading 1",
                preview_text="Methods",
                is_heading=True,
                heading_level=1
            ),
            ParagraphSkeleton(
                index=3,
                style_name="Small Text",  # Uses problematic style
                preview_text="This text is too small for accessibility compliance.",
                is_heading=False
            )
        ],
        headings=[
            HeadingReference(
                paragraph_index=0,
                level=1,
                text="Introduction",
                style_name="Heading 1"
            ),
            HeadingReference(
                paragraph_index=2,
                level=1,
                text="Methods",
                style_name="Heading 1"
            )
        ],
        fields=[
            FieldReference(
                paragraph_index=4,
                field_type="TOC",
                field_code="TOC \\o \"1-3\" \\h \\z \\u",
                result_text="Introduction\t1\nMethods\t2"
            )
        ],
        tables=[
            TableSkeleton(
                paragraph_index=5,
                rows=3,
                columns=2,
                has_header=False  # Missing header for accessibility
            )
        ]
    )


def demonstrate_quality_metrics_analysis():
    """Demonstrate detailed quality metrics analysis."""
    
    print("\n=== Quality Metrics Analysis ===\n")
    
    sample_structure = create_sample_structure()
    
    with AdvancedValidator(visible=False) as validator:
        metrics = validator.generate_quality_metrics(sample_structure, "sample.docx")
        
        print("Detailed Quality Analysis:")
        print(f"Total styles analyzed: {metrics.total_styles_checked}")
        print(f"Total cross-references checked: {metrics.total_cross_references_checked}")
        print(f"Total accessibility checks: {metrics.total_accessibility_checks}")
        print(f"Total formatting checks: {metrics.total_formatting_checks}")
        print()
        
        # Score breakdown
        print("Score Breakdown:")
        print(f"  Style Consistency: {metrics.style_consistency_score:.2f}")
        print(f"    Weight: 30% of overall score")
        print(f"    Contribution: {metrics.style_consistency_score * 0.3:.2f}")
        print()
        
        print(f"  Cross-Reference Integrity: {metrics.cross_reference_integrity_score:.2f}")
        print(f"    Weight: 20% of overall score")
        print(f"    Contribution: {metrics.cross_reference_integrity_score * 0.2:.2f}")
        print()
        
        print(f"  Accessibility: {metrics.accessibility_score:.2f}")
        print(f"    Weight: 25% of overall score")
        print(f"    Contribution: {metrics.accessibility_score * 0.25:.2f}")
        print()
        
        print(f"  Formatting Quality: {metrics.formatting_quality_score:.2f}")
        print(f"    Weight: 25% of overall score")
        print(f"    Contribution: {metrics.formatting_quality_score * 0.25:.2f}")
        print()
        
        print(f"Overall Score: {metrics.overall_score:.2f}")


if __name__ == "__main__":
    try:
        demonstrate_advanced_validation()
        demonstrate_quality_metrics_analysis()
        
        print("\n=== Demo Complete ===")
        print("The AdvancedValidator provides comprehensive document quality assurance")
        print("beyond basic validation, helping ensure professional document standards.")
        
    except Exception as e:
        logger.error(f"Demo failed: {e}")
        print(f"Error running demo: {e}")
        print("Note: This demo requires the full AutoWord vNext environment.")