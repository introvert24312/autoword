#!/usr/bin/env python3
"""
Enhanced Advanced Validation Demo for AutoWord vNext

This demo showcases the comprehensive advanced validation and quality assurance
capabilities implemented in task 18, including:

1. Enhanced document integrity verification beyond basic assertions
2. Comprehensive style consistency checking across entire document  
3. Advanced cross-reference validation and repair capabilities
4. Detailed document accessibility compliance checking
5. Enhanced formatting quality metrics and reporting

Usage:
    python enhanced_advanced_validation_demo.py [document.docx]
"""

import sys
import os
import logging
from pathlib import Path

# Add the vnext directory to Python path
sys.path.insert(0, str(Path(__file__).parent.parent))

from validator.advanced_validator import AdvancedValidator, QualityMetrics
from extractor.document_extractor import DocumentExtractor
from models import StructureV1


def setup_logging():
    """Set up logging for the demo."""
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )


def demonstrate_enhanced_validation(docx_path: str):
    """
    Demonstrate enhanced advanced validation capabilities.
    
    Args:
        docx_path: Path to DOCX file to validate
    """
    print("=" * 80)
    print("Enhanced Advanced Validation Demo")
    print("=" * 80)
    print(f"Analyzing document: {docx_path}")
    print()
    
    try:
        # Step 1: Extract document structure
        print("Step 1: Extracting document structure...")
        with DocumentExtractor(visible=False) as extractor:
            structure = extractor.extract_structure(docx_path)
        
        print(f"✓ Extracted structure with {len(structure.paragraphs)} paragraphs, "
              f"{len(structure.headings)} headings, {len(structure.styles)} styles")
        print()
        
        # Step 2: Perform enhanced advanced validation
        print("Step 2: Performing enhanced advanced validation...")
        with AdvancedValidator(visible=False) as validator:
            
            # 2.1 Enhanced document integrity verification
            print("  2.1 Enhanced Document Integrity Verification...")
            integrity_result = validator.validate_document_integrity(structure, docx_path)
            print(f"      ✓ Integrity check: {'PASS' if integrity_result.is_valid else 'FAIL'}")
            if integrity_result.errors:
                print(f"      ⚠ Found {len(integrity_result.errors)} integrity issues:")
                for error in integrity_result.errors[:3]:
                    print(f"        - {error}")
                if len(integrity_result.errors) > 3:
                    print(f"        ... and {len(integrity_result.errors) - 3} more")
            print()
            
            # 2.2 Comprehensive style consistency checking
            print("  2.2 Comprehensive Style Consistency Checking...")
            style_result = validator.check_style_consistency(structure)
            print(f"      ✓ Style consistency: {'PASS' if style_result.is_valid else 'FAIL'}")
            if style_result.errors:
                print(f"      ⚠ Found {len(style_result.errors)} style issues:")
                for error in style_result.errors[:3]:
                    print(f"        - {error}")
                if len(style_result.errors) > 3:
                    print(f"        ... and {len(style_result.errors) - 3} more")
            print()
            
            # 2.3 Advanced cross-reference validation with repair capabilities
            print("  2.3 Advanced Cross-Reference Validation with Repair...")
            xref_result = validator.validate_cross_references(structure, docx_path)
            print(f"      ✓ Cross-references: {'PASS' if xref_result.is_valid else 'FAIL'}")
            if xref_result.errors:
                print(f"      ⚠ Found {len(xref_result.errors)} cross-reference issues:")
                for error in xref_result.errors[:3]:
                    print(f"        - {error}")
                if len(xref_result.errors) > 3:
                    print(f"        ... and {len(xref_result.errors) - 3} more")
            
            if xref_result.warnings:
                print(f"      💡 Repair recommendations:")
                for warning in xref_result.warnings[:3]:
                    print(f"        - {warning}")
                if len(xref_result.warnings) > 3:
                    print(f"        ... and {len(xref_result.warnings) - 3} more")
            print()
            
            # 2.4 Detailed accessibility compliance checking
            print("  2.4 Detailed Accessibility Compliance Checking...")
            accessibility_result = validator.check_accessibility_compliance(structure)
            print(f"      ✓ Accessibility: {'PASS' if accessibility_result.is_valid else 'FAIL'}")
            if accessibility_result.errors:
                print(f"      ⚠ Found {len(accessibility_result.errors)} accessibility issues:")
                for error in accessibility_result.errors[:3]:
                    print(f"        - {error}")
                if len(accessibility_result.errors) > 3:
                    print(f"        ... and {len(accessibility_result.errors) - 3} more")
            
            if accessibility_result.warnings:
                print(f"      💡 Accessibility recommendations:")
                for warning in accessibility_result.warnings[:3]:
                    print(f"        - {warning}")
                if len(accessibility_result.warnings) > 3:
                    print(f"        ... and {len(accessibility_result.warnings) - 3} more")
            print()
            
            # 2.5 Enhanced formatting quality metrics and reporting
            print("  2.5 Enhanced Formatting Quality Metrics and Reporting...")
            quality_metrics = validator.generate_quality_metrics(structure, docx_path)
            print(f"      ✓ Quality analysis complete")
            print()
        
        # Step 3: Display comprehensive quality report
        print("Step 3: Comprehensive Quality Report")
        print("-" * 50)
        print(quality_metrics.get_summary_report())
        print()
        
        # Step 4: Detailed metrics breakdown
        print("Step 4: Detailed Metrics Breakdown")
        print("-" * 50)
        
        print(f"Style Consistency Score: {quality_metrics.style_consistency_score:.1%}")
        if quality_metrics.inconsistent_styles:
            print("  Style Issues:")
            for issue in quality_metrics.inconsistent_styles[:5]:
                print(f"    - {issue}")
            if len(quality_metrics.inconsistent_styles) > 5:
                print(f"    ... and {len(quality_metrics.inconsistent_styles) - 5} more")
        print()
        
        print(f"Cross-Reference Integrity: {quality_metrics.cross_reference_integrity_score:.1%}")
        if quality_metrics.broken_cross_references:
            print("  Cross-Reference Issues:")
            for issue in quality_metrics.broken_cross_references[:5]:
                print(f"    - {issue}")
            if len(quality_metrics.broken_cross_references) > 5:
                print(f"    ... and {len(quality_metrics.broken_cross_references) - 5} more")
        print()
        
        print(f"Accessibility Compliance: {quality_metrics.accessibility_score:.1%}")
        if quality_metrics.accessibility_issues:
            print("  Accessibility Issues:")
            for issue in quality_metrics.accessibility_issues[:5]:
                print(f"    - {issue}")
            if len(quality_metrics.accessibility_issues) > 5:
                print(f"    ... and {len(quality_metrics.accessibility_issues) - 5} more")
        print()
        
        print(f"Formatting Quality: {quality_metrics.formatting_quality_score:.1%}")
        if quality_metrics.formatting_issues:
            print("  Formatting Issues:")
            for issue in quality_metrics.formatting_issues[:5]:
                print(f"    - {issue}")
            if len(quality_metrics.formatting_issues) > 5:
                print(f"    ... and {len(quality_metrics.formatting_issues) - 5} more")
        print()
        
        # Step 5: Improvement recommendations
        if quality_metrics.improvement_recommendations:
            print("Step 5: Quality Improvement Recommendations")
            print("-" * 50)
            for i, recommendation in enumerate(quality_metrics.improvement_recommendations, 1):
                print(f"{i}. {recommendation}")
            print()
        
        # Step 6: Export quality report
        print("Step 6: Exporting Quality Report")
        print("-" * 50)
        
        # Create output directory
        output_dir = Path(docx_path).parent / "quality_reports"
        output_dir.mkdir(exist_ok=True)
        
        # Export detailed metrics as JSON
        import json
        metrics_file = output_dir / f"{Path(docx_path).stem}_quality_metrics.json"
        with open(metrics_file, 'w', encoding='utf-8') as f:
            json.dump(quality_metrics.to_dict(), f, indent=2, ensure_ascii=False)
        
        # Export summary report as text
        report_file = output_dir / f"{Path(docx_path).stem}_quality_report.txt"
        with open(report_file, 'w', encoding='utf-8') as f:
            f.write(quality_metrics.get_summary_report())
            f.write("\n\n")
            f.write("Detailed Analysis\n")
            f.write("=" * 20 + "\n\n")
            
            if quality_metrics.inconsistent_styles:
                f.write("Style Issues:\n")
                for issue in quality_metrics.inconsistent_styles:
                    f.write(f"  - {issue}\n")
                f.write("\n")
            
            if quality_metrics.broken_cross_references:
                f.write("Cross-Reference Issues:\n")
                for issue in quality_metrics.broken_cross_references:
                    f.write(f"  - {issue}\n")
                f.write("\n")
            
            if quality_metrics.accessibility_issues:
                f.write("Accessibility Issues:\n")
                for issue in quality_metrics.accessibility_issues:
                    f.write(f"  - {issue}\n")
                f.write("\n")
            
            if quality_metrics.formatting_issues:
                f.write("Formatting Issues:\n")
                for issue in quality_metrics.formatting_issues:
                    f.write(f"  - {issue}\n")
                f.write("\n")
            
            if quality_metrics.improvement_recommendations:
                f.write("Improvement Recommendations:\n")
                for i, rec in enumerate(quality_metrics.improvement_recommendations, 1):
                    f.write(f"  {i}. {rec}\n")
        
        print(f"✓ Quality metrics exported to: {metrics_file}")
        print(f"✓ Quality report exported to: {report_file}")
        print()
        
        # Final summary
        print("=" * 80)
        print("Enhanced Advanced Validation Complete")
        print("=" * 80)
        print(f"Overall Quality Grade: {quality_metrics.quality_grade}")
        print(f"Overall Quality Score: {quality_metrics.overall_score:.1%}")
        
        total_issues = (len(quality_metrics.inconsistent_styles) + 
                       len(quality_metrics.broken_cross_references) + 
                       len(quality_metrics.accessibility_issues) + 
                       len(quality_metrics.formatting_issues))
        print(f"Total Issues Found: {total_issues}")
        print(f"Recommendations Provided: {len(quality_metrics.improvement_recommendations)}")
        
        if quality_metrics.overall_score >= 0.8:
            print("🎉 Document has high quality!")
        elif quality_metrics.overall_score >= 0.6:
            print("⚠ Document has moderate quality with room for improvement")
        else:
            print("❌ Document needs significant quality improvements")
        
    except Exception as e:
        print(f"❌ Error during enhanced validation: {e}")
        import traceback
        traceback.print_exc()


def main():
    """Main function to run the enhanced advanced validation demo."""
    setup_logging()
    
    # Get document path from command line or use default
    if len(sys.argv) > 1:
        docx_path = sys.argv[1]
    else:
        # Look for test documents
        test_docs = [
            "test_data/scenario_1_normal_paper/input.docx",
            "test_data/scenario_2_no_toc/input.docx",
            "../test_data/scenario_1_normal_paper/input.docx",
            "../../examples/sample_document.docx"
        ]
        
        docx_path = None
        for test_doc in test_docs:
            if os.path.exists(test_doc):
                docx_path = test_doc
                break
        
        if not docx_path:
            print("Usage: python enhanced_advanced_validation_demo.py [document.docx]")
            print("\nNo test document found. Please provide a DOCX file path.")
            return 1
    
    if not os.path.exists(docx_path):
        print(f"Error: Document not found: {docx_path}")
        return 1
    
    demonstrate_enhanced_validation(docx_path)
    return 0


if __name__ == "__main__":
    sys.exit(main())