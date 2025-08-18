#!/usr/bin/env python3
"""
Test suite for enhanced advanced validation capabilities (Task 18).

This test suite verifies the implementation of:
1. Enhanced document integrity verification beyond basic assertions
2. Comprehensive style consistency checking across entire document
3. Advanced cross-reference validation and repair capabilities  
4. Detailed document accessibility compliance checking
5. Enhanced formatting quality metrics and reporting
"""

import unittest
import sys
import os
from pathlib import Path
from unittest.mock import Mock, patch

# Add the vnext directory to Python path
sys.path.insert(0, str(Path(__file__).parent.parent))

from validator.advanced_validator import AdvancedValidator, QualityMetrics
from models import (
    StructureV1, DocumentMetadata, StyleDefinition, FontSpec, ParagraphSpec,
    ParagraphSkeleton, HeadingReference, FieldReference, TableSkeleton,
    StyleType, LineSpacingMode
)


class TestEnhancedAdvancedValidation(unittest.TestCase):
    """Test enhanced advanced validation capabilities."""
    
    def setUp(self):
        """Set up test fixtures."""
        self.validator = AdvancedValidator(visible=False)
        
        # Create sample document structure for testing
        self.sample_structure = StructureV1(
            metadata=DocumentMetadata(
                title="Test Document",
                author="Test Author",
                page_count=5,
                word_count=1000,
                paragraph_count=50
            ),
            styles=[
                StyleDefinition(
                    name="Heading 1",
                    type=StyleType.PARAGRAPH,
                    font=FontSpec(east_asian="楷体", latin="Times New Roman", size_pt=12, bold=True),
                    paragraph=ParagraphSpec(line_spacing_mode=LineSpacingMode.MULTIPLE, line_spacing_value=2.0)
                ),
                StyleDefinition(
                    name="Heading 2", 
                    type=StyleType.PARAGRAPH,
                    font=FontSpec(east_asian="宋体", latin="Times New Roman", size_pt=12, bold=True),
                    paragraph=ParagraphSpec(line_spacing_mode=LineSpacingMode.MULTIPLE, line_spacing_value=2.0)
                ),
                StyleDefinition(
                    name="Normal",
                    type=StyleType.PARAGRAPH,
                    font=FontSpec(east_asian="宋体", latin="Times New Roman", size_pt=12, bold=False),
                    paragraph=ParagraphSpec(line_spacing_mode=LineSpacingMode.MULTIPLE, line_spacing_value=2.0)
                )
            ],
            paragraphs=[
                ParagraphSkeleton(index=0, style_name="Heading 1", preview_text="Introduction", is_heading=True, heading_level=1),
                ParagraphSkeleton(index=1, style_name="Normal", preview_text="This is the introduction paragraph..."),
                ParagraphSkeleton(index=2, style_name="Heading 2", preview_text="Background", is_heading=True, heading_level=2),
                ParagraphSkeleton(index=3, style_name="Normal", preview_text="Background information..."),
                ParagraphSkeleton(index=4, style_name="Heading 1", preview_text="Conclusion", is_heading=True, heading_level=1),
            ],
            headings=[
                HeadingReference(paragraph_index=0, level=1, text="Introduction", style_name="Heading 1"),
                HeadingReference(paragraph_index=2, level=2, text="Background", style_name="Heading 2"),
                HeadingReference(paragraph_index=4, level=1, text="Conclusion", style_name="Heading 1"),
            ],
            fields=[
                FieldReference(paragraph_index=1, field_type="TOC", field_code="TOC \\o \"1-3\"", result_text="Introduction\t1\nBackground\t1\nConclusion\t2"),
                FieldReference(paragraph_index=3, field_type="REF", field_code="REF Introduction", result_text="Introduction"),
            ],
            tables=[
                TableSkeleton(paragraph_index=3, rows=3, columns=2, has_header=True)
            ]
        )
    
    def test_enhanced_document_integrity_verification(self):
        """Test enhanced document integrity verification beyond basic assertions."""
        # Test with valid structure
        result = self.validator._validate_structure_integrity(self.sample_structure)
        self.assertEqual(len(result), 0, "Valid structure should have no integrity errors")
        
        # Test with missing metadata
        invalid_structure = StructureV1(
            metadata=None,
            styles=[],
            paragraphs=[],
            headings=[],
            fields=[],
            tables=[]
        )
        result = self.validator._validate_structure_integrity(invalid_structure)
        self.assertGreater(len(result), 0, "Missing metadata should trigger integrity error")
        self.assertTrue(any("Missing metadata" in error for error in result))
        
        # Test with invalid paragraph indexing
        invalid_structure = StructureV1(
            metadata=DocumentMetadata(title="Test"),
            styles=[],
            paragraphs=[
                ParagraphSkeleton(index=0, preview_text="First"),
                ParagraphSkeleton(index=2, preview_text="Third"),  # Missing index 1
            ],
            headings=[],
            fields=[],
            tables=[]
        )
        result = self.validator._validate_structure_integrity(invalid_structure)
        self.assertTrue(any("Missing paragraph indexes" in error for error in result))
        
        # Test with orphaned style references
        invalid_structure = StructureV1(
            metadata=DocumentMetadata(title="Test"),
            styles=[
                StyleDefinition(name="Style1", type=StyleType.PARAGRAPH, font=FontSpec(size_pt=12))
            ],
            paragraphs=[
                ParagraphSkeleton(index=0, style_name="NonExistentStyle", preview_text="Test")
            ],
            headings=[],
            fields=[],
            tables=[]
        )
        result = self.validator._validate_structure_integrity(invalid_structure)
        self.assertTrue(any("Orphaned style references" in error for error in result))
    
    def test_comprehensive_style_consistency_checking(self):
        """Test comprehensive style consistency checking across entire document."""
        # Test with consistent styles
        result = self.validator._check_style_definitions(self.sample_structure)
        self.assertEqual(len(result), 0, "Valid styles should have no consistency errors")
        
        # Test with duplicate style names
        invalid_structure = StructureV1(
            metadata=DocumentMetadata(title="Test"),
            styles=[
                StyleDefinition(name="Heading 1", type=StyleType.PARAGRAPH, font=FontSpec(size_pt=12)),
                StyleDefinition(name="Heading 1", type=StyleType.PARAGRAPH, font=FontSpec(size_pt=14)),  # Duplicate
            ],
            paragraphs=[],
            headings=[],
            fields=[],
            tables=[]
        )
        result = self.validator._check_style_definitions(invalid_structure)
        self.assertTrue(any("Duplicate style names" in error for error in result))
        
        # Test with incomplete style definitions
        invalid_structure = StructureV1(
            metadata=DocumentMetadata(title="Test"),
            styles=[
                StyleDefinition(name="Incomplete", type=StyleType.PARAGRAPH, font=None),  # Missing font
            ],
            paragraphs=[],
            headings=[],
            fields=[],
            tables=[]
        )
        result = self.validator._check_style_definitions(invalid_structure)
        self.assertTrue(any("missing font specification" in error for error in result))
        
        # Test enhanced style consistency scoring
        score = self.validator._calculate_enhanced_style_consistency_score(self.sample_structure)
        self.assertGreaterEqual(score, 0.0)
        self.assertLessEqual(score, 1.0)
        self.assertGreater(score, 0.8, "Well-formed structure should have high style consistency score")
    
    def test_advanced_cross_reference_validation_and_repair(self):
        """Test advanced cross-reference validation and repair capabilities."""
        # Test with valid cross-references
        result = self.validator._validate_field_references(self.sample_structure)
        self.assertEqual(len(result), 0, "Valid cross-references should have no errors")
        
        # Test with broken cross-references
        invalid_structure = StructureV1(
            metadata=DocumentMetadata(title="Test"),
            styles=[],
            paragraphs=[ParagraphSkeleton(index=0, preview_text="Test")],
            headings=[],
            fields=[
                FieldReference(paragraph_index=0, field_type="REF", field_code="REF NonExistent", result_text="Error! Reference source not found."),
            ],
            tables=[]
        )
        result = self.validator._validate_field_references(invalid_structure)
        self.assertTrue(any("Field error" in error for error in result))
        
        # Test repair recommendations
        repair_recommendations = self.validator._generate_cross_reference_repair_recommendations(invalid_structure)
        self.assertGreater(len(repair_recommendations), 0, "Should provide repair recommendations")
        self.assertTrue(any("Update all fields" in rec for rec in repair_recommendations))
        
        # Test enhanced cross-reference scoring
        score = self.validator._calculate_enhanced_cross_reference_score(self.sample_structure)
        self.assertGreaterEqual(score, 0.0)
        self.assertLessEqual(score, 1.0)
    
    def test_detailed_accessibility_compliance_checking(self):
        """Test detailed document accessibility compliance checking."""
        # Test with accessible structure
        result = self.validator._check_heading_accessibility(self.sample_structure)
        self.assertEqual(len(result), 0, "Well-structured headings should have no accessibility errors")
        
        # Test with no headings
        invalid_structure = StructureV1(
            metadata=DocumentMetadata(title="Test"),
            styles=[],
            paragraphs=[ParagraphSkeleton(index=0, preview_text="Test")],
            headings=[],
            fields=[],
            tables=[]
        )
        result = self.validator._check_heading_accessibility(invalid_structure)
        self.assertTrue(any("no headings for navigation" in error for error in result))
        
        # Test with skipped heading levels
        invalid_structure = StructureV1(
            metadata=DocumentMetadata(title="Test"),
            styles=[],
            paragraphs=[],
            headings=[
                HeadingReference(paragraph_index=0, level=1, text="Title"),
                HeadingReference(paragraph_index=1, level=3, text="Subsection"),  # Skips level 2
            ],
            fields=[],
            tables=[]
        )
        result = self.validator._check_heading_accessibility(invalid_structure)
        self.assertTrue(any("Heading level skipped" in error for error in result))
        
        # Test with short heading text
        invalid_structure = StructureV1(
            metadata=DocumentMetadata(title="Test"),
            styles=[],
            paragraphs=[],
            headings=[
                HeadingReference(paragraph_index=0, level=1, text="A"),  # Too short
            ],
            fields=[],
            tables=[]
        )
        result = self.validator._check_heading_accessibility(invalid_structure)
        self.assertTrue(any("too short for screen readers" in error for error in result))
        
        # Test table accessibility
        result = self.validator._check_table_accessibility(self.sample_structure)
        self.assertEqual(len(result), 0, "Table with header should have no accessibility errors")
        
        # Test table without header
        invalid_structure = StructureV1(
            metadata=DocumentMetadata(title="Test"),
            styles=[],
            paragraphs=[],
            headings=[],
            fields=[],
            tables=[
                TableSkeleton(paragraph_index=0, rows=2, columns=2, has_header=False)
            ]
        )
        result = self.validator._check_table_accessibility(invalid_structure)
        self.assertTrue(any("should have header row" in error for error in result))
        
        # Test enhanced accessibility scoring
        score = self.validator._calculate_enhanced_accessibility_score(self.sample_structure)
        self.assertGreaterEqual(score, 0.0)
        self.assertLessEqual(score, 1.0)
    
    def test_enhanced_formatting_quality_metrics_and_reporting(self):
        """Test enhanced formatting quality metrics and reporting."""
        # Test quality metrics generation
        with patch.object(self.validator, '_word_app', None):  # Skip Word COM operations
            metrics = self.validator.generate_quality_metrics(self.sample_structure, "test.docx")
        
        # Verify all scores are within valid range
        self.assertGreaterEqual(metrics.style_consistency_score, 0.0)
        self.assertLessEqual(metrics.style_consistency_score, 1.0)
        self.assertGreaterEqual(metrics.cross_reference_integrity_score, 0.0)
        self.assertLessEqual(metrics.cross_reference_integrity_score, 1.0)
        self.assertGreaterEqual(metrics.accessibility_score, 0.0)
        self.assertLessEqual(metrics.accessibility_score, 1.0)
        self.assertGreaterEqual(metrics.formatting_quality_score, 0.0)
        self.assertLessEqual(metrics.formatting_quality_score, 1.0)
        self.assertGreaterEqual(metrics.overall_score, 0.0)
        self.assertLessEqual(metrics.overall_score, 1.0)
        
        # Verify quality grade is assigned
        self.assertIn(metrics.quality_grade, ["A+", "A", "A-", "B+", "B", "B-", "C+", "C", "C-", "D", "F"])
        
        # Verify metrics can be serialized
        metrics_dict = metrics.to_dict()
        self.assertIn('scores', metrics_dict)
        self.assertIn('grade', metrics_dict)
        self.assertIn('issues', metrics_dict)
        self.assertIn('recommendations', metrics_dict)
        
        # Verify summary report generation
        summary = metrics.get_summary_report()
        self.assertIn("Document Quality Assessment", summary)
        self.assertIn("Overall Grade:", summary)
        self.assertIn("Detailed Scores:", summary)
        
        # Test with problematic document
        problematic_structure = StructureV1(
            metadata=DocumentMetadata(title="Test"),
            styles=[
                StyleDefinition(name="Bad Style", type=StyleType.PARAGRAPH, font=None),  # Missing font
            ],
            paragraphs=[
                ParagraphSkeleton(index=0, style_name="NonExistent", preview_text="Test")  # Orphaned style
            ],
            headings=[],  # No headings
            fields=[
                FieldReference(paragraph_index=0, field_type="REF", result_text="Error!")  # Broken field
            ],
            tables=[
                TableSkeleton(paragraph_index=0, rows=2, columns=2, has_header=False)  # No header
            ]
        )
        
        with patch.object(self.validator, '_word_app', None):
            problematic_metrics = self.validator.generate_quality_metrics(problematic_structure, "test.docx")
        
        # Should have lower scores due to issues
        self.assertLess(problematic_metrics.overall_score, 0.8)
        self.assertGreater(len(problematic_metrics.inconsistent_styles), 0)
        self.assertGreater(len(problematic_metrics.accessibility_issues), 0)
        self.assertGreater(len(problematic_metrics.improvement_recommendations), 0)
    
    def test_quality_grade_calculation(self):
        """Test quality grade calculation."""
        test_cases = [
            (0.95, "A+"),
            (0.90, "A"),
            (0.85, "A-"),
            (0.80, "B+"),
            (0.75, "B"),
            (0.70, "B-"),
            (0.65, "C+"),
            (0.60, "C"),
            (0.55, "C-"),
            (0.50, "D"),
            (0.40, "F"),
        ]
        
        for score, expected_grade in test_cases:
            grade = self.validator._calculate_quality_grade(score)
            self.assertEqual(grade, expected_grade, f"Score {score} should map to grade {expected_grade}")
    
    def test_enhanced_validation_integration(self):
        """Test integration of all enhanced validation components."""
        # Mock Word COM to avoid requiring actual Word installation
        with patch.object(self.validator, '_word_app', None):
            with patch.object(self.validator, '_com_initialized', False):
                
                # Test document integrity validation
                integrity_result = self.validator.validate_document_integrity(self.sample_structure, "test.docx")
                self.assertIsInstance(integrity_result.is_valid, bool)
                
                # Test style consistency validation
                style_result = self.validator.check_style_consistency(self.sample_structure)
                self.assertIsInstance(style_result.is_valid, bool)
                
                # Test cross-reference validation
                xref_result = self.validator.validate_cross_references(self.sample_structure, "test.docx")
                self.assertIsInstance(xref_result.is_valid, bool)
                
                # Test accessibility validation
                accessibility_result = self.validator.check_accessibility_compliance(self.sample_structure)
                self.assertIsInstance(accessibility_result.is_valid, bool)
                
                # Test quality metrics generation
                quality_metrics = self.validator.generate_quality_metrics(self.sample_structure, "test.docx")
                self.assertIsInstance(quality_metrics, QualityMetrics)
                self.assertIsInstance(quality_metrics.overall_score, float)
                self.assertIsInstance(quality_metrics.quality_grade, str)


class TestQualityMetrics(unittest.TestCase):
    """Test QualityMetrics class functionality."""
    
    def test_quality_metrics_initialization(self):
        """Test QualityMetrics initialization."""
        metrics = QualityMetrics()
        
        # Check default values
        self.assertEqual(metrics.style_consistency_score, 0.0)
        self.assertEqual(metrics.cross_reference_integrity_score, 0.0)
        self.assertEqual(metrics.accessibility_score, 0.0)
        self.assertEqual(metrics.formatting_quality_score, 0.0)
        self.assertEqual(metrics.overall_score, 0.0)
        self.assertEqual(metrics.quality_grade, "F")
        
        # Check lists are initialized
        self.assertIsInstance(metrics.inconsistent_styles, list)
        self.assertIsInstance(metrics.broken_cross_references, list)
        self.assertIsInstance(metrics.accessibility_issues, list)
        self.assertIsInstance(metrics.formatting_issues, list)
        self.assertIsInstance(metrics.improvement_recommendations, list)
    
    def test_quality_metrics_serialization(self):
        """Test QualityMetrics serialization to dictionary."""
        metrics = QualityMetrics()
        metrics.style_consistency_score = 0.8
        metrics.overall_score = 0.75
        metrics.quality_grade = "B"
        metrics.inconsistent_styles = ["Style issue 1", "Style issue 2"]
        metrics.improvement_recommendations = ["Recommendation 1"]
        
        metrics_dict = metrics.to_dict()
        
        # Check structure
        self.assertIn('scores', metrics_dict)
        self.assertIn('grade', metrics_dict)
        self.assertIn('issues', metrics_dict)
        self.assertIn('recommendations', metrics_dict)
        self.assertIn('statistics', metrics_dict)
        
        # Check values
        self.assertEqual(metrics_dict['scores']['style_consistency'], 0.8)
        self.assertEqual(metrics_dict['scores']['overall'], 0.75)
        self.assertEqual(metrics_dict['grade'], "B")
        self.assertEqual(len(metrics_dict['issues']['inconsistent_styles']), 2)
        self.assertEqual(len(metrics_dict['recommendations']), 1)
    
    def test_quality_metrics_summary_report(self):
        """Test QualityMetrics summary report generation."""
        metrics = QualityMetrics()
        metrics.overall_score = 0.85
        metrics.quality_grade = "A-"
        metrics.style_consistency_score = 0.9
        metrics.cross_reference_integrity_score = 0.8
        metrics.accessibility_score = 0.85
        metrics.formatting_quality_score = 0.85
        metrics.improvement_recommendations = ["Improve cross-references", "Add more headings"]
        
        summary = metrics.get_summary_report()
        
        # Check content
        self.assertIn("Document Quality Assessment", summary)
        self.assertIn("Overall Grade: A-", summary)
        self.assertIn("85.0%", summary)  # Overall score
        self.assertIn("Style Consistency: 90.0%", summary)
        self.assertIn("Cross-Reference Integrity: 80.0%", summary)
        self.assertIn("Accessibility Compliance: 85.0%", summary)
        self.assertIn("Formatting Quality: 85.0%", summary)
        self.assertIn("Top Recommendations:", summary)
        self.assertIn("Improve cross-references", summary)


if __name__ == '__main__':
    # Set up test environment
    import logging
    logging.basicConfig(level=logging.WARNING)  # Reduce log noise during tests
    
    # Run tests
    unittest.main(verbosity=2)