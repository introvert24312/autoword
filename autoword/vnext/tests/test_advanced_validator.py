"""
Tests for advanced validation and quality assurance functionality.

This module tests the AdvancedValidator class and its comprehensive
document validation capabilities including integrity checks, style
consistency, cross-reference validation, accessibility compliance,
and quality metrics generation.
"""

import pytest
from unittest.mock import Mock, patch, MagicMock
from datetime import datetime
from pathlib import Path

from autoword.vnext.validator.advanced_validator import AdvancedValidator, QualityMetrics
from autoword.vnext.models import (
    StructureV1, DocumentMetadata, StyleDefinition, FontSpec, ParagraphSpec,
    ParagraphSkeleton, HeadingReference, FieldReference, TableSkeleton,
    ValidationResult, LineSpacingMode, StyleType
)
from autoword.vnext.exceptions import ValidationError


class TestAdvancedValidator:
    """Test cases for AdvancedValidator."""
    
    @pytest.fixture
    def sample_structure(self):
        """Create a sample document structure for testing."""
        return StructureV1(
            metadata=DocumentMetadata(
                title="Test Document",
                author="Test Author",
                creation_time=datetime.now(),
                modified_time=datetime.now(),
                page_count=5,
                paragraph_count=20,
                word_count=500
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
                    preview_text="This is the first paragraph of content...",
                    is_heading=False
                ),
                ParagraphSkeleton(
                    index=2,
                    style_name="Heading 1",
                    preview_text="Conclusion",
                    is_heading=True,
                    heading_level=1
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
                    text="Conclusion",
                    style_name="Heading 1"
                )
            ],
            fields=[
                FieldReference(
                    paragraph_index=3,
                    field_type="TOC",
                    field_code="TOC \\o \"1-3\" \\h \\z \\u",
                    result_text="Introduction\t1\nConclusion\t2"
                )
            ],
            tables=[
                TableSkeleton(
                    paragraph_index=4,
                    rows=3,
                    columns=2,
                    has_header=True
                )
            ]
        )
    
    @pytest.fixture
    def mock_word_app(self):
        """Create a mock Word application for testing."""
        mock_app = Mock()
        mock_doc = Mock()
        mock_app.Documents.Open.return_value = mock_doc
        
        # Mock document properties
        mock_doc.SpellingErrors.Count = 0
        mock_doc.GrammarErrors.Count = 0
        mock_doc.Hyperlinks = []
        mock_doc.Fields = []
        mock_doc.ProtectionType = 0
        mock_doc.TrackRevisions = False
        mock_doc.Comments.Count = 0
        
        return mock_app
    
    def test_init(self):
        """Test AdvancedValidator initialization."""
        validator = AdvancedValidator(visible=True)
        assert validator.visible is True
        assert validator._word_app is None
        assert validator._com_initialized is False
        assert validator.min_heading_hierarchy_score == 0.8
        assert validator.max_style_variations_per_type == 3
    
    @patch('autoword.vnext.validator.advanced_validator.pythoncom')
    @patch('autoword.vnext.validator.advanced_validator.win32')
    def test_context_manager(self, mock_win32, mock_pythoncom):
        """Test context manager functionality."""
        mock_app = Mock()
        mock_win32.gencache.EnsureDispatch.return_value = mock_app
        
        with AdvancedValidator() as validator:
            assert validator._word_app is mock_app
            assert validator._com_initialized is True
            mock_pythoncom.CoInitialize.assert_called_once()
        
        mock_app.Quit.assert_called_once()
        mock_pythoncom.CoUninitialize.assert_called_once()
    
    def test_validate_document_integrity_basic(self, sample_structure):
        """Test basic document integrity validation."""
        validator = AdvancedValidator()
        
        result = validator.validate_document_integrity(sample_structure, "test.docx")
        
        assert isinstance(result, ValidationResult)
        # Should pass basic integrity checks with sample structure
        assert result.is_valid is True
        assert len(result.errors) == 0
    
    def test_validate_document_integrity_missing_metadata(self, sample_structure):
        """Test document integrity validation with missing metadata."""
        validator = AdvancedValidator()
        
        # Create structure without metadata (using model_copy to bypass validation)
        structure_dict = sample_structure.model_dump()
        structure_dict['metadata'] = None
        
        # Create a new structure that bypasses Pydantic validation for this test
        from autoword.vnext.models import StructureV1
        import json
        
        # Manually create structure with None metadata for testing
        test_structure = StructureV1.model_validate({
            **structure_dict,
            'metadata': {
                'title': None,
                'author': None,
                'creation_time': None,
                'modified_time': None,
                'page_count': None,
                'paragraph_count': None,
                'word_count': None
            }
        })
        
        # Now set metadata to None after creation (for testing purposes)
        object.__setattr__(test_structure, 'metadata', None)
        
        result = validator.validate_document_integrity(test_structure, "test.docx")
        
        assert result.is_valid is False
        assert any("Missing metadata" in error for error in result.errors)
    
    def test_validate_document_integrity_invalid_paragraph_indexes(self, sample_structure):
        """Test document integrity validation with invalid paragraph indexes."""
        validator = AdvancedValidator()
        # Create non-sequential paragraph indexes
        sample_structure.paragraphs[1].index = 5
        
        result = validator.validate_document_integrity(sample_structure, "test.docx")
        
        assert result.is_valid is False
        assert any("not sequential" in error for error in result.errors)
    
    def test_validate_document_integrity_invalid_heading_reference(self, sample_structure):
        """Test document integrity validation with invalid heading reference."""
        validator = AdvancedValidator()
        # Add heading that references non-existent paragraph
        sample_structure.headings.append(
            HeadingReference(
                paragraph_index=999,
                level=1,
                text="Invalid Heading",
                style_name="Heading 1"
            )
        )
        
        result = validator.validate_document_integrity(sample_structure, "test.docx")
        
        assert result.is_valid is False
        assert any("invalid paragraph index" in error for error in result.errors)
    
    def test_check_style_consistency_basic(self, sample_structure):
        """Test basic style consistency checking."""
        validator = AdvancedValidator()
        
        result = validator.check_style_consistency(sample_structure)
        
        assert isinstance(result, ValidationResult)
        assert result.is_valid is True
    
    def test_check_style_consistency_duplicate_styles(self, sample_structure):
        """Test style consistency with duplicate style names."""
        validator = AdvancedValidator()
        # Add duplicate style
        sample_structure.styles.append(
            StyleDefinition(
                name="Heading 1",  # Duplicate name
                type=StyleType.PARAGRAPH,
                font=FontSpec(size_pt=14),
                paragraph=ParagraphSpec(line_spacing_value=1.5)
            )
        )
        
        result = validator.check_style_consistency(sample_structure)
        
        assert result.is_valid is False
        assert any("Duplicate style names" in error for error in result.errors)
    
    def test_check_style_consistency_undefined_style_usage(self, sample_structure):
        """Test style consistency with undefined style usage."""
        validator = AdvancedValidator()
        # Use undefined style in paragraph
        sample_structure.paragraphs[0].style_name = "Undefined Style"
        
        result = validator.check_style_consistency(sample_structure)
        
        assert result.is_valid is False
        assert any("Undefined styles in use" in error for error in result.errors)
    
    def test_check_style_consistency_incomplete_style_definition(self, sample_structure):
        """Test style consistency with incomplete style definitions."""
        validator = AdvancedValidator()
        # Remove font specification from style
        sample_structure.styles[0].font = None
        
        result = validator.check_style_consistency(sample_structure)
        
        assert result.is_valid is False
        assert any("missing font specification" in error for error in result.errors)
    
    def test_validate_cross_references_basic(self, sample_structure):
        """Test basic cross-reference validation."""
        validator = AdvancedValidator()
        
        result = validator.validate_cross_references(sample_structure, "test.docx")
        
        assert isinstance(result, ValidationResult)
        assert result.is_valid is True
    
    def test_validate_cross_references_broken_field(self, sample_structure):
        """Test cross-reference validation with broken field."""
        validator = AdvancedValidator()
        # Add broken field reference
        sample_structure.fields.append(
            FieldReference(
                paragraph_index=5,
                field_type="REF",
                field_code="REF NonExistentHeading",
                result_text="Error! Reference source not found."
            )
        )
        
        result = validator.validate_cross_references(sample_structure, "test.docx")
        
        assert result.is_valid is False
        assert any("Field error" in error for error in result.errors)
    
    def test_validate_cross_references_empty_toc(self, sample_structure):
        """Test cross-reference validation with empty TOC."""
        validator = AdvancedValidator()
        # Make TOC field empty
        sample_structure.fields[0].result_text = ""
        
        result = validator.validate_cross_references(sample_structure, "test.docx")
        
        assert result.is_valid is False
        assert any("Empty TOC" in error for error in result.errors)
    
    def test_check_accessibility_compliance_basic(self, sample_structure):
        """Test basic accessibility compliance checking."""
        validator = AdvancedValidator()
        
        result = validator.check_accessibility_compliance(sample_structure)
        
        assert isinstance(result, ValidationResult)
        assert result.is_valid is True
    
    def test_check_accessibility_compliance_no_headings(self, sample_structure):
        """Test accessibility compliance with no headings."""
        validator = AdvancedValidator()
        sample_structure.headings = []
        
        result = validator.check_accessibility_compliance(sample_structure)
        
        assert result.is_valid is False
        assert any("no headings for navigation" in error for error in result.errors)
    
    def test_check_accessibility_compliance_skipped_heading_levels(self, sample_structure):
        """Test accessibility compliance with skipped heading levels."""
        validator = AdvancedValidator()
        # Add heading that skips level 2
        sample_structure.headings.append(
            HeadingReference(
                paragraph_index=3,
                level=3,  # Skips level 2
                text="Subsection",
                style_name="Heading 3"
            )
        )
        
        result = validator.check_accessibility_compliance(sample_structure)
        
        assert result.is_valid is False
        assert any("Heading level skipped" in error for error in result.errors)
    
    def test_check_accessibility_compliance_small_font(self, sample_structure):
        """Test accessibility compliance with small font size."""
        validator = AdvancedValidator()
        # Set font size too small
        sample_structure.styles[0].font.size_pt = 8
        
        result = validator.check_accessibility_compliance(sample_structure)
        
        assert result.is_valid is False
        assert any("font size" in error and "too small" in error for error in result.errors)
    
    def test_check_accessibility_compliance_table_without_header(self, sample_structure):
        """Test accessibility compliance with table without header."""
        validator = AdvancedValidator()
        sample_structure.tables[0].has_header = False
        
        result = validator.check_accessibility_compliance(sample_structure)
        
        assert result.is_valid is False
        assert any("should have header row" in error for error in result.errors)
    
    def test_generate_quality_metrics_basic(self, sample_structure):
        """Test basic quality metrics generation."""
        validator = AdvancedValidator()
        
        metrics = validator.generate_quality_metrics(sample_structure, "test.docx")
        
        assert isinstance(metrics, QualityMetrics)
        assert 0.0 <= metrics.overall_score <= 1.0
        assert 0.0 <= metrics.style_consistency_score <= 1.0
        assert 0.0 <= metrics.cross_reference_integrity_score <= 1.0
        assert 0.0 <= metrics.accessibility_score <= 1.0
        assert 0.0 <= metrics.formatting_quality_score <= 1.0
        assert metrics.total_styles_checked == len(sample_structure.styles)
    
    def test_generate_quality_metrics_with_issues(self, sample_structure):
        """Test quality metrics generation with various issues."""
        validator = AdvancedValidator()
        
        # Add issues to reduce scores
        sample_structure.styles[0].font = None  # Incomplete style
        sample_structure.fields[0].result_text = "Error!"  # Broken field
        sample_structure.headings = []  # No headings
        
        metrics = validator.generate_quality_metrics(sample_structure, "test.docx")
        
        assert metrics.overall_score < 1.0
        assert len(metrics.inconsistent_styles) > 0
        assert len(metrics.broken_cross_references) > 0
        assert len(metrics.accessibility_issues) > 0
    
    def test_heading_hierarchy_validation(self, sample_structure):
        """Test heading hierarchy validation."""
        validator = AdvancedValidator()
        
        # Add heading with skipped level
        sample_structure.headings.append(
            HeadingReference(
                paragraph_index=3,
                level=3,  # Skips level 2
                text="Deep Section",
                style_name="Heading 3"
            )
        )
        
        errors = validator._validate_heading_hierarchy(sample_structure)
        
        assert len(errors) > 0
        assert any("skips levels" in error for error in errors)
    
    def test_heading_hierarchy_validation_empty_heading(self, sample_structure):
        """Test heading hierarchy validation with empty heading."""
        validator = AdvancedValidator()
        sample_structure.headings[0].text = ""
        
        errors = validator._validate_heading_hierarchy(sample_structure)
        
        assert len(errors) > 0
        assert any("Empty heading" in error for error in errors)
    
    def test_heading_hierarchy_validation_duplicate_headings(self, sample_structure):
        """Test heading hierarchy validation with duplicate headings."""
        validator = AdvancedValidator()
        # Add duplicate heading at same level
        sample_structure.headings.append(
            HeadingReference(
                paragraph_index=4,
                level=1,
                text="Introduction",  # Duplicate
                style_name="Heading 1"
            )
        )
        
        errors = validator._validate_heading_hierarchy(sample_structure)
        
        assert len(errors) > 0
        assert any("Duplicate headings" in error for error in errors)
    
    def test_table_integrity_validation(self, sample_structure):
        """Test table integrity validation."""
        validator = AdvancedValidator()
        
        # Create a table with invalid dimensions by bypassing Pydantic validation
        from autoword.vnext.models import TableSkeleton
        
        # Create valid table first
        valid_table = TableSkeleton(
            paragraph_index=5,
            rows=1,
            columns=2,
            has_header=False
        )
        
        # Manually set invalid dimensions for testing
        object.__setattr__(valid_table, 'rows', 0)
        
        sample_structure.tables.append(valid_table)
        
        errors = validator._validate_table_integrity(sample_structure)
        
        assert len(errors) > 0
        assert any("invalid dimensions" in error for error in errors)
    
    def test_field_integrity_validation(self, sample_structure):
        """Test field integrity validation."""
        validator = AdvancedValidator()
        
        # Add field with empty field code
        sample_structure.fields.append(
            FieldReference(
                paragraph_index=6,
                field_type="PAGE",
                field_code="",  # Empty
                result_text="1"
            )
        )
        
        errors = validator._validate_field_integrity(sample_structure)
        
        assert len(errors) > 0
        assert any("empty field code" in error for error in errors)
    
    def test_style_consistency_font_variations(self, sample_structure):
        """Test style consistency with excessive font variations."""
        validator = AdvancedValidator()
        validator.max_style_variations_per_type = 2  # Lower threshold
        
        # Add styles with many different font sizes
        for i, size in enumerate([10, 11, 12, 14, 16]):
            sample_structure.styles.append(
                StyleDefinition(
                    name=f"Style{i}",
                    type=StyleType.PARAGRAPH,
                    font=FontSpec(size_pt=size),
                    paragraph=ParagraphSpec(line_spacing_value=2.0)
                )
            )
        
        errors = validator._check_font_consistency(sample_structure)
        
        assert len(errors) > 0
        assert any("Too many font size variations" in error for error in errors)
    
    def test_cross_reference_repair_recommendations(self, sample_structure):
        """Test cross-reference repair recommendations."""
        validator = AdvancedValidator()
        
        warnings = validator._generate_cross_reference_repair_recommendations(sample_structure)
        
        assert len(warnings) > 0
        assert any("Update all fields" in warning for warning in warnings)
    
    def test_accessibility_recommendations(self, sample_structure):
        """Test accessibility improvement recommendations."""
        validator = AdvancedValidator()
        
        warnings = validator._generate_accessibility_recommendations(sample_structure)
        
        assert len(warnings) > 0
        assert any("heading styles" in warning for warning in warnings)
        assert any("color contrast" in warning for warning in warnings)
    
    @patch('autoword.vnext.validator.advanced_validator.os.path.exists')
    def test_validate_document_integrity_with_word_com(self, mock_exists, sample_structure, mock_word_app):
        """Test document integrity validation with Word COM integration."""
        mock_exists.return_value = True
        
        validator = AdvancedValidator()
        validator._word_app = mock_word_app
        
        result = validator.validate_document_integrity(sample_structure, "test.docx")
        
        assert isinstance(result, ValidationResult)
        mock_word_app.Documents.Open.assert_called_once_with("test.docx")
    
    def test_quality_metrics_calculation_methods(self, sample_structure):
        """Test individual quality metrics calculation methods."""
        validator = AdvancedValidator()
        
        # Test style consistency score
        style_score = validator._calculate_style_consistency_score(sample_structure)
        assert 0.0 <= style_score <= 1.0
        
        # Test cross-reference score
        ref_score = validator._calculate_cross_reference_score(sample_structure)
        assert 0.0 <= ref_score <= 1.0
        
        # Test accessibility score
        access_score = validator._calculate_accessibility_score(sample_structure)
        assert 0.0 <= access_score <= 1.0
        
        # Test formatting quality score
        format_score = validator._calculate_formatting_quality_score(sample_structure)
        assert 0.0 <= format_score <= 1.0
    
    def test_issue_collection_methods(self, sample_structure):
        """Test issue collection methods."""
        validator = AdvancedValidator()
        
        # Add some issues
        sample_structure.styles[0].font = None  # Missing font
        sample_structure.fields[0].result_text = "Error!"  # Broken field
        
        inconsistent_styles = validator._collect_inconsistent_styles(sample_structure)
        broken_refs = validator._collect_broken_cross_references(sample_structure)
        accessibility_issues = validator._collect_accessibility_issues(sample_structure)
        formatting_issues = validator._collect_formatting_issues(sample_structure)
        
        assert len(inconsistent_styles) > 0
        assert len(broken_refs) > 0
        assert isinstance(accessibility_issues, list)
        assert isinstance(formatting_issues, list)


class TestQualityMetrics:
    """Test cases for QualityMetrics class."""
    
    def test_quality_metrics_initialization(self):
        """Test QualityMetrics initialization."""
        metrics = QualityMetrics()
        
        assert metrics.style_consistency_score == 0.0
        assert metrics.cross_reference_integrity_score == 0.0
        assert metrics.accessibility_score == 0.0
        assert metrics.formatting_quality_score == 0.0
        assert metrics.overall_score == 0.0
        
        assert isinstance(metrics.inconsistent_styles, list)
        assert isinstance(metrics.broken_cross_references, list)
        assert isinstance(metrics.accessibility_issues, list)
        assert isinstance(metrics.formatting_issues, list)
        
        assert metrics.total_styles_checked == 0
        assert metrics.total_cross_references_checked == 0
        assert metrics.total_accessibility_checks == 0
        assert metrics.total_formatting_checks == 0


if __name__ == "__main__":
    pytest.main([__file__])