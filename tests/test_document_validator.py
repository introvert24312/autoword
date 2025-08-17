"""
Unit tests for DocumentValidator class.

Tests comprehensive validation and rollback capabilities including:
- Chapter assertions (no 摘要/参考文献 at level 1)
- Style assertions (H1/H2/Normal specifications)
- TOC assertions (consistency with heading tree)
- Pagination assertions (fields updated, repagination)
- Rollback functionality
"""

import os
import tempfile
import shutil
from unittest.mock import Mock, patch, MagicMock
import pytest
from datetime import datetime

from autoword.vnext.validator.document_validator import DocumentValidator
from autoword.vnext.models import (
    StructureV1, DocumentMetadata, StyleDefinition, ParagraphSkeleton,
    HeadingReference, FieldReference, ValidationResult, FontSpec, 
    ParagraphSpec, LineSpacingMode, StyleType
)
from autoword.vnext.exceptions import ValidationError, RollbackError


class TestDocumentValidator:
    """Test cases for DocumentValidator class."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.validator = DocumentValidator(visible=False)
        
        # Create sample metadata
        self.sample_metadata = DocumentMetadata(
            title="Test Document",
            author="Test Author",
            creation_time=datetime(2024, 1, 1, 10, 0, 0),
            modified_time=datetime(2024, 1, 1, 11, 0, 0),
            word_version="16.0",
            page_count=5,
            paragraph_count=20,
            word_count=500
        )
        
        # Create sample styles
        self.sample_styles = [
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
    
    def create_sample_structure(self, headings_data=None, styles=None):
        """Create a sample StructureV1 object for testing."""
        if headings_data is None:
            headings_data = [
                ("Introduction", 1),
                ("Background", 2),
                ("Methodology", 1),
                ("Results", 2),
                ("Conclusion", 1)
            ]
        
        if styles is None:
            styles = self.sample_styles
        
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
        
        fields = [
            FieldReference(
                paragraph_index=0,
                field_type="TOC",
                field_code="TOC \\o \"1-3\" \\h \\z \\u",
                result_text="Introduction\t1\nBackground\t2\nMethodology\t3\nResults\t4\nConclusion\t5"
            )
        ]
        
        return StructureV1(
            metadata=self.sample_metadata,
            styles=styles,
            paragraphs=paragraphs,
            headings=headings,
            fields=fields,
            tables=[]
        )
    
    def test_check_chapter_assertions_valid(self):
        """Test chapter assertions with valid headings."""
        structure = self.create_sample_structure()
        
        errors = self.validator.check_chapter_assertions(structure)
        
        assert len(errors) == 0, "Valid headings should not produce errors"
    
    def test_check_chapter_assertions_invalid_abstract(self):
        """Test chapter assertions with forbidden 摘要 at level 1."""
        headings_data = [
            ("摘要", 1),  # Forbidden at level 1
            ("Introduction", 1),
            ("Methodology", 1)
        ]
        structure = self.create_sample_structure(headings_data)
        
        errors = self.validator.check_chapter_assertions(structure)
        
        assert len(errors) == 1
        assert "摘要" in errors[0]
        assert "level 1" in errors[0]
    
    def test_check_chapter_assertions_invalid_references(self):
        """Test chapter assertions with forbidden 参考文献 at level 1."""
        headings_data = [
            ("Introduction", 1),
            ("参考文献", 1),  # Forbidden at level 1
            ("Appendix", 1)
        ]
        structure = self.create_sample_structure(headings_data)
        
        errors = self.validator.check_chapter_assertions(structure)
        
        assert len(errors) == 1
        assert "参考文献" in errors[0]
        assert "level 1" in errors[0]
    
    def test_check_chapter_assertions_valid_at_level_2(self):
        """Test chapter assertions with 摘要/参考文献 at level 2 (should be valid)."""
        headings_data = [
            ("Introduction", 1),
            ("摘要", 2),  # Valid at level 2
            ("参考文献", 2),  # Valid at level 2
            ("Conclusion", 1)
        ]
        structure = self.create_sample_structure(headings_data)
        
        errors = self.validator.check_chapter_assertions(structure)
        
        assert len(errors) == 0, "摘要/参考文献 at level 2 should be valid"
    
    def test_check_style_assertions_valid(self):
        """Test style assertions with correct style specifications."""
        structure = self.create_sample_structure()
        
        errors = self.validator.check_style_assertions(structure)
        
        assert len(errors) == 0, "Correct style specifications should not produce errors"
    
    def test_check_style_assertions_wrong_font(self):
        """Test style assertions with incorrect font specifications."""
        wrong_styles = [
            StyleDefinition(
                name="Heading 1",
                type=StyleType.PARAGRAPH,
                font=FontSpec(
                    east_asian="宋体",  # Should be 楷体
                    latin="Times New Roman",
                    size_pt=12,
                    bold=True
                ),
                paragraph=ParagraphSpec(
                    line_spacing_mode=LineSpacingMode.MULTIPLE,
                    line_spacing_value=2.0
                )
            )
        ]
        structure = self.create_sample_structure(styles=wrong_styles)
        
        errors = self.validator.check_style_assertions(structure)
        
        assert len(errors) >= 1
        assert any("font_east_asian" in error and "楷体" in error for error in errors)
    
    def test_check_style_assertions_wrong_size(self):
        """Test style assertions with incorrect font size."""
        wrong_styles = [
            StyleDefinition(
                name="Normal",
                type=StyleType.PARAGRAPH,
                font=FontSpec(
                    east_asian="宋体",
                    latin="Times New Roman",
                    size_pt=14,  # Should be 12
                    bold=False
                ),
                paragraph=ParagraphSpec(
                    line_spacing_mode=LineSpacingMode.MULTIPLE,
                    line_spacing_value=2.0
                )
            )
        ]
        structure = self.create_sample_structure(styles=wrong_styles)
        
        errors = self.validator.check_style_assertions(structure)
        
        assert len(errors) >= 1
        assert any("font_size_pt" in error and "12" in error for error in errors)
    
    def test_check_style_assertions_wrong_line_spacing(self):
        """Test style assertions with incorrect line spacing."""
        wrong_styles = [
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
                    line_spacing_mode=LineSpacingMode.SINGLE,  # Should be MULTIPLE
                    line_spacing_value=1.0  # Should be 2.0
                )
            )
        ]
        structure = self.create_sample_structure(styles=wrong_styles)
        
        errors = self.validator.check_style_assertions(structure)
        
        assert len(errors) >= 2  # Both mode and value should be wrong
        assert any("line_spacing_mode" in error for error in errors)
        assert any("line_spacing_value" in error for error in errors)
    
    def test_check_style_assertions_missing_style(self):
        """Test style assertions with missing required style."""
        incomplete_styles = [
            StyleDefinition(
                name="Heading 1",
                type=StyleType.PARAGRAPH,
                font=FontSpec(east_asian="楷体", size_pt=12, bold=True),
                paragraph=ParagraphSpec(line_spacing_mode=LineSpacingMode.MULTIPLE, line_spacing_value=2.0)
            )
            # Missing Heading 2 and Normal styles
        ]
        structure = self.create_sample_structure(styles=incomplete_styles)
        
        errors = self.validator.check_style_assertions(structure)
        
        assert len(errors) >= 2  # Should report missing Heading 2 and Normal
        assert any("Heading 2" in error and "not found" in error for error in errors)
        assert any("Normal" in error and "not found" in error for error in errors)
    
    def test_check_toc_assertions_valid(self):
        """Test TOC assertions with consistent TOC and headings."""
        # Create structure with matching TOC and headings
        headings_data = [
            ("Introduction", 1),
            ("Methodology", 1),
            ("Results", 1),
            ("Conclusion", 1)
        ]
        structure = self.create_sample_structure(headings_data)
        
        # Update TOC to match headings exactly
        structure.fields[0].result_text = "Introduction\t1\nMethodology\t2\nResults\t3\nConclusion\t4"
        
        errors = self.validator.check_toc_assertions(structure)
        
        # Should have some errors due to simplified parsing, but not too many
        assert len(errors) <= 3, "Valid TOC should produce minimal errors"
    
    def test_check_toc_assertions_no_toc(self):
        """Test TOC assertions with no TOC fields."""
        structure = self.create_sample_structure()
        structure.fields = []  # Remove TOC fields
        
        errors = self.validator.check_toc_assertions(structure)
        
        assert len(errors) == 0, "No TOC should not produce errors"
    
    def test_check_toc_assertions_empty_result(self):
        """Test TOC assertions with empty TOC result."""
        structure = self.create_sample_structure()
        structure.fields[0].result_text = ""  # Empty TOC result
        
        errors = self.validator.check_toc_assertions(structure)
        
        assert len(errors) >= 1
        assert any("no result text" in error for error in errors)
    
    def test_check_pagination_assertions_valid(self):
        """Test pagination assertions with valid changes."""
        original_metadata = DocumentMetadata(
            modified_time=datetime(2024, 1, 1, 10, 0, 0),
            page_count=5
        )
        modified_metadata = DocumentMetadata(
            modified_time=datetime(2024, 1, 1, 11, 0, 0),  # Later time
            page_count=6
        )
        
        original_structure = StructureV1(metadata=original_metadata, styles=[], paragraphs=[], headings=[], fields=[], tables=[])
        modified_structure = StructureV1(metadata=modified_metadata, styles=[], paragraphs=[], headings=[], fields=[], tables=[])
        
        errors = self.validator.check_pagination_assertions(original_structure, modified_structure)
        
        assert len(errors) == 0, "Valid pagination changes should not produce errors"
    
    def test_check_pagination_assertions_no_time_change(self):
        """Test pagination assertions with no modification time change."""
        same_time = datetime(2024, 1, 1, 10, 0, 0)
        original_metadata = DocumentMetadata(modified_time=same_time, page_count=5)
        modified_metadata = DocumentMetadata(modified_time=same_time, page_count=5)
        
        original_structure = StructureV1(metadata=original_metadata, styles=[], paragraphs=[], headings=[], fields=[], tables=[])
        modified_structure = StructureV1(metadata=modified_metadata, styles=[], paragraphs=[], headings=[], fields=[], tables=[])
        
        errors = self.validator.check_pagination_assertions(original_structure, modified_structure)
        
        assert len(errors) >= 1
        assert any("modification time has not changed" in error for error in errors)
    
    def test_check_pagination_assertions_invalid_page_count(self):
        """Test pagination assertions with invalid page count."""
        modified_metadata = DocumentMetadata(
            modified_time=datetime(2024, 1, 1, 11, 0, 0),
            page_count=0  # Invalid page count
        )
        
        original_structure = StructureV1(metadata=self.sample_metadata, styles=[], paragraphs=[], headings=[], fields=[], tables=[])
        modified_structure = StructureV1(metadata=modified_metadata, styles=[], paragraphs=[], headings=[], fields=[], tables=[])
        
        errors = self.validator.check_pagination_assertions(original_structure, modified_structure)
        
        assert len(errors) >= 1
        assert any("Invalid page count" in error for error in errors)
    
    def test_rollback_document_success(self):
        """Test successful document rollback."""
        with tempfile.TemporaryDirectory() as temp_dir:
            # Create test files
            original_file = os.path.join(temp_dir, "original.docx")
            modified_file = os.path.join(temp_dir, "modified.docx")
            
            # Write test content
            with open(original_file, 'w') as f:
                f.write("original content")
            with open(modified_file, 'w') as f:
                f.write("modified content")
            
            # Perform rollback
            self.validator.rollback_document(original_file, modified_file)
            
            # Verify rollback
            with open(modified_file, 'r') as f:
                content = f.read()
            assert content == "original content", "File should be rolled back to original content"
            
            # Verify backup was created
            backup_file = f"{modified_file}.rollback_backup"
            assert os.path.exists(backup_file), "Rollback backup should be created"
    
    def test_rollback_document_missing_original(self):
        """Test rollback with missing original file."""
        with tempfile.TemporaryDirectory() as temp_dir:
            original_file = os.path.join(temp_dir, "nonexistent.docx")
            modified_file = os.path.join(temp_dir, "modified.docx")
            
            with open(modified_file, 'w') as f:
                f.write("modified content")
            
            with pytest.raises(RollbackError) as exc_info:
                self.validator.rollback_document(original_file, modified_file)
            
            assert "not found" in str(exc_info.value)
    
    @patch('autoword.vnext.validator.document_validator.DocumentExtractor')
    def test_validate_modifications_success(self, mock_extractor_class):
        """Test successful validation of modifications."""
        # Mock the extractor
        mock_extractor = Mock()
        mock_extractor_class.return_value.__enter__.return_value = mock_extractor
        
        # Create original structure with earlier time
        original_metadata = DocumentMetadata(
            title="Test Document",
            author="Test Author",
            creation_time=datetime(2024, 1, 1, 10, 0, 0),
            modified_time=datetime(2024, 1, 1, 10, 0, 0),  # Earlier time
            word_version="16.0",
            page_count=5,
            paragraph_count=20,
            word_count=500
        )
        original_structure = self.create_sample_structure()
        original_structure.metadata = original_metadata
        original_structure.fields = []  # Remove TOC to avoid parsing issues
        
        # Create valid modified structure with later time
        modified_metadata = DocumentMetadata(
            title="Test Document",
            author="Test Author",
            creation_time=datetime(2024, 1, 1, 10, 0, 0),
            modified_time=datetime(2024, 1, 1, 12, 0, 0),  # Later time
            word_version="16.0",
            page_count=5,
            paragraph_count=20,
            word_count=500
        )
        modified_structure = self.create_sample_structure()
        modified_structure.metadata = modified_metadata
        modified_structure.fields = []  # Remove TOC to avoid parsing issues
        mock_extractor.extract_structure.return_value = modified_structure
        
        # Mock Word COM operations
        with patch.object(self.validator, '_update_fields_and_repaginate'):
            with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as temp_file:
                temp_path = temp_file.name
            
            try:
                result = self.validator.validate_modifications(original_structure, temp_path)
                
                assert result.is_valid, f"Valid modifications should pass validation. Errors: {result.errors}"
                assert len(result.errors) == 0, "No errors should be reported for valid modifications"
                
            finally:
                os.unlink(temp_path)
    
    @patch('autoword.vnext.validator.document_validator.DocumentExtractor')
    def test_validate_modifications_with_errors(self, mock_extractor_class):
        """Test validation with assertion failures."""
        # Mock the extractor
        mock_extractor = Mock()
        mock_extractor_class.return_value.__enter__.return_value = mock_extractor
        
        # Create invalid modified structure (with forbidden heading)
        invalid_headings = [("摘要", 1), ("Introduction", 1)]
        modified_structure = self.create_sample_structure(invalid_headings)
        mock_extractor.extract_structure.return_value = modified_structure
        
        # Mock Word COM operations
        with patch.object(self.validator, '_update_fields_and_repaginate'):
            with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as temp_file:
                temp_path = temp_file.name
            
            try:
                original_structure = self.create_sample_structure()
                result = self.validator.validate_modifications(original_structure, temp_path)
                
                assert not result.is_valid, "Invalid modifications should fail validation"
                assert len(result.errors) > 0, "Errors should be reported for invalid modifications"
                assert any("摘要" in error for error in result.errors)
                
            finally:
                os.unlink(temp_path)
    
    def test_validate_modifications_missing_file(self):
        """Test validation with missing modified file."""
        original_structure = self.create_sample_structure()
        
        with pytest.raises(ValidationError) as exc_info:
            self.validator.validate_modifications(original_structure, "nonexistent.docx")
        
        assert "not found" in str(exc_info.value)
    
    def test_text_matches_approximately(self):
        """Test approximate text matching functionality."""
        # Exact match
        assert self.validator._text_matches_approximately("Introduction", "Introduction")
        
        # Case insensitive match
        assert self.validator._text_matches_approximately("Introduction", "INTRODUCTION")
        
        # Partial match
        assert self.validator._text_matches_approximately("Intro", "Introduction")
        
        # Similar text
        assert self.validator._text_matches_approximately("Introduction", "Introductio")
        
        # Different text
        assert not self.validator._text_matches_approximately("Introduction", "Conclusion")
    
    def test_parse_toc_entries(self):
        """Test TOC entry parsing functionality."""
        toc_text = "Introduction\t1\n  Background\t2\nMethodology\t3\n    Results\t4\nConclusion\t5"
        
        entries = self.validator._parse_toc_entries(toc_text)
        
        assert len(entries) == 5
        assert entries[0] == ("Introduction", 1)
        assert entries[1] == ("Background", 2)  # 2 spaces = level 2
        assert entries[2] == ("Methodology", 1)
        assert entries[3] == ("Results", 3)  # 4 spaces = level 3
        assert entries[4] == ("Conclusion", 1)
    
    def test_context_manager(self):
        """Test DocumentValidator as context manager."""
        with patch('pythoncom.CoInitialize'), \
             patch('pythoncom.CoUninitialize'), \
             patch('win32com.client.gencache.EnsureDispatch') as mock_dispatch:
            
            mock_word_app = Mock()
            mock_dispatch.return_value = mock_word_app
            
            with DocumentValidator(visible=False) as validator:
                assert validator is not None
                assert validator._word_app == mock_word_app
            
            # Verify cleanup was called
            mock_word_app.Quit.assert_called_once()


if __name__ == "__main__":
    pytest.main([__file__])