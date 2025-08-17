"""
Unit tests for AutoWord vNext DocumentExtractor.

This module provides comprehensive unit tests for the DocumentExtractor class,
including mock Word COM objects and validation of extraction functionality.
"""

import os
import pytest
import tempfile
import zipfile
from unittest.mock import Mock, MagicMock, patch, mock_open
from datetime import datetime
from pathlib import Path

from autoword.vnext.extractor.document_extractor import DocumentExtractor
from autoword.vnext.models import (
    StructureV1, InventoryFullV1, DocumentMetadata, StyleDefinition,
    ParagraphSkeleton, HeadingReference, FieldReference, TableSkeleton,
    FontSpec, ParagraphSpec, MediaReference, ContentControlReference,
    FormulaReference, ChartReference, StyleType, LineSpacingMode
)
from autoword.vnext.exceptions import ExtractionError


class TestDocumentExtractor:
    """Test cases for DocumentExtractor class."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.extractor = DocumentExtractor(visible=False)
        self.test_docx_path = "test_document.docx"
    
    def teardown_method(self):
        """Clean up after tests."""
        pass
    
    @patch('autoword.vnext.extractor.document_extractor.pythoncom')
    @patch('autoword.vnext.extractor.document_extractor.win32')
    def test_context_manager_initialization(self, mock_win32, mock_pythoncom):
        """Test context manager initialization and cleanup."""
        # Setup mocks
        mock_word_app = Mock()
        mock_win32.gencache.EnsureDispatch.return_value = mock_word_app
        
        # Test context manager
        with self.extractor as extractor:
            assert extractor is self.extractor
            mock_pythoncom.CoInitialize.assert_called_once()
            mock_win32.gencache.EnsureDispatch.assert_called_once_with('Word.Application')
            assert mock_word_app.Visible == False
            assert mock_word_app.DisplayAlerts == 0
            assert mock_word_app.ScreenUpdating == False
        
        # Verify cleanup
        mock_word_app.Quit.assert_called_once_with(SaveChanges=0)
        mock_pythoncom.CoUninitialize.assert_called_once()
    
    @patch('autoword.vnext.extractor.document_extractor.pythoncom')
    @patch('autoword.vnext.extractor.document_extractor.win32')
    def test_context_manager_initialization_failure(self, mock_win32, mock_pythoncom):
        """Test context manager initialization failure handling."""
        # Setup mock to raise exception
        mock_win32.gencache.EnsureDispatch.side_effect = Exception("COM initialization failed")
        
        # Test exception handling
        with pytest.raises(ExtractionError) as exc_info:
            with self.extractor:
                pass
        
        assert "Failed to initialize Word COM for extraction" in str(exc_info.value)
        mock_pythoncom.CoUninitialize.assert_called_once()
    
    @patch('os.path.exists')
    def test_extract_structure_file_not_found(self, mock_exists):
        """Test extract_structure with non-existent file."""
        mock_exists.return_value = False
        
        with pytest.raises(ExtractionError) as exc_info:
            self.extractor.extract_structure("nonexistent.docx")
        
        assert "DOCX file not found" in str(exc_info.value)
        assert exc_info.value.details['docx_path'] == "nonexistent.docx"
    
    @patch('os.path.exists')
    @patch('autoword.vnext.extractor.document_extractor.pythoncom')
    @patch('autoword.vnext.extractor.document_extractor.win32')
    def test_extract_structure_success(self, mock_win32, mock_pythoncom, mock_exists):
        """Test successful structure extraction."""
        # Setup mocks
        mock_exists.return_value = True
        mock_word_app = Mock()
        mock_win32.gencache.EnsureDispatch.return_value = mock_word_app
        
        # Mock document
        mock_doc = Mock()
        mock_word_app.Documents.Open.return_value = mock_doc
        
        # Mock document properties
        mock_props = Mock()
        mock_doc.BuiltInDocumentProperties = mock_props
        
        # Mock individual properties
        title_prop = Mock()
        title_prop.Value = None
        mock_props.side_effect = lambda prop_name: title_prop
        
        # Mock statistics
        mock_doc.ComputeStatistics.side_effect = [10, 50, 500]  # pages, paragraphs, words
        
        # Mock Word version
        self.extractor._word_app = mock_word_app
        mock_word_app.Version = None
        
        # Mock styles
        mock_style = Mock()
        mock_style.NameLocal = "Normal"
        mock_style.Type = 1  # wdStyleTypeParagraph
        mock_style.BuiltIn = True
        mock_style.InUse = True
        mock_style.Font = Mock()
        mock_style.Font.Name = "Calibri"
        mock_style.Font.Size = 11
        mock_style.Font.Bold = False
        mock_style.Font.Italic = False
        mock_style.Font.Color = -9999999
        # Mock base style attributes to return None
        mock_style.BaseStyle = None
        mock_style.NextParagraphStyle = None
        mock_doc.Styles = [mock_style]
        
        # Mock paragraphs
        mock_para = Mock()
        mock_para.Range.Text = "This is a test paragraph."
        mock_para.Style.NameLocal = "Normal"
        mock_para.OutlineLevel = 10  # wdOutlineLevelBodyText
        mock_doc.Paragraphs = [mock_para]
        
        # Mock fields and tables
        mock_doc.Fields = []
        mock_doc.Tables = []
        
        # Test extraction
        with self.extractor:
            structure = self.extractor.extract_structure(self.test_docx_path)
        
        # Verify results
        assert isinstance(structure, StructureV1)
        assert structure.schema_version == "structure.v1"
        assert len(structure.paragraphs) == 1
        assert structure.paragraphs[0].preview_text == "This is a test paragraph."
        assert len(structure.styles) == 1
        assert structure.styles[0].name == "Normal"
        
        # Verify document was closed
        mock_doc.Close.assert_called_once_with(SaveChanges=0)
    
    @patch('os.path.exists')
    @patch('autoword.vnext.extractor.document_extractor.pythoncom')
    @patch('autoword.vnext.extractor.document_extractor.win32')
    def test_extract_structure_with_headings(self, mock_win32, mock_pythoncom, mock_exists):
        """Test structure extraction with headings."""
        # Setup mocks
        mock_exists.return_value = True
        mock_word_app = Mock()
        mock_win32.gencache.EnsureDispatch.return_value = mock_word_app
        
        # Mock document
        mock_doc = Mock()
        mock_word_app.Documents.Open.return_value = mock_doc
        
        # Mock basic properties
        mock_props = Mock()
        mock_doc.BuiltInDocumentProperties = mock_props
        title_prop = Mock()
        title_prop.Value = None
        mock_props.side_effect = lambda prop_name: title_prop
        
        # Mock Word version
        self.extractor._word_app = mock_word_app
        mock_word_app.Version = None
        mock_doc.ComputeStatistics.side_effect = [1, 2, 20]
        mock_doc.Styles = []
        mock_doc.Fields = []
        mock_doc.Tables = []
        
        # Mock paragraphs with headings
        mock_para1 = Mock()
        mock_para1.Range.Text = "Chapter 1: Introduction"
        mock_para1.Style.NameLocal = "Heading 1"
        mock_para1.OutlineLevel = 1  # Heading level 1
        
        mock_para2 = Mock()
        mock_para2.Range.Text = "This is the introduction content."
        mock_para2.Style.NameLocal = "Normal"
        mock_para2.OutlineLevel = 10  # Body text
        
        mock_doc.Paragraphs = [mock_para1, mock_para2]
        
        # Test extraction
        with self.extractor:
            structure = self.extractor.extract_structure(self.test_docx_path)
        
        # Verify headings
        assert len(structure.headings) == 1
        assert structure.headings[0].level == 1
        assert structure.headings[0].text == "Chapter 1: Introduction"
        assert structure.headings[0].paragraph_index == 0
        
        # Verify paragraphs
        assert len(structure.paragraphs) == 2
        assert structure.paragraphs[0].is_heading == True
        assert structure.paragraphs[0].heading_level == 1
        assert structure.paragraphs[1].is_heading == False
    
    @patch('os.path.exists')
    def test_extract_inventory_file_not_found(self, mock_exists):
        """Test extract_inventory with non-existent file."""
        mock_exists.return_value = False
        
        with pytest.raises(ExtractionError) as exc_info:
            self.extractor.extract_inventory("nonexistent.docx")
        
        assert "DOCX file not found" in str(exc_info.value)
        assert exc_info.value.details['extraction_stage'] == "inventory"
    
    @patch('os.path.exists')
    @patch('autoword.vnext.extractor.document_extractor.zipfile.ZipFile')
    @patch('autoword.vnext.extractor.document_extractor.pythoncom')
    @patch('autoword.vnext.extractor.document_extractor.win32')
    def test_extract_inventory_success(self, mock_win32, mock_pythoncom, mock_zipfile, mock_exists):
        """Test successful inventory extraction."""
        # Setup mocks
        mock_exists.return_value = True
        mock_word_app = Mock()
        mock_win32.gencache.EnsureDispatch.return_value = mock_word_app
        
        # Mock document
        mock_doc = Mock()
        mock_word_app.Documents.Open.return_value = mock_doc
        mock_doc.ContentControls = []
        mock_doc.InlineShapes = []
        
        # Mock ZIP file
        mock_zip = Mock()
        mock_zipfile.return_value.__enter__.return_value = mock_zip
        mock_zip.namelist.return_value = [
            'word/document.xml',
            'word/styles.xml',
            'word/media/image1.png',
            'word/_rels/document.xml.rels'
        ]
        mock_zip.read.return_value = b'<xml>test content</xml>'
        
        # Mock file info for media
        mock_file_info = Mock()
        mock_file_info.file_size = 1024
        mock_zip.getinfo.return_value = mock_file_info
        
        # Test extraction
        with self.extractor:
            inventory = self.extractor.extract_inventory(self.test_docx_path)
        
        # Verify results
        assert isinstance(inventory, InventoryFullV1)
        assert inventory.schema_version == "inventory.full.v1"
        assert len(inventory.ooxml_fragments) > 0
        assert 'word/document.xml' in inventory.ooxml_fragments
        assert len(inventory.media_indexes) == 1
        assert 'word/media/image1.png' in inventory.media_indexes
        
        # Verify document was closed
        mock_doc.Close.assert_called_once_with(SaveChanges=0)
    
    @patch('os.path.exists')
    @patch('autoword.vnext.extractor.document_extractor.pythoncom')
    @patch('autoword.vnext.extractor.document_extractor.win32')
    def test_process_document_success(self, mock_win32, mock_pythoncom, mock_exists):
        """Test successful document processing."""
        # Setup mocks
        mock_exists.return_value = True
        mock_word_app = Mock()
        mock_win32.gencache.EnsureDispatch.return_value = mock_word_app
        
        # Mock document for structure extraction
        mock_doc = Mock()
        mock_word_app.Documents.Open.return_value = mock_doc
        
        # Mock minimal document structure
        mock_props = Mock()
        mock_doc.BuiltInDocumentProperties = mock_props
        title_prop = Mock()
        title_prop.Value = None
        mock_props.side_effect = lambda prop_name: title_prop
        
        # Mock Word version
        self.extractor._word_app = mock_word_app
        mock_word_app.Version = None
        mock_doc.ComputeStatistics.side_effect = [1, 1, 10]
        mock_doc.Styles = []
        mock_doc.Paragraphs = []
        mock_doc.Fields = []
        mock_doc.Tables = []
        mock_doc.ContentControls = []
        mock_doc.InlineShapes = []
        
        # Mock ZIP file for inventory
        with patch('autoword.vnext.extractor.document_extractor.zipfile.ZipFile') as mock_zipfile:
            mock_zip = Mock()
            mock_zipfile.return_value.__enter__.return_value = mock_zip
            mock_zip.namelist.return_value = ['word/document.xml']
            mock_zip.read.return_value = b'<xml>test</xml>'
            
            # Test processing
            with self.extractor:
                structure, inventory = self.extractor.process_document(self.test_docx_path)
        
        # Verify results
        assert isinstance(structure, StructureV1)
        assert isinstance(inventory, InventoryFullV1)
        
        # Verify document was opened twice (once for structure, once for inventory)
        assert mock_word_app.Documents.Open.call_count == 2
    
    def test_extract_metadata_with_mock_properties(self):
        """Test metadata extraction with mock document properties."""
        # Create mock document
        mock_doc = Mock()
        
        # Mock built-in properties
        mock_props = Mock()
        mock_doc.BuiltInDocumentProperties = mock_props
        
        # Mock individual properties
        title_prop = Mock()
        title_prop.Value = "Test Document Title"
        author_prop = Mock()
        author_prop.Value = "Test Author"
        creation_prop = Mock()
        creation_prop.Value = "2024-01-01T10:00:00Z"
        modified_prop = Mock()
        modified_prop.Value = "2024-01-02T15:30:00Z"
        
        mock_props.side_effect = lambda prop_name: {
            "Title": title_prop,
            "Author": author_prop,
            "Creation Date": creation_prop,
            "Last Save Time": modified_prop
        }.get(prop_name, Mock(Value=None))
        
        # Mock statistics - need to handle exceptions
        def mock_compute_stats(stat_type):
            if stat_type == 2:  # wdStatisticPages
                return 5
            elif stat_type == 1:  # wdStatisticParagraphs
                return 25
            elif stat_type == 0:  # wdStatisticWords
                return 250
            return 0
        
        mock_doc.ComputeStatistics.side_effect = mock_compute_stats
        
        # Mock Word version
        self.extractor._word_app = Mock()
        self.extractor._word_app.Version = "16.0"
        
        # Test metadata extraction
        metadata = self.extractor._extract_metadata(mock_doc)
        
        # Verify results
        assert metadata.title == "Test Document Title"
        assert metadata.author == "Test Author"
        assert metadata.word_version == "16.0"
        assert metadata.page_count == 5
        assert metadata.paragraph_count == 25
        assert metadata.word_count == 250
        assert isinstance(metadata.creation_time, datetime)
        assert isinstance(metadata.modified_time, datetime)
    
    def test_extract_styles_with_mock_styles(self):
        """Test style extraction with mock Word styles."""
        # Create mock document
        mock_doc = Mock()
        
        # Create mock style
        mock_style = Mock()
        mock_style.NameLocal = "Heading 1"
        mock_style.Type = 1  # wdStyleTypeParagraph
        mock_style.BuiltIn = True
        mock_style.InUse = True
        
        # Mock font
        mock_font = Mock()
        mock_font.Name = "Calibri"
        mock_font.NameFarEast = "宋体"
        mock_font.Size = 16
        mock_font.Bold = True
        mock_font.Italic = False
        mock_font.Color = 0x000000  # Black
        mock_style.Font = mock_font
        
        # Mock paragraph format
        mock_para_format = Mock()
        mock_para_format.LineSpacingRule = 0  # wdLineSpaceSingle
        mock_para_format.LineSpacing = 1.0
        mock_para_format.SpaceBefore = 12
        mock_para_format.SpaceAfter = 6
        mock_para_format.LeftIndent = 0
        mock_para_format.RightIndent = 0
        mock_para_format.FirstLineIndent = 0
        mock_style.ParagraphFormat = mock_para_format
        
        # Mock base style attributes to return None
        mock_style.BaseStyle = None
        mock_style.NextParagraphStyle = None
        
        mock_doc.Styles = [mock_style]
        
        # Test style extraction
        styles = self.extractor._extract_styles(mock_doc)
        
        # Verify results
        assert len(styles) == 1
        style = styles[0]
        assert style.name == "Heading 1"
        assert style.type == StyleType.PARAGRAPH
        assert style.font.latin == "Calibri"
        assert style.font.east_asian == "宋体"
        assert style.font.size_pt == 16
        assert style.font.bold == True
        assert style.font.italic == False
        assert style.paragraph.line_spacing_mode == LineSpacingMode.SINGLE
        assert style.paragraph.space_before_pt == 12
        assert style.paragraph.space_after_pt == 6
    
    def test_extract_paragraphs_with_mock_paragraphs(self):
        """Test paragraph extraction with mock Word paragraphs."""
        # Create mock document
        mock_doc = Mock()
        
        # Create mock paragraphs
        mock_para1 = Mock()
        mock_para1.Range.Text = "This is a normal paragraph with some content."
        mock_para1.Style.NameLocal = "Normal"
        mock_para1.OutlineLevel = 10  # wdOutlineLevelBodyText
        
        mock_para2 = Mock()
        mock_para2.Range.Text = "Chapter 1: Introduction"
        mock_para2.Style.NameLocal = "Heading 1"
        mock_para2.OutlineLevel = 1  # Heading level 1
        
        mock_doc.Paragraphs = [mock_para1, mock_para2]
        
        # Test paragraph extraction
        paragraphs = self.extractor._extract_paragraphs(mock_doc)
        
        # Verify results
        assert len(paragraphs) == 2
        
        # Check normal paragraph
        para1 = paragraphs[0]
        assert para1.index == 0
        assert para1.style_name == "Normal"
        assert para1.preview_text == "This is a normal paragraph with some content."
        assert para1.is_heading == False
        assert para1.heading_level is None
        
        # Check heading paragraph
        para2 = paragraphs[1]
        assert para2.index == 1
        assert para2.style_name == "Heading 1"
        assert para2.preview_text == "Chapter 1: Introduction"
        assert para2.is_heading == True
        assert para2.heading_level == 1
    
    def test_extract_fields_with_mock_fields(self):
        """Test field extraction with mock Word fields."""
        # Create mock document
        mock_doc = Mock()
        
        # Create mock field
        mock_field = Mock()
        mock_field.Type = 13  # wdFieldTOC
        mock_field.Code.Text = "TOC \\o \"1-3\" \\h \\z \\u"
        mock_field.Result.Text = "Table of Contents"
        
        # Mock field range and paragraph
        mock_field_range = Mock()
        mock_field_range.Paragraphs.return_value = [Mock()]
        mock_field_range.Paragraphs.return_value[0].Range.Start = 100
        mock_field.Range = mock_field_range
        
        # Mock document paragraphs
        mock_para = Mock()
        mock_para.Range.Start = 100
        mock_doc.Paragraphs = [mock_para]
        mock_doc.Fields = [mock_field]
        
        # Test field extraction
        fields = self.extractor._extract_fields(mock_doc)
        
        # Verify results
        assert len(fields) == 1
        field = fields[0]
        assert field.paragraph_index == 0
        assert field.field_type == "13"
        assert field.field_code == "TOC \\o \"1-3\" \\h \\z \\u"
        assert field.result_text == "Table of Contents"
    
    def test_extract_tables_with_mock_tables(self):
        """Test table extraction with mock Word tables."""
        # Create mock document
        mock_doc = Mock()
        
        # Create mock table
        mock_table = Mock()
        mock_table.Rows.Count = 3
        mock_table.Columns.Count = 2
        
        # Mock table range
        mock_table_range = Mock()
        mock_table_range.Start = 200
        mock_table_range.End = 300
        mock_table.Range = mock_table_range
        
        # Mock document paragraphs
        mock_para = Mock()
        mock_para.Range.Start = 250
        mock_para.Range.End = 260
        mock_doc.Paragraphs = [mock_para]
        mock_doc.Tables = [mock_table]
        
        # Test table extraction
        tables = self.extractor._extract_tables(mock_doc)
        
        # Verify results
        assert len(tables) == 1
        table = tables[0]
        assert table.paragraph_index == 0
        assert table.rows == 3
        assert table.columns == 2
        assert table.has_header == True  # Simplified assumption
    
    def test_rgb_to_hex_conversion(self):
        """Test RGB to hex color conversion."""
        # Test normal color
        hex_color = self.extractor._rgb_to_hex(0xFF0000)  # Red
        assert hex_color == "#0000FF"  # Note: RGB order is reversed in Word
        
        # Test automatic color
        hex_color = self.extractor._rgb_to_hex(-9999999)
        assert hex_color is None
        
        # Test black
        hex_color = self.extractor._rgb_to_hex(0x000000)
        assert hex_color == "#000000"
    
    @patch('autoword.vnext.extractor.document_extractor.zipfile.ZipFile')
    def test_extract_ooxml_fragments(self, mock_zipfile):
        """Test OOXML fragment extraction."""
        # Mock ZIP file
        mock_zip = Mock()
        mock_zipfile.return_value.__enter__.return_value = mock_zip
        mock_zip.namelist.return_value = [
            'word/document.xml',
            'word/styles.xml',
            'word/_rels/document.xml.rels',
            'docProps/core.xml'
        ]
        mock_zip.read.return_value = b'<xml>test content</xml>'
        
        # Test OOXML extraction
        fragments = self.extractor._extract_ooxml_fragments(self.test_docx_path)
        
        # Verify results
        assert len(fragments) == 4
        assert 'word/document.xml' in fragments
        assert 'word/styles.xml' in fragments
        assert 'word/_rels/document.xml.rels' in fragments
        assert 'docProps/core.xml' in fragments
        assert all(content == '<xml>test content</xml>' for content in fragments.values())
    
    @patch('autoword.vnext.extractor.document_extractor.zipfile.ZipFile')
    def test_extract_media_indexes(self, mock_zipfile):
        """Test media index extraction."""
        # Mock ZIP file
        mock_zip = Mock()
        mock_zipfile.return_value.__enter__.return_value = mock_zip
        mock_zip.namelist.return_value = [
            'word/media/image1.png',
            'word/media/image2.jpg',
            'word/media/chart1.emf'
        ]
        
        # Mock file info
        mock_file_info = Mock()
        mock_file_info.file_size = 2048
        mock_zip.getinfo.return_value = mock_file_info
        
        # Test media extraction
        media_indexes = self.extractor._extract_media_indexes(self.test_docx_path)
        
        # Verify results
        assert len(media_indexes) == 3
        
        png_media = media_indexes['word/media/image1.png']
        assert png_media.media_id == 'image1'
        assert png_media.content_type == 'image/png'
        assert png_media.file_extension == '.png'
        assert png_media.size_bytes == 2048
        
        jpg_media = media_indexes['word/media/image2.jpg']
        assert jpg_media.content_type == 'image/jpeg'
        
        emf_media = media_indexes['word/media/chart1.emf']
        assert emf_media.content_type == 'image/emf'


class TestDocumentExtractorIntegration:
    """Integration tests for DocumentExtractor with real-like scenarios."""
    
    def setup_method(self):
        """Set up integration test fixtures."""
        self.extractor = DocumentExtractor(visible=False)
        self.test_docx_path = "test_document.docx"
    
    def test_extraction_error_handling(self):
        """Test error handling during extraction."""
        # Test with non-existent file
        with pytest.raises(ExtractionError) as exc_info:
            self.extractor.extract_structure("nonexistent.docx")
        
        assert "DOCX file not found" in str(exc_info.value)
        assert exc_info.value.details['docx_path'] == "nonexistent.docx"
    
    @patch('os.path.exists')
    @patch('autoword.vnext.extractor.document_extractor.pythoncom')
    @patch('autoword.vnext.extractor.document_extractor.win32')
    def test_com_error_handling(self, mock_win32, mock_pythoncom, mock_exists):
        """Test COM error handling during extraction."""
        # Setup mocks
        mock_exists.return_value = True
        mock_word_app = Mock()
        mock_win32.gencache.EnsureDispatch.return_value = mock_word_app
        
        # Mock document opening to raise exception
        mock_word_app.Documents.Open.side_effect = Exception("COM error")
        
        # Test error handling
        with pytest.raises(ExtractionError) as exc_info:
            with self.extractor:
                self.extractor.extract_structure(self.test_docx_path)
        
        assert "Failed to extract structure" in str(exc_info.value)
        assert exc_info.value.details['extraction_stage'] == "structure"
    
    def test_long_text_truncation(self):
        """Test that long paragraph text is properly truncated."""
        # Create a long text string
        long_text = "A" * 200
        
        # Create mock paragraph
        mock_para = Mock()
        mock_para.Range.Text = long_text
        mock_para.Style.NameLocal = "Normal"
        mock_para.OutlineLevel = 10
        
        mock_doc = Mock()
        mock_doc.Paragraphs = [mock_para]
        
        # Test paragraph extraction
        paragraphs = self.extractor._extract_paragraphs(mock_doc)
        
        # Verify truncation
        assert len(paragraphs) == 1
        assert len(paragraphs[0].preview_text) == 120
        assert paragraphs[0].preview_text == "A" * 120


if __name__ == "__main__":
    pytest.main([__file__])