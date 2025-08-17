"""
Test complex document element handling in the extractor.

This module tests the enhanced extractor functionality for:
- Tables with cell references to paragraph indexes
- Formula and content control detection with inventory storage
- Chart/SmartArt/OLE object handling as OOXML/binary in inventory
- Footnote/endnote reference preservation in structure extraction
- Cross-reference identification and relationship mapping
"""

import pytest
import os
import tempfile
from unittest.mock import Mock, MagicMock, patch
from pathlib import Path

from autoword.vnext.extractor.document_extractor import DocumentExtractor
from autoword.vnext.models import (
    StructureV1, InventoryFullV1, TableSkeleton, FootnoteReference,
    EndnoteReference, CrossReference, FormulaReference, ChartReference,
    ContentControlReference
)
from autoword.vnext.exceptions import ExtractionError


class TestComplexDocumentElements:
    """Test complex document element extraction functionality."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.test_docx_path = "test_document.docx"
        
    def create_mock_word_app(self):
        """Create a comprehensive mock Word application."""
        mock_app = Mock()
        mock_doc = Mock()
        
        # Mock basic document properties
        mock_doc.Paragraphs = []
        mock_doc.Styles = []
        mock_doc.Fields = []
        mock_doc.Tables = []
        mock_doc.InlineShapes = []
        mock_doc.Shapes = []
        mock_doc.Footnotes = []
        mock_doc.Endnotes = []
        mock_doc.Hyperlinks = []
        mock_doc.Bookmarks = []
        mock_doc.ContentControls = []
        
        # Mock document opening
        mock_app.Documents.Open.return_value = mock_doc
        
        return mock_app, mock_doc
    
    def test_enhanced_table_extraction_with_cell_references(self):
        """Test enhanced table extraction with cell-to-paragraph mapping."""
        mock_app, mock_doc = self.create_mock_word_app()
        
        # Create mock table with cells
        mock_table = Mock()
        mock_table.Rows.Count = 2
        mock_table.Columns.Count = 2
        mock_table.Range.Start = 100
        mock_table.Range.End = 200
        
        # Mock table rows collection
        mock_rows = Mock()
        mock_rows.Count = 2
        mock_table.Rows = mock_rows
        
        # Mock cells with paragraphs
        mock_cell11 = Mock()
        mock_cell12 = Mock()
        mock_cell21 = Mock()
        mock_cell22 = Mock()
        
        # Mock cell ranges and paragraphs
        mock_para1 = Mock()
        mock_para1.Range.Start = 110
        mock_para1.Range.End = 120
        mock_cell11.Range.Paragraphs = [mock_para1]
        
        mock_para2 = Mock()
        mock_para2.Range.Start = 130
        mock_para2.Range.End = 140
        mock_cell12.Range.Paragraphs = [mock_para2]
        
        # Mock table.Cell() method
        def mock_cell(row, col):
            if row == 1 and col == 1:
                return mock_cell11
            elif row == 1 and col == 2:
                return mock_cell12
            elif row == 2 and col == 1:
                return mock_cell21
            elif row == 2 and col == 2:
                return mock_cell22
            
        mock_table.Cell = mock_cell
        
        # Mock document paragraphs
        mock_doc_para1 = Mock()
        mock_doc_para1.Range.Start = 110
        mock_doc_para1.Range.End = 120
        mock_doc_para2 = Mock()
        mock_doc_para2.Range.Start = 130
        mock_doc_para2.Range.End = 140
        
        mock_doc.Paragraphs = [mock_doc_para1, mock_doc_para2]
        mock_doc.Tables = [mock_table]
        
        with patch('pythoncom.CoInitialize'), \
             patch('pythoncom.CoUninitialize'), \
             patch('win32com.client.gencache.EnsureDispatch', return_value=mock_app):
            
            with DocumentExtractor() as extractor:
                tables = extractor._extract_tables(mock_doc)
                
                assert len(tables) == 1
                table = tables[0]
                assert isinstance(table, TableSkeleton)
                assert table.rows == 2
                assert table.columns == 2
                assert table.cell_references is not None
                assert table.cell_paragraph_map is not None
                assert "1,1" in table.cell_paragraph_map
                assert "1,2" in table.cell_paragraph_map
    
    def test_enhanced_formula_extraction(self):
        """Test enhanced formula and equation extraction."""
        mock_app, mock_doc = self.create_mock_word_app()
        
        # Create mock inline shapes with formulas
        mock_shape1 = Mock()
        mock_shape1.Type = 8  # wdInlineShapeOLEControlObject
        mock_shape1.Range.Start = 50
        mock_shape1.Range.End = 60
        mock_shape1.OLEFormat.ClassType = "Microsoft Equation 3.0"
        
        mock_shape2 = Mock()
        mock_shape2.Type = 7  # wdInlineShapeEmbeddedOLEObject
        mock_shape2.Range.Start = 150
        mock_shape2.Range.End = 160
        mock_shape2.OLEFormat.ClassType = "Excel.Sheet"
        
        # Mock OMath objects
        mock_omath = Mock()
        mock_omath.Range.Start = 250
        mock_omath.Range.End = 260
        mock_omath.Range.Text = "x^2 + y^2 = z^2"
        
        # Mock document paragraphs
        mock_para1 = Mock()
        mock_para1.Range.Start = 40
        mock_para1.Range.End = 70
        mock_para2 = Mock()
        mock_para2.Range.Start = 140
        mock_para2.Range.End = 170
        mock_para3 = Mock()
        mock_para3.Range.Start = 240
        mock_para3.Range.End = 270
        
        mock_doc.Paragraphs = [mock_para1, mock_para2, mock_para3]
        mock_doc.InlineShapes = [mock_shape1, mock_shape2]
        
        # Mock OMaths
        mock_doc.Range.return_value.OMaths = [mock_omath]
        
        with patch('pythoncom.CoInitialize'), \
             patch('pythoncom.CoUninitialize'), \
             patch('win32com.client.gencache.EnsureDispatch', return_value=mock_app):
            
            with DocumentExtractor() as extractor:
                formulas = extractor._extract_formulas(mock_doc)
                
                assert len(formulas) >= 2  # At least the two inline shapes
                
                # Check equation formula
                equation_formulas = [f for f in formulas if f.formula_type == "equation"]
                assert len(equation_formulas) >= 1
                
                # Check Excel object formula
                excel_formulas = [f for f in formulas if f.formula_type == "excel_object"]
                assert len(excel_formulas) >= 1
    
    def test_enhanced_chart_extraction(self):
        """Test enhanced chart/SmartArt/OLE object extraction."""
        mock_app, mock_doc = self.create_mock_word_app()
        
        # Create mock inline shapes with various object types
        mock_chart = Mock()
        mock_chart.Type = 12  # wdInlineShapeChart
        mock_chart.Range.Start = 50
        mock_chart.Range.End = 60
        mock_chart.Chart.ChartTitle.Text = "Sales Chart"
        
        mock_smartart = Mock()
        mock_smartart.Type = 15  # wdInlineShapeSmartArt
        mock_smartart.Range.Start = 150
        mock_smartart.Range.End = 160
        mock_smartart.AlternativeText = None  # No alternative text
        
        mock_picture = Mock()
        mock_picture.Type = 1  # wdInlineShapePicture
        mock_picture.Range.Start = 250
        mock_picture.Range.End = 260
        mock_picture.AlternativeText = "Company Logo"
        
        # Mock document paragraphs
        mock_para1 = Mock()
        mock_para1.Range.Start = 40
        mock_para1.Range.End = 70
        mock_para2 = Mock()
        mock_para2.Range.Start = 140
        mock_para2.Range.End = 170
        mock_para3 = Mock()
        mock_para3.Range.Start = 240
        mock_para3.Range.End = 270
        
        mock_doc.Paragraphs = [mock_para1, mock_para2, mock_para3]
        mock_doc.InlineShapes = [mock_chart, mock_smartart, mock_picture]
        
        with patch('pythoncom.CoInitialize'), \
             patch('pythoncom.CoUninitialize'), \
             patch('win32com.client.gencache.EnsureDispatch', return_value=mock_app):
            
            with DocumentExtractor() as extractor:
                charts = extractor._extract_charts(mock_doc)
                
                assert len(charts) >= 3
                
                # Check chart
                chart_objects = [c for c in charts if c.chart_type == "chart"]
                assert len(chart_objects) >= 1
                assert chart_objects[0].title == "Sales Chart"
                
                # Check SmartArt
                smartart_objects = [c for c in charts if c.chart_type == "smartart"]
                assert len(smartart_objects) >= 1
                
                # Check picture
                picture_objects = [c for c in charts if c.chart_type == "picture"]
                assert len(picture_objects) >= 1
                assert picture_objects[0].title == "Company Logo"
    
    def test_footnote_extraction(self):
        """Test footnote reference extraction."""
        mock_app, mock_doc = self.create_mock_word_app()
        
        # Create mock footnotes
        mock_footnote1 = Mock()
        mock_footnote1.Index = 1
        mock_footnote1.Reference.Start = 50
        mock_footnote1.Reference.End = 51
        mock_footnote1.Range.Text = "This is the first footnote text."
        
        mock_footnote2 = Mock()
        mock_footnote2.Index = 2
        mock_footnote2.Reference.Start = 150
        mock_footnote2.Reference.End = 151
        mock_footnote2.Range.Text = "This is the second footnote with much longer text that should be truncated to 120 characters maximum length for preview purposes."
        
        # Mock document paragraphs
        mock_para1 = Mock()
        mock_para1.Range.Start = 40
        mock_para1.Range.End = 70
        mock_para2 = Mock()
        mock_para2.Range.Start = 140
        mock_para2.Range.End = 170
        
        mock_doc.Paragraphs = [mock_para1, mock_para2]
        mock_doc.Footnotes = [mock_footnote1, mock_footnote2]
        
        with patch('pythoncom.CoInitialize'), \
             patch('pythoncom.CoUninitialize'), \
             patch('win32com.client.gencache.EnsureDispatch', return_value=mock_app):
            
            with DocumentExtractor() as extractor:
                footnotes = extractor._extract_footnotes(mock_doc)
                
                assert len(footnotes) == 2
                
                # Check first footnote
                assert footnotes[0].footnote_id == "footnote_1"
                assert footnotes[0].reference_mark == "1"
                assert footnotes[0].paragraph_index == 0
                assert footnotes[0].text_preview == "This is the first footnote text."
                
                # Check second footnote (should be truncated)
                assert footnotes[1].footnote_id == "footnote_2"
                assert footnotes[1].reference_mark == "2"
                assert footnotes[1].paragraph_index == 1
                assert len(footnotes[1].text_preview) <= 120
    
    def test_endnote_extraction(self):
        """Test endnote reference extraction."""
        mock_app, mock_doc = self.create_mock_word_app()
        
        # Create mock endnotes
        mock_endnote1 = Mock()
        mock_endnote1.Index = 1
        mock_endnote1.Reference.Start = 50
        mock_endnote1.Reference.End = 51
        mock_endnote1.Range.Text = "This is the first endnote text."
        
        # Mock document paragraphs
        mock_para1 = Mock()
        mock_para1.Range.Start = 40
        mock_para1.Range.End = 70
        
        mock_doc.Paragraphs = [mock_para1]
        mock_doc.Endnotes = [mock_endnote1]
        
        with patch('pythoncom.CoInitialize'), \
             patch('pythoncom.CoUninitialize'), \
             patch('win32com.client.gencache.EnsureDispatch', return_value=mock_app):
            
            with DocumentExtractor() as extractor:
                endnotes = extractor._extract_endnotes(mock_doc)
                
                assert len(endnotes) == 1
                assert endnotes[0].endnote_id == "endnote_1"
                assert endnotes[0].reference_mark == "1"
                assert endnotes[0].paragraph_index == 0
                assert endnotes[0].text_preview == "This is the first endnote text."
    
    def test_cross_reference_extraction(self):
        """Test cross-reference identification and relationship mapping."""
        mock_app, mock_doc = self.create_mock_word_app()
        
        # Create mock cross-reference field
        mock_field = Mock()
        mock_field.Type = 37  # wdFieldRef
        mock_field.Range.Start = 50
        mock_field.Range.End = 60
        mock_field.Code.Text = "REF _Ref123456789 \\h"
        mock_field.Result.Text = "Figure 1"
        
        # Create mock hyperlink
        mock_hyperlink = Mock()
        mock_hyperlink.Range.Start = 150
        mock_hyperlink.Range.End = 160
        mock_hyperlink.TextToDisplay = "See Chapter 2"
        mock_hyperlink.Address = None
        mock_hyperlink.SubAddress = "_Toc987654321"
        
        # Create mock bookmark
        mock_bookmark = Mock()
        mock_bookmark.Name = "_Ref123456789"
        mock_bookmark.Range.Start = 250
        mock_bookmark.Range.End = 260
        
        # Mock document paragraphs
        mock_para1 = Mock()
        mock_para1.Range.Start = 40
        mock_para1.Range.End = 70
        mock_para2 = Mock()
        mock_para2.Range.Start = 140
        mock_para2.Range.End = 170
        mock_para3 = Mock()
        mock_para3.Range.Start = 240
        mock_para3.Range.End = 270
        
        mock_doc.Paragraphs = [mock_para1, mock_para2, mock_para3]
        mock_doc.Fields = [mock_field]
        mock_doc.Hyperlinks = [mock_hyperlink]
        mock_doc.Bookmarks = [mock_bookmark]
        
        with patch('pythoncom.CoInitialize'), \
             patch('pythoncom.CoUninitialize'), \
             patch('win32com.client.gencache.EnsureDispatch', return_value=mock_app):
            
            with DocumentExtractor() as extractor:
                cross_refs = extractor._extract_cross_references(mock_doc)
                
                assert len(cross_refs) >= 2
                
                # Check cross-reference field
                ref_fields = [cr for cr in cross_refs if cr.reference_type == "bookmark"]
                assert len(ref_fields) >= 1
                assert ref_fields[0].source_paragraph_index == 0
                assert ref_fields[0].target_paragraph_index == 2
                assert ref_fields[0].reference_text == "Figure 1"
                assert ref_fields[0].target_id == "_Ref123456789"
                
                # Check internal hyperlink
                internal_links = [cr for cr in cross_refs if cr.reference_type == "internal_link"]
                assert len(internal_links) >= 1
                assert internal_links[0].source_paragraph_index == 1
                assert internal_links[0].reference_text == "See Chapter 2"
                assert internal_links[0].target_id == "_Toc987654321"
    
    def test_enhanced_ooxml_fragment_extraction(self):
        """Test enhanced OOXML fragment extraction with binary object handling."""
        # Create a mock ZIP file structure
        mock_zip_files = {
            'word/document.xml': b'<document>content</document>',
            'word/charts/chart1.xml': b'<chart>chart data</chart>',
            'word/charts/chart1.bin': b'\x00\x01\x02\x03',  # Binary data
            'word/drawings/drawing1.xml': b'<drawing>drawing data</drawing>',
            'word/embeddings/oleObject1.bin': b'\x04\x05\x06\x07',  # Binary data
            'word/activeX/activeX1.xml': b'<activeX>control data</activeX>',
            'customXml/item1.xml': b'<customXml>custom data</customXml>',
            'word/vbaProject.bin': b'\x08\x09\x0A\x0B'  # Binary data
        }
        
        with patch('zipfile.ZipFile') as mock_zipfile:
            mock_zip = Mock()
            mock_zipfile.return_value.__enter__.return_value = mock_zip
            mock_zip.namelist.return_value = list(mock_zip_files.keys())
            
            def mock_read(filename):
                return mock_zip_files[filename]
            
            mock_zip.read = mock_read
            
            extractor = DocumentExtractor()
            fragments = extractor._extract_ooxml_fragments("test.docx")
            
            # Check that XML files are stored as text
            assert 'word/document.xml' in fragments
            assert fragments['word/document.xml'] == '<document>content</document>'
            
            # Check that binary files are stored as base64
            assert 'word/charts/chart1.bin' in fragments
            assert fragments['word/charts/chart1.bin'].startswith('base64:')
            
            # Check that various object types are extracted
            assert 'word/embeddings/oleObject1.bin' in fragments
            assert 'word/vbaProject.bin' in fragments
    
    def test_complete_inventory_extraction(self):
        """Test complete inventory extraction with all complex elements."""
        mock_app, mock_doc = self.create_mock_word_app()
        
        # Set up minimal mocks for all extraction methods
        mock_doc.ContentControls = []
        mock_doc.InlineShapes = []
        mock_doc.Footnotes = []
        mock_doc.Endnotes = []
        mock_doc.Fields = []
        mock_doc.Hyperlinks = []
        mock_doc.Bookmarks = []
        
        with patch('pythoncom.CoInitialize'), \
             patch('pythoncom.CoUninitialize'), \
             patch('win32com.client.gencache.EnsureDispatch', return_value=mock_app), \
             patch.object(DocumentExtractor, '_extract_ooxml_fragments', return_value={}), \
             patch.object(DocumentExtractor, '_extract_media_indexes', return_value={}), \
             patch('os.path.exists', return_value=True):
            
            with DocumentExtractor() as extractor:
                inventory = extractor.extract_inventory(self.test_docx_path)
                
                assert isinstance(inventory, InventoryFullV1)
                assert hasattr(inventory, 'footnotes')
                assert hasattr(inventory, 'endnotes')
                assert hasattr(inventory, 'cross_references')
                assert isinstance(inventory.footnotes, list)
                assert isinstance(inventory.endnotes, list)
                assert isinstance(inventory.cross_references, list)


if __name__ == "__main__":
    pytest.main([__file__])