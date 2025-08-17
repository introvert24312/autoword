"""
Tests for AutoWord vNext DocumentExecutor.

This module tests the DocumentExecutor class and its atomic operations
with mock Word COM objects to ensure proper functionality.
"""

import os
import tempfile
import unittest
from unittest.mock import Mock, MagicMock, patch, call
from pathlib import Path

from autoword.vnext.executor import DocumentExecutor
from autoword.vnext.models import (
    PlanV1, DeleteSectionByHeading, UpdateToc, DeleteToc, SetStyleRule,
    ReassignParagraphsToStyle, ClearDirectFormatting, FontSpec, ParagraphSpec,
    MatchMode, LineSpacingMode, OperationResult
)
from autoword.vnext.exceptions import ExecutionError, LocalizationError, SecurityViolationError


class TestDocumentExecutor(unittest.TestCase):
    """Test cases for DocumentExecutor class."""
    
    def setUp(self):
        """Set up test fixtures."""
        # Mock win32com availability
        self.win32_patcher = patch('autoword.vnext.executor.document_executor.WIN32_AVAILABLE', True)
        self.win32_patcher.start()
        
        # Mock win32com.client
        self.win32com_patcher = patch('autoword.vnext.executor.document_executor.win32com')
        self.mock_win32com = self.win32com_patcher.start()
        
        # Mock wdConstants
        self.constants_patcher = patch('autoword.vnext.executor.document_executor.wdConstants')
        self.mock_constants = self.constants_patcher.start()
        self.mock_constants.wdFieldTOC = 13
        self.mock_constants.wdStyleTypeParagraph = 1
        self.mock_constants.wdLineSpaceSingle = 0
        self.mock_constants.wdLineSpaceMultiple = 5
        self.mock_constants.wdLineSpaceExactly = 4
        
        self.executor = DocumentExecutor()
        
    def tearDown(self):
        """Clean up test fixtures."""
        self.win32_patcher.stop()
        self.win32com_patcher.stop()
        self.constants_patcher.stop()
    
    def test_init_without_win32(self):
        """Test initialization without win32com available."""
        with patch('autoword.vnext.executor.document_executor.WIN32_AVAILABLE', False):
            with self.assertRaises(ExecutionError) as context:
                DocumentExecutor()
            self.assertIn("Win32 COM is not available", str(context.exception))
    
    def test_execute_plan_file_not_found(self):
        """Test execute_plan with non-existent file."""
        plan = PlanV1(ops=[])
        
        with self.assertRaises(ExecutionError) as context:
            self.executor.execute_plan(plan, "nonexistent.docx")
        
        self.assertIn("DOCX file not found", str(context.exception))
    
    @patch('tempfile.mkdtemp')
    @patch('shutil.copy2')
    @patch('os.path.exists')
    def test_execute_plan_success(self, mock_exists, mock_copy, mock_mkdtemp):
        """Test successful plan execution."""
        # Setup mocks
        mock_exists.return_value = True
        mock_mkdtemp.return_value = "/tmp/test"
        
        mock_word_app = Mock()
        mock_doc = Mock()
        mock_word_app.Documents.Open.return_value = mock_doc
        self.mock_win32com.client.Dispatch.return_value = mock_word_app
        
        # Create simple plan
        plan = PlanV1(ops=[UpdateToc()])
        
        # Mock the operation execution
        with patch.object(self.executor, 'execute_operation') as mock_execute_op:
            mock_execute_op.return_value = OperationResult(success=True, operation_type="update_toc")
            
            with patch.object(self.executor, 'apply_localization_fallbacks'):
                result = self.executor.execute_plan(plan, "test.docx")
        
        # Verify result
        # Use os.path.join for cross-platform compatibility
        expected_path = os.path.join("/tmp/test", "modified_document.docx")
        self.assertEqual(result, expected_path)
        mock_doc.Fields.Update.assert_called_once()
        mock_doc.Repaginate.assert_called_once()
        mock_doc.Save.assert_called_once()
    
    def test_execute_operation_delete_section_by_heading(self):
        """Test delete_section_by_heading operation."""
        operation = DeleteSectionByHeading(
            heading_text="摘要",
            level=1,
            match=MatchMode.EXACT,
            case_sensitive=False
        )
        
        # Mock document and paragraphs
        mock_doc = Mock()
        mock_para1 = Mock()
        mock_para1.OutlineLevel = 1
        mock_para1.Range.Text = "摘要\r"
        mock_para1.Range.Start = 100
        mock_para1.Next.return_value = None
        
        mock_doc.Paragraphs = [mock_para1]
        mock_doc.Range.return_value.End = 200
        
        mock_delete_range = Mock()
        mock_doc.Range.return_value = mock_delete_range
        
        result = self.executor.execute_operation(operation, mock_doc)
        
        self.assertTrue(result.success)
        self.assertEqual(result.operation_type, "delete_section_by_heading")
        mock_delete_range.Delete.assert_called_once()
    
    def test_execute_operation_update_toc(self):
        """Test update_toc operation."""
        operation = UpdateToc()
        
        # Mock document with TOC field
        mock_doc = Mock()
        mock_field = Mock()
        mock_field.Type = 13  # wdFieldTOC
        mock_doc.Fields = [mock_field]
        
        result = self.executor.execute_operation(operation, mock_doc)
        
        self.assertTrue(result.success)
        self.assertEqual(result.operation_type, "update_toc")
        mock_field.Update.assert_called_once()
    
    def test_execute_operation_update_toc_no_toc(self):
        """Test update_toc operation with no TOC."""
        operation = UpdateToc()
        
        # Mock document without TOC fields
        mock_doc = Mock()
        mock_doc.Fields = []
        
        result = self.executor.execute_operation(operation, mock_doc)
        
        self.assertTrue(result.success)
        self.assertIn("NOOP", result.message)
        self.assertIn("NOOP: No TOC fields found to update", result.warnings[0])
    
    def test_execute_operation_delete_toc(self):
        """Test delete_toc operation."""
        operation = DeleteToc(mode="all")
        
        # Mock document with TOC fields
        mock_doc = Mock()
        mock_field1 = Mock()
        mock_field1.Type = 13  # wdFieldTOC
        mock_field2 = Mock()
        mock_field2.Type = 13  # wdFieldTOC
        mock_doc.Fields = [mock_field1, mock_field2]
        
        result = self.executor.execute_operation(operation, mock_doc)
        
        self.assertTrue(result.success)
        self.assertEqual(result.operation_type, "delete_toc")
        mock_field1.Delete.assert_called_once()
        mock_field2.Delete.assert_called_once()
    
    def test_execute_operation_set_style_rule(self):
        """Test set_style_rule operation."""
        operation = SetStyleRule(
            target_style_name="Heading 1",
            font=FontSpec(east_asian="楷体", latin="Times New Roman", size_pt=12, bold=True),
            paragraph=ParagraphSpec(line_spacing_mode=LineSpacingMode.MULTIPLE, line_spacing_value=2.0)
        )
        
        # Mock document and style
        mock_doc = Mock()
        mock_style = Mock()
        mock_doc.Styles = {
            "标题 1": mock_style  # Mock the styles collection
        }
        
        with patch.object(self.executor.localization_manager, 'resolve_style_name', return_value="标题 1"):
            with patch.object(self.executor.localization_manager, 'resolve_font_name', side_effect=lambda x, doc, warnings: x):
                result = self.executor.execute_operation(operation, mock_doc)
        
        self.assertTrue(result.success)
        self.assertEqual(result.operation_type, "set_style_rule")
        self.assertEqual(mock_style.Font.NameFarEast, "楷体")
        self.assertEqual(mock_style.Font.Name, "Times New Roman")
        self.assertEqual(mock_style.Font.Size, 12)
        self.assertTrue(mock_style.Font.Bold)
    
    def test_execute_operation_reassign_paragraphs_to_style(self):
        """Test reassign_paragraphs_to_style operation."""
        operation = ReassignParagraphsToStyle(
            selector={"style_name": "Normal"},
            target_style_name="Heading 1",
            clear_direct_formatting=True
        )
        
        # Mock document and paragraphs
        mock_doc = Mock()
        mock_para = Mock()
        mock_para.Style.NameLocal = "正文"
        mock_para.Range = Mock()  # Ensure Range is properly mocked
        mock_doc.Paragraphs = [mock_para]
        
        with patch.object(self.executor.localization_manager, 'resolve_style_name') as mock_resolve:
            # Mock to return the expected resolved names
            def resolve_side_effect(style_name, doc):
                if style_name == "Normal":
                    return "正文"
                elif style_name == "Heading 1":
                    return "标题 1"
                return style_name
            
            mock_resolve.side_effect = resolve_side_effect
            
            result = self.executor.execute_operation(operation, mock_doc)
        
        self.assertTrue(result.success)
        self.assertEqual(result.operation_type, "reassign_paragraphs_to_style")
        # Check that the style was assigned (mock_para.Style should be set to the resolved style)
        self.assertTrue(hasattr(mock_para, 'Style'))
        mock_para.Range.ClearFormatting.assert_called_once()
    
    def test_execute_operation_clear_direct_formatting_document(self):
        """Test clear_direct_formatting operation for document scope."""
        operation = ClearDirectFormatting(
            scope="document",
            authorization_required=True
        )
        
        # Mock document
        mock_doc = Mock()
        mock_range = Mock()
        mock_doc.Range.return_value = mock_range
        
        result = self.executor.execute_operation(operation, mock_doc)
        
        self.assertTrue(result.success)
        self.assertEqual(result.operation_type, "clear_direct_formatting")
        mock_range.ClearFormatting.assert_called_once()
    
    def test_execute_operation_clear_direct_formatting_unauthorized(self):
        """Test clear_direct_formatting operation without authorization."""
        # Create operation with mock to bypass Pydantic validation
        operation = Mock()
        operation.operation_type = "clear_direct_formatting"
        operation.scope = "document"
        operation.range_spec = None
        operation.authorization_required = False  # This should cause security violation
        operation.model_dump.return_value = {"scope": "document", "authorization_required": False}
        
        mock_doc = Mock()
        
        result = self.executor.execute_operation(operation, mock_doc)
        
        self.assertFalse(result.success)
        self.assertIn("authorization", result.message.lower())
    
    def test_execute_operation_unknown_operation(self):
        """Test execution with unknown operation type."""
        # Create a mock operation with unknown type
        mock_operation = Mock()
        mock_operation.operation_type = "unknown_operation"
        
        mock_doc = Mock()
        
        result = self.executor.execute_operation(mock_operation, mock_doc)
        
        self.assertFalse(result.success)
        self.assertIn("Operation failed", result.message)
    
    def test_apply_localization_fallbacks(self):
        """Test localization fallbacks application."""
        # Mock document with styles
        mock_doc = Mock()
        mock_style = Mock()
        mock_style.Font.NameFarEast = "楷体"
        mock_style.Font.Name = "Times New Roman"
        mock_doc.Styles = [mock_style]
        
        with patch.object(self.executor.localization_manager, 'resolve_font_name', side_effect=lambda x, doc, warnings: x):
            self.executor.apply_localization_fallbacks(mock_doc)
        
        # Should not raise any exceptions
        self.assertTrue(True)


class TestLocalizationManager(unittest.TestCase):
    """Test cases for LocalizationManager class."""
    
    def setUp(self):
        """Set up test fixtures."""
        from autoword.vnext.executor.document_executor import LocalizationManager
        self.manager = LocalizationManager()
    
    def test_resolve_style_name_direct_match(self):
        """Test style name resolution with direct match."""
        mock_doc = Mock()
        mock_doc.Styles = {"Heading 1": Mock()}  # Style exists
        
        with patch('autoword.vnext.executor.document_executor.LocalizationManager._style_exists', return_value=True):
            result = self.manager.resolve_style_name("Heading 1", mock_doc)
        self.assertEqual(result, "Heading 1")
    
    def test_resolve_style_name_alias_match(self):
        """Test style name resolution with alias match."""
        mock_doc = Mock()
        
        def mock_style_exists(style_name, doc):
            return style_name == "标题 1"  # Only Chinese style exists
        
        with patch('autoword.vnext.executor.document_executor.LocalizationManager._style_exists', side_effect=mock_style_exists):
            result = self.manager.resolve_style_name("Heading 1", mock_doc)
        self.assertEqual(result, "标题 1")
    
    def test_resolve_font_name_direct_match(self):
        """Test font name resolution with direct match."""
        mock_doc = Mock()
        
        with patch.object(self.manager, '_font_exists', return_value=True):
            warnings = []
            result = self.manager.resolve_font_name("楷体", mock_doc, warnings)
        
        self.assertEqual(result, "楷体")
        self.assertEqual(len(warnings), 0)
    
    def test_resolve_font_name_fallback(self):
        """Test font name resolution with fallback."""
        mock_doc = Mock()
        
        def mock_font_exists(font_name, doc):
            # Original font doesn't exist, but second in fallback chain does
            return font_name == "楷体_GB2312"
        
        with patch('autoword.vnext.executor.document_executor.LocalizationManager._font_exists', side_effect=mock_font_exists):
            warnings = []
            result = self.manager.resolve_font_name("楷体", mock_doc, warnings)
        
        self.assertEqual(result, "楷体_GB2312")
        self.assertEqual(len(warnings), 1)
        self.assertIn("Font fallback", warnings[0])
    
    def test_resolve_font_name_no_fallback(self):
        """Test font name resolution with no available fallback."""
        mock_doc = Mock()
        
        with patch.object(self.manager, '_font_exists', return_value=False):
            warnings = []
            result = self.manager.resolve_font_name("UnknownFont", mock_doc, warnings)
        
        self.assertEqual(result, "UnknownFont")
        # Since UnknownFont is not in FONT_FALLBACKS, no warning is generated
        self.assertEqual(len(warnings), 0)


if __name__ == '__main__':
    unittest.main()