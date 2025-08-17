"""
Document executor for executing atomic operations through Word COM.

This module implements the DocumentExecutor class that executes atomic operations
through Word COM automation with strict safety controls and localization support.
"""

import os
import re
import shutil
import tempfile
import logging
from typing import Dict, List, Optional, Any, Union
from pathlib import Path

try:
    import win32com.client
    from win32com.client import constants as wdConstants
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False
    wdConstants = None

from ..models import (
    PlanV1, OperationResult, AtomicOperationUnion,
    DeleteSectionByHeading, UpdateToc, DeleteToc, SetStyleRule,
    ReassignParagraphsToStyle, ClearDirectFormatting,
    FontSpec, ParagraphSpec, MatchMode, LineSpacingMode
)
from ..exceptions import ExecutionError, LocalizationError, SecurityViolationError
from ..localization import LocalizationManager
from ..constraints import RuntimeConstraintEnforcer


logger = logging.getLogger(__name__)


# LocalizationManager moved to separate module


class DocumentExecutor:
    """Execute atomic operations through Word COM with strict safety controls."""
    
    def __init__(self, warnings_log_path: Optional[str] = None):
        """
        Initialize document executor.
        
        Args:
            warnings_log_path: Path to warnings.log file for logging fallbacks
        """
        if not WIN32_AVAILABLE:
            raise ExecutionError(
                "Win32 COM is not available. This module requires pywin32.",
                details={"missing_dependency": "pywin32"}
            )
        
        self.localization_manager = LocalizationManager(warnings_log_path)
        self._word_app = None
        
    def execute_plan(self, plan: PlanV1, docx_path: str) -> str:
        """
        Execute complete plan and return modified DOCX path.
        
        Args:
            plan: Execution plan with atomic operations
            docx_path: Path to original DOCX file
            
        Returns:
            str: Path to modified DOCX file
            
        Raises:
            ExecutionError: If execution fails
        """
        if not os.path.exists(docx_path):
            raise ExecutionError(f"DOCX file not found: {docx_path}")
            
        # Create temporary copy for modification
        temp_dir = tempfile.mkdtemp(prefix="autoword_executor_")
        temp_docx = os.path.join(temp_dir, "modified_document.docx")
        shutil.copy2(docx_path, temp_docx)
        
        word_app = None
        doc = None
        
        try:
            # Initialize Word application
            word_app = win32com.client.Dispatch("Word.Application")
            word_app.Visible = False
            word_app.DisplayAlerts = False
            self._word_app = word_app
            
            # Open document
            doc = word_app.Documents.Open(temp_docx)
            
            # Apply localization fallbacks
            self.apply_localization_fallbacks(doc)
            
            # Execute all operations in sequence
            warnings = []
            for i, operation in enumerate(plan.ops):
                try:
                    logger.info(f"Executing operation {i+1}/{len(plan.ops)}: {operation.operation_type}")
                    result = self.execute_operation(operation, doc)
                    
                    if not result.success:
                        operation_data = operation.model_dump() if hasattr(operation, 'model_dump') else str(operation)
                        raise ExecutionError(
                            f"Operation failed: {result.message}",
                            operation_type=operation.operation_type,
                            operation_data=operation_data
                        )
                    
                    warnings.extend(result.warnings)
                    
                except Exception as e:
                    operation_data = operation.model_dump() if hasattr(operation, 'model_dump') else str(operation)
                    raise ExecutionError(
                        f"Failed to execute operation {i+1}: {str(e)}",
                        operation_type=operation.operation_type,
                        operation_data=operation_data
                    ) from e
            
            # Force field updates and repagination
            doc.Fields.Update()
            doc.Repaginate()
            
            # Save document
            doc.Save()
            
            # Write warnings to log if configured
            all_warnings = warnings + self.localization_manager.get_warnings()
            if all_warnings:
                self.localization_manager.write_warnings_log(warnings)
            
            logger.info(f"Plan execution completed successfully. Warnings: {len(all_warnings)}")
            return temp_docx
            
        except Exception as e:
            # Clean up temporary file on error
            if os.path.exists(temp_docx):
                try:
                    os.remove(temp_docx)
                except:
                    pass
            raise
            
        finally:
            # Clean up COM objects
            if doc:
                try:
                    doc.Close(SaveChanges=False)
                except:
                    pass
            if word_app:
                try:
                    word_app.Quit()
                except:
                    pass
            self._word_app = None
    
    def execute_operation(self, operation: AtomicOperationUnion, doc: object) -> OperationResult:
        """
        Execute single atomic operation.
        
        Args:
            operation: Atomic operation to execute
            doc: Word document COM object
            
        Returns:
            OperationResult: Result of operation execution
            
        Raises:
            ExecutionError: If operation fails
        """
        operation_type = operation.operation_type
        warnings = []
        
        try:
            if isinstance(operation, DeleteSectionByHeading) or operation_type == "delete_section_by_heading":
                success, message = self._delete_section_by_heading(operation, doc, warnings)
            elif isinstance(operation, UpdateToc) or operation_type == "update_toc":
                success, message = self._update_toc(operation, doc, warnings)
            elif isinstance(operation, DeleteToc) or operation_type == "delete_toc":
                success, message = self._delete_toc(operation, doc, warnings)
            elif isinstance(operation, SetStyleRule) or operation_type == "set_style_rule":
                success, message = self._set_style_rule(operation, doc, warnings)
            elif isinstance(operation, ReassignParagraphsToStyle) or operation_type == "reassign_paragraphs_to_style":
                success, message = self._reassign_paragraphs_to_style(operation, doc, warnings)
            elif isinstance(operation, ClearDirectFormatting) or operation_type == "clear_direct_formatting":
                success, message = self._clear_direct_formatting(operation, doc, warnings)
            else:
                raise SecurityViolationError(
                    f"Unknown operation type: {operation_type}",
                    operation_type=operation_type,
                    security_context={"whitelist_violation": True}
                )
            
            return OperationResult(
                success=success,
                operation_type=operation_type,
                message=message,
                warnings=warnings
            )
            
        except Exception as e:
            return OperationResult(
                success=False,
                operation_type=operation_type,
                message=f"Operation failed: {str(e)}",
                warnings=warnings
            )
    
    def apply_localization_fallbacks(self, doc: object):
        """
        Apply style aliases and font fallbacks.
        
        Args:
            doc: Word document COM object
            
        Raises:
            LocalizationError: If localization fails
        """
        try:
            # Detect document styles for better mapping
            style_mapping = self.localization_manager.detect_document_styles(doc)
            
            # Apply font fallbacks to all styles
            for style in doc.Styles:
                try:
                    # Check East Asian font
                    if hasattr(style.Font, 'NameFarEast'):
                        original_font = style.Font.NameFarEast
                        if original_font:
                            resolved_font = self.localization_manager.resolve_font_name(original_font, doc)
                            if resolved_font != original_font:
                                style.Font.NameFarEast = resolved_font
                    
                    # Check Latin font
                    if hasattr(style.Font, 'Name'):
                        original_font = style.Font.Name
                        if original_font:
                            resolved_font = self.localization_manager.resolve_font_name(original_font, doc)
                            if resolved_font != original_font:
                                style.Font.Name = resolved_font
                                
                except Exception as e:
                    self.localization_manager._log_warning(f"Failed to apply font fallback to style {getattr(style, 'NameLocal', 'unknown')}: {str(e)}")
                    
        except Exception as e:
            raise LocalizationError(
                f"Failed to apply localization fallbacks: {str(e)}"
            ) from e
    
    def _delete_section_by_heading(self, operation, doc: object, warnings: List[str]) -> tuple[bool, str]:
        """Delete section by heading text."""
        heading_text = getattr(operation, 'heading_text', '')
        level = getattr(operation, 'level', 1)
        match_mode = getattr(operation, 'match', MatchMode.EXACT)
        case_sensitive = getattr(operation, 'case_sensitive', False)
        occurrence_index = getattr(operation, 'occurrence_index', None)
        
        try:
            # Find all headings that match criteria
            matching_headings = []
            
            for para in doc.Paragraphs:
                if para.OutlineLevel == level:
                    para_text = para.Range.Text.strip()
                    
                    # Apply matching logic
                    is_match = False
                    if match_mode == MatchMode.EXACT:
                        is_match = (para_text == heading_text) if case_sensitive else (para_text.lower() == heading_text.lower())
                    elif match_mode == MatchMode.CONTAINS:
                        is_match = (heading_text in para_text) if case_sensitive else (heading_text.lower() in para_text.lower())
                    elif match_mode == MatchMode.REGEX:
                        flags = 0 if case_sensitive else re.IGNORECASE
                        is_match = bool(re.search(heading_text, para_text, flags))
                    
                    if is_match:
                        matching_headings.append(para)
            
            if not matching_headings:
                warnings.append(f"NOOP: No heading found matching '{heading_text}' at level {level}")
                return True, f"No matching heading found (NOOP)"
            
            # Select specific occurrence if specified
            if occurrence_index is not None:
                if occurrence_index > len(matching_headings):
                    warnings.append(f"NOOP: Occurrence index {occurrence_index} exceeds matches ({len(matching_headings)})")
                    return True, f"Occurrence index out of range (NOOP)"
                target_heading = matching_headings[occurrence_index - 1]
            else:
                target_heading = matching_headings[0]
                if len(matching_headings) > 1:
                    warnings.append(f"Multiple headings matched, using first occurrence")
            
            # Find the range to delete (from heading to next same-level heading)
            start_range = target_heading.Range.Start
            end_range = doc.Range().End
            
            # Look for next heading at same or higher level
            current_para = target_heading.Next()
            while current_para:
                if current_para.OutlineLevel <= level:
                    end_range = current_para.Range.Start
                    break
                current_para = current_para.Next()
            
            # Delete the range
            delete_range = doc.Range(start_range, end_range)
            delete_range.Delete()
            
            return True, f"Deleted section starting with heading '{heading_text}'"
            
        except Exception as e:
            operation_data = operation.model_dump() if hasattr(operation, 'model_dump') else str(operation)
            raise ExecutionError(
                f"Failed to delete section by heading: {str(e)}",
                operation_type="delete_section_by_heading",
                operation_data=operation_data
            ) from e
    
    def _update_toc(self, operation, doc: object, warnings: List[str]) -> tuple[bool, str]:
        """Update table of contents."""
        try:
            toc_updated = False
            
            # Find and update all TOC fields
            for field in doc.Fields:
                if field.Type == wdConstants.wdFieldTOC:
                    field.Update()
                    toc_updated = True
            
            if not toc_updated:
                warnings.append("NOOP: No TOC fields found to update")
                return True, "No TOC found (NOOP)"
            
            return True, "TOC updated successfully"
            
        except Exception as e:
            raise ExecutionError(
                f"Failed to update TOC: {str(e)}",
                operation_type="update_toc"
            ) from e
    
    def _delete_toc(self, operation, doc: object, warnings: List[str]) -> tuple[bool, str]:
        """Delete table of contents."""
        mode = getattr(operation, 'mode', 'all')
        
        try:
            toc_fields = []
            
            # Find all TOC fields
            for field in doc.Fields:
                if field.Type == wdConstants.wdFieldTOC:
                    toc_fields.append(field)
            
            if not toc_fields:
                warnings.append("NOOP: No TOC fields found to delete")
                return True, "No TOC found (NOOP)"
            
            # Delete based on mode
            if mode == "all":
                for field in toc_fields:
                    field.Delete()
                return True, f"Deleted {len(toc_fields)} TOC field(s)"
            elif mode == "first":
                toc_fields[0].Delete()
                return True, "Deleted first TOC field"
            elif mode == "last":
                toc_fields[-1].Delete()
                return True, "Deleted last TOC field"
            else:
                raise ExecutionError(f"Invalid TOC deletion mode: {mode}")
                
        except Exception as e:
            operation_data = operation.model_dump() if hasattr(operation, 'model_dump') else str(operation)
            raise ExecutionError(
                f"Failed to delete TOC: {str(e)}",
                operation_type="delete_toc",
                operation_data=operation_data
            ) from e
    
    def _set_style_rule(self, operation, doc: object, warnings: List[str]) -> tuple[bool, str]:
        """Set style rule for font and paragraph formatting."""
        target_style_name = getattr(operation, 'target_style_name', '')
        font_spec = getattr(operation, 'font', None)
        paragraph_spec = getattr(operation, 'paragraph', None)
        
        try:
            # Resolve style name using localization
            resolved_style_name = self.localization_manager.resolve_style_name(target_style_name, doc)
            
            # Get or create style
            try:
                style = doc.Styles[resolved_style_name]
            except:
                # Style doesn't exist, create it
                style = doc.Styles.Add(resolved_style_name, wdConstants.wdStyleTypeParagraph)
                warnings.append(f"Created new style: {resolved_style_name}")
            
            # Apply font specifications
            if font_spec:
                if font_spec.east_asian:
                    resolved_font = self.localization_manager.resolve_font_name(font_spec.east_asian, doc)
                    style.Font.NameFarEast = resolved_font
                
                if font_spec.latin:
                    resolved_font = self.localization_manager.resolve_font_name(font_spec.latin, doc)
                    style.Font.Name = resolved_font
                
                if font_spec.size_pt is not None:
                    style.Font.Size = font_spec.size_pt
                
                if font_spec.bold is not None:
                    style.Font.Bold = font_spec.bold
                
                if font_spec.italic is not None:
                    style.Font.Italic = font_spec.italic
                
                if font_spec.color_hex:
                    # Convert hex color to RGB
                    hex_color = font_spec.color_hex.lstrip('#')
                    rgb = int(hex_color, 16)
                    style.Font.Color = rgb
            
            # Apply paragraph specifications
            if paragraph_spec:
                if paragraph_spec.line_spacing_mode and paragraph_spec.line_spacing_value is not None:
                    if paragraph_spec.line_spacing_mode == LineSpacingMode.SINGLE:
                        style.ParagraphFormat.LineSpacingRule = wdConstants.wdLineSpaceSingle
                    elif paragraph_spec.line_spacing_mode == LineSpacingMode.MULTIPLE:
                        style.ParagraphFormat.LineSpacingRule = wdConstants.wdLineSpaceMultiple
                        style.ParagraphFormat.LineSpacing = paragraph_spec.line_spacing_value
                    elif paragraph_spec.line_spacing_mode == LineSpacingMode.EXACTLY:
                        style.ParagraphFormat.LineSpacingRule = wdConstants.wdLineSpaceExactly
                        style.ParagraphFormat.LineSpacing = paragraph_spec.line_spacing_value
                
                if paragraph_spec.space_before_pt is not None:
                    style.ParagraphFormat.SpaceBefore = paragraph_spec.space_before_pt
                
                if paragraph_spec.space_after_pt is not None:
                    style.ParagraphFormat.SpaceAfter = paragraph_spec.space_after_pt
                
                if paragraph_spec.indent_left_pt is not None:
                    style.ParagraphFormat.LeftIndent = paragraph_spec.indent_left_pt
                
                if paragraph_spec.indent_right_pt is not None:
                    style.ParagraphFormat.RightIndent = paragraph_spec.indent_right_pt
                
                if paragraph_spec.indent_first_line_pt is not None:
                    style.ParagraphFormat.FirstLineIndent = paragraph_spec.indent_first_line_pt
            
            return True, f"Style rule applied to '{resolved_style_name}'"
            
        except Exception as e:
            operation_data = operation.model_dump() if hasattr(operation, 'model_dump') else str(operation)
            raise ExecutionError(
                f"Failed to set style rule: {str(e)}",
                operation_type="set_style_rule",
                operation_data=operation_data
            ) from e
    
    def _reassign_paragraphs_to_style(self, operation, doc: object, warnings: List[str]) -> tuple[bool, str]:
        """Reassign paragraphs to target style."""
        selector = getattr(operation, 'selector', {})
        target_style_name = getattr(operation, 'target_style_name', '')
        clear_direct_formatting = getattr(operation, 'clear_direct_formatting', False)
        
        try:
            # Resolve target style name
            resolved_target_style = self.localization_manager.resolve_style_name(target_style_name, doc)
            
            # Find paragraphs matching selector
            matching_paragraphs = []
            
            for para in doc.Paragraphs:
                is_match = True
                
                # Apply selector criteria
                if "style_name" in selector:
                    current_style = para.Style.NameLocal
                    expected_style = self.localization_manager.resolve_style_name(selector["style_name"], doc)
                    if current_style != expected_style:
                        is_match = False
                
                if "outline_level" in selector:
                    if para.OutlineLevel != selector["outline_level"]:
                        is_match = False
                
                if "text_contains" in selector:
                    para_text = para.Range.Text.strip()
                    if selector["text_contains"] not in para_text:
                        is_match = False
                
                if "text_regex" in selector:
                    para_text = para.Range.Text.strip()
                    if not re.search(selector["text_regex"], para_text):
                        is_match = False
                
                if is_match:
                    matching_paragraphs.append(para)
            
            if not matching_paragraphs:
                warnings.append("NOOP: No paragraphs found matching selector criteria")
                return True, "No matching paragraphs found (NOOP)"
            
            # Reassign paragraphs to target style
            for para in matching_paragraphs:
                para.Style = resolved_target_style
                
                if clear_direct_formatting:
                    para.Range.ClearFormatting()
            
            return True, f"Reassigned {len(matching_paragraphs)} paragraph(s) to style '{resolved_target_style}'"
            
        except Exception as e:
            operation_data = operation.model_dump() if hasattr(operation, 'model_dump') else str(operation)
            raise ExecutionError(
                f"Failed to reassign paragraphs to style: {str(e)}",
                operation_type="reassign_paragraphs_to_style",
                operation_data=operation_data
            ) from e
    
    def _clear_direct_formatting(self, operation, doc: object, warnings: List[str]) -> tuple[bool, str]:
        """Clear direct formatting with explicit authorization."""
        scope = operation.scope
        range_spec = getattr(operation, 'range_spec', None)
        
        # Security check - this operation requires explicit authorization
        if not getattr(operation, 'authorization_required', True):
            raise SecurityViolationError(
                "clear_direct_formatting requires explicit authorization",
                operation_type="clear_direct_formatting",
                security_context={"authorization_missing": True}
            )
        
        try:
            if scope == "document":
                # Clear formatting for entire document
                doc.Range().ClearFormatting()
                return True, "Cleared direct formatting for entire document"
                
            elif scope == "selection":
                # Clear formatting for current selection (if any)
                if doc.Application.Selection.Range.Text:
                    doc.Application.Selection.Range.ClearFormatting()
                    return True, "Cleared direct formatting for selection"
                else:
                    warnings.append("NOOP: No selection to clear formatting")
                    return True, "No selection found (NOOP)"
                    
            elif scope == "range":
                # Clear formatting for specified range
                if not range_spec:
                    raise ExecutionError("Range specification required for range scope")
                
                start = range_spec.get("start", 0)
                end = range_spec.get("end", doc.Range().End)
                
                range_obj = doc.Range(start, end)
                range_obj.ClearFormatting()
                
                return True, f"Cleared direct formatting for range {start}-{end}"
            
            else:
                raise ExecutionError(f"Invalid scope for clear_direct_formatting: {scope}")
                
        except Exception as e:
            operation_data = operation.model_dump() if hasattr(operation, 'model_dump') else str(operation)
            raise ExecutionError(
                f"Failed to clear direct formatting: {str(e)}",
                operation_type="clear_direct_formatting",
                operation_data=operation_data
            ) from e