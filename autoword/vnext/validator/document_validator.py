"""
Document validator for comprehensive validation and rollback capabilities.

This module implements the Validator component of the AutoWord vNext pipeline,
responsible for validating document modifications against strict assertions
and providing rollback functionality when validation fails.
"""

import os
import shutil
import logging
from typing import List, Optional, Dict, Any
from pathlib import Path

import pythoncom
import win32com.client as win32

from ..models import StructureV1, ValidationResult, StyleDefinition, FontSpec, ParagraphSpec, LineSpacingMode
from ..exceptions import ValidationError, RollbackError
from ..extractor.document_extractor import DocumentExtractor


logger = logging.getLogger(__name__)


class DocumentValidator:
    """Validate document modifications against strict assertions with rollback capability."""
    
    def __init__(self, visible: bool = False):
        """
        Initialize document validator.
        
        Args:
            visible: Whether to show Word application window during validation
        """
        self.visible = visible
        self._word_app = None
        self._com_initialized = False
    
    def __enter__(self):
        """Enter context manager for COM resource management."""
        try:
            # Initialize COM
            pythoncom.CoInitialize()
            self._com_initialized = True
            logger.debug("COM initialized for validation")
            
            # Create Word application instance
            self._word_app = win32.gencache.EnsureDispatch('Word.Application')
            self._word_app.Visible = self.visible
            self._word_app.DisplayAlerts = 0  # Disable alerts
            self._word_app.ScreenUpdating = False  # Improve performance
            
            logger.info(f"Word application started for validation (visible={self.visible})")
            return self
            
        except Exception as e:
            self._cleanup()
            raise ValidationError(f"Failed to initialize Word COM for validation: {e}")
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Exit context manager and cleanup resources."""
        self._cleanup()
    
    def _cleanup(self):
        """Clean up COM resources."""
        try:
            if self._word_app:
                self._word_app.ScreenUpdating = True
                self._word_app.DisplayAlerts = -1
                self._word_app.Quit(SaveChanges=0)
                self._word_app = None
                logger.info("Word application closed after validation")
        except Exception as e:
            logger.warning(f"Error during Word cleanup: {e}")
        
        finally:
            if self._com_initialized:
                try:
                    pythoncom.CoUninitialize()
                    self._com_initialized = False
                    logger.debug("COM uninitialized after validation")
                except Exception as e:
                    logger.warning(f"Error during COM cleanup: {e}")
    
    def validate_modifications(self, original_structure: StructureV1, modified_docx: str, 
                             original_docx: Optional[str] = None) -> ValidationResult:
        """
        Validate all assertions and generate comparison structure.
        
        Args:
            original_structure: Original document structure
            modified_docx: Path to modified DOCX file
            original_docx: Path to original DOCX file for rollback
            
        Returns:
            ValidationResult: Validation result with detailed errors
            
        Raises:
            ValidationError: If validation process fails
        """
        if not os.path.exists(modified_docx):
            raise ValidationError(f"Modified DOCX file not found: {modified_docx}")
        
        try:
            logger.info(f"Validating modifications in: {modified_docx}")
            
            # First, update fields and repaginate the document
            self._update_fields_and_repaginate(modified_docx)
            
            # Extract structure from modified document
            with DocumentExtractor(visible=self.visible) as extractor:
                modified_structure = extractor.extract_structure(modified_docx)
            
            # Collect all validation errors
            all_errors = []
            all_warnings = []
            
            # Run all assertion checks
            chapter_errors = self.check_chapter_assertions(modified_structure)
            all_errors.extend(chapter_errors)
            
            style_errors = self.check_style_assertions(modified_structure)
            all_errors.extend(style_errors)
            
            toc_errors = self.check_toc_assertions(modified_structure)
            all_errors.extend(toc_errors)
            
            pagination_errors = self.check_pagination_assertions(original_structure, modified_structure)
            all_errors.extend(pagination_errors)
            
            # Check if validation passed
            is_valid = len(all_errors) == 0
            
            if not is_valid:
                logger.warning(f"Validation failed with {len(all_errors)} errors")
                
                # Perform rollback if original DOCX is provided
                if original_docx:
                    try:
                        self.rollback_document(original_docx, modified_docx)
                        all_warnings.append(f"Document rolled back to original: {original_docx}")
                    except RollbackError as e:
                        all_errors.append(f"Rollback failed: {e}")
            else:
                logger.info("All validation assertions passed")
            
            return ValidationResult(
                is_valid=is_valid,
                errors=all_errors,
                warnings=all_warnings
            )
            
        except Exception as e:
            if isinstance(e, ValidationError):
                raise
            raise ValidationError(
                f"Failed to validate modifications: {e}",
                validation_stage="validation"
            )
    
    def check_chapter_assertions(self, structure: StructureV1) -> List[str]:
        """
        Verify no 摘要/参考文献 at level 1.
        
        Args:
            structure: Document structure to validate
            
        Returns:
            List[str]: List of assertion failures
        """
        errors = []
        
        try:
            # Check for forbidden headings at level 1
            forbidden_headings = ["摘要", "参考文献", "Abstract", "References"]
            
            for heading in structure.headings:
                if heading.level == 1:
                    heading_text = heading.text.strip()
                    
                    # Check for exact matches and partial matches
                    for forbidden in forbidden_headings:
                        if forbidden.lower() in heading_text.lower():
                            errors.append(
                                f"Chapter assertion failed: Found forbidden heading '{heading_text}' "
                                f"at level 1 (paragraph {heading.paragraph_index}). "
                                f"'{forbidden}' should not be at level 1."
                            )
                            break
            
            logger.info(f"Chapter assertions checked: {len(errors)} failures found")
            
        except Exception as e:
            logger.error(f"Error during chapter assertion checking: {e}")
            errors.append(f"Chapter assertion check failed: {e}")
        
        return errors
    
    def check_style_assertions(self, structure: StructureV1) -> List[str]:
        """
        Verify H1/H2/Normal style specifications.
        
        Expected specifications:
        - H1 (Heading 1/标题 1): 楷体, 12pt, bold, 2.0 line spacing
        - H2 (Heading 2/标题 2): 宋体, 12pt, bold, 2.0 line spacing  
        - Normal (正文): 宋体, 12pt, 2.0 line spacing
        
        Args:
            structure: Document structure to validate
            
        Returns:
            List[str]: List of assertion failures
        """
        errors = []
        
        try:
            # Define expected style specifications
            expected_styles = {
                "Heading 1": {
                    "aliases": ["Heading 1", "标题 1"],
                    "font_east_asian": "楷体",
                    "font_size_pt": 12,
                    "font_bold": True,
                    "line_spacing_mode": LineSpacingMode.MULTIPLE,
                    "line_spacing_value": 2.0
                },
                "Heading 2": {
                    "aliases": ["Heading 2", "标题 2"],
                    "font_east_asian": "宋体",
                    "font_size_pt": 12,
                    "font_bold": True,
                    "line_spacing_mode": LineSpacingMode.MULTIPLE,
                    "line_spacing_value": 2.0
                },
                "Normal": {
                    "aliases": ["Normal", "正文"],
                    "font_east_asian": "宋体",
                    "font_size_pt": 12,
                    "font_bold": False,
                    "line_spacing_mode": LineSpacingMode.MULTIPLE,
                    "line_spacing_value": 2.0
                }
            }
            
            # Check each expected style
            for style_key, expected in expected_styles.items():
                style_found = False
                
                for style_def in structure.styles:
                    if style_def.name in expected["aliases"]:
                        style_found = True
                        
                        # Check font specifications
                        if style_def.font:
                            font = style_def.font
                            
                            # Check East Asian font
                            if font.east_asian != expected["font_east_asian"]:
                                errors.append(
                                    f"Style assertion failed: {style_def.name} font_east_asian is "
                                    f"'{font.east_asian}', expected '{expected['font_east_asian']}'"
                                )
                            
                            # Check font size
                            if font.size_pt != expected["font_size_pt"]:
                                errors.append(
                                    f"Style assertion failed: {style_def.name} font_size_pt is "
                                    f"{font.size_pt}, expected {expected['font_size_pt']}"
                                )
                            
                            # Check bold setting
                            if font.bold != expected["font_bold"]:
                                errors.append(
                                    f"Style assertion failed: {style_def.name} font_bold is "
                                    f"{font.bold}, expected {expected['font_bold']}"
                                )
                        else:
                            errors.append(f"Style assertion failed: {style_def.name} has no font specification")
                        
                        # Check paragraph specifications
                        if style_def.paragraph:
                            para = style_def.paragraph
                            
                            # Check line spacing mode
                            if para.line_spacing_mode != expected["line_spacing_mode"]:
                                errors.append(
                                    f"Style assertion failed: {style_def.name} line_spacing_mode is "
                                    f"'{para.line_spacing_mode}', expected '{expected['line_spacing_mode']}'"
                                )
                            
                            # Check line spacing value
                            if para.line_spacing_value != expected["line_spacing_value"]:
                                errors.append(
                                    f"Style assertion failed: {style_def.name} line_spacing_value is "
                                    f"{para.line_spacing_value}, expected {expected['line_spacing_value']}"
                                )
                        else:
                            errors.append(f"Style assertion failed: {style_def.name} has no paragraph specification")
                        
                        break
                
                if not style_found:
                    errors.append(f"Style assertion failed: Required style '{style_key}' not found in document")
            
            logger.info(f"Style assertions checked: {len(errors)} failures found")
            
        except Exception as e:
            logger.error(f"Error during style assertion checking: {e}")
            errors.append(f"Style assertion check failed: {e}")
        
        return errors
    
    def check_toc_assertions(self, structure: StructureV1) -> List[str]:
        """
        Verify TOC consistency with heading tree.
        
        Args:
            structure: Document structure to validate
            
        Returns:
            List[str]: List of assertion failures
        """
        errors = []
        
        try:
            # Find TOC fields in the document
            toc_fields = []
            for field in structure.fields:
                if "TOC" in field.field_code.upper() if field.field_code else False:
                    toc_fields.append(field)
            
            if not toc_fields:
                # No TOC found - this is acceptable, just log it
                logger.info("No TOC fields found in document")
                return errors
            
            # For each TOC field, verify consistency
            for toc_field in toc_fields:
                try:
                    # Extract TOC entries from result text
                    if not toc_field.result_text:
                        errors.append(f"TOC assertion failed: TOC field at paragraph {toc_field.paragraph_index} has no result text")
                        continue
                    
                    # Parse TOC entries (simplified parsing)
                    toc_entries = self._parse_toc_entries(toc_field.result_text)
                    
                    # Compare with actual headings
                    heading_entries = [(h.text.strip(), h.level) for h in structure.headings]
                    
                    # Check if TOC entries match headings
                    if len(toc_entries) != len(heading_entries):
                        errors.append(
                            f"TOC assertion failed: TOC has {len(toc_entries)} entries, "
                            f"but document has {len(heading_entries)} headings"
                        )
                    
                    # Check individual entries (simplified check)
                    for i, (toc_text, toc_level) in enumerate(toc_entries):
                        if i < len(heading_entries):
                            heading_text, heading_level = heading_entries[i]
                            
                            # Check if heading text matches (allowing for minor differences)
                            if not self._text_matches_approximately(toc_text, heading_text):
                                errors.append(
                                    f"TOC assertion failed: TOC entry '{toc_text}' does not match "
                                    f"heading '{heading_text}' at position {i}"
                                )
                            
                            # Check if levels match
                            if toc_level != heading_level:
                                errors.append(
                                    f"TOC assertion failed: TOC entry '{toc_text}' has level {toc_level}, "
                                    f"but heading has level {heading_level}"
                                )
                
                except Exception as e:
                    errors.append(f"TOC assertion failed: Error processing TOC field: {e}")
            
            logger.info(f"TOC assertions checked: {len(errors)} failures found")
            
        except Exception as e:
            logger.error(f"Error during TOC assertion checking: {e}")
            errors.append(f"TOC assertion check failed: {e}")
        
        return errors
    
    def check_pagination_assertions(self, original_structure: StructureV1, 
                                  modified_structure: StructureV1) -> List[str]:
        """
        Verify pagination assertions with Fields.Update() and Repaginate() verification.
        
        Args:
            original_structure: Original document structure
            modified_structure: Modified document structure
            
        Returns:
            List[str]: List of assertion failures
        """
        errors = []
        
        try:
            # Check if modification time has changed (indicating document was modified)
            if (original_structure.metadata.modified_time and 
                modified_structure.metadata.modified_time):
                
                if (original_structure.metadata.modified_time >= 
                    modified_structure.metadata.modified_time):
                    errors.append(
                        "Pagination assertion failed: Document modification time has not changed, "
                        "indicating fields may not have been updated"
                    )
            
            # Check if page count is reasonable (not zero or negative)
            if modified_structure.metadata.page_count is not None:
                if modified_structure.metadata.page_count <= 0:
                    errors.append(
                        f"Pagination assertion failed: Invalid page count: {modified_structure.metadata.page_count}"
                    )
            
            # Check if fields have been updated by comparing field results
            original_fields = {f.field_code: f.result_text for f in original_structure.fields if f.field_code}
            modified_fields = {f.field_code: f.result_text for f in modified_structure.fields if f.field_code}
            
            # Look for fields that should have been updated
            for field_code in original_fields:
                if field_code in modified_fields:
                    # For page number fields, the result should potentially be different
                    if "PAGE" in field_code.upper() or "NUMPAGES" in field_code.upper():
                        # Page fields should have valid numeric results
                        try:
                            if modified_fields[field_code]:
                                int(modified_fields[field_code].strip())
                        except (ValueError, AttributeError):
                            errors.append(
                                f"Pagination assertion failed: Page field '{field_code}' "
                                f"has invalid result: '{modified_fields[field_code]}'"
                            )
            
            logger.info(f"Pagination assertions checked: {len(errors)} failures found")
            
        except Exception as e:
            logger.error(f"Error during pagination assertion checking: {e}")
            errors.append(f"Pagination assertion check failed: {e}")
        
        return errors
    
    def rollback_document(self, original_docx: str, modified_docx: str):
        """
        Rollback functionality that restores original DOCX on validation failure.
        
        Args:
            original_docx: Path to original DOCX file
            modified_docx: Path to modified DOCX file to be replaced
            
        Raises:
            RollbackError: If rollback fails
        """
        if not os.path.exists(original_docx):
            raise RollbackError(
                f"Cannot rollback: Original DOCX file not found: {original_docx}",
                original_docx=original_docx,
                rollback_reason="validation_failure"
            )
        
        try:
            logger.info(f"Rolling back document from {original_docx} to {modified_docx}")
            
            # Create backup of modified file before rollback
            backup_path = f"{modified_docx}.rollback_backup"
            if os.path.exists(modified_docx):
                shutil.copy2(modified_docx, backup_path)
                logger.debug(f"Created rollback backup: {backup_path}")
            
            # Copy original file over modified file
            shutil.copy2(original_docx, modified_docx)
            
            logger.info(f"Document successfully rolled back to original: {original_docx}")
            
        except Exception as e:
            raise RollbackError(
                f"Failed to rollback document: {e}",
                original_docx=original_docx,
                rollback_reason="validation_failure",
                rollback_exception=e
            )
    
    def _update_fields_and_repaginate(self, docx_path: str):
        """
        Update all fields and repaginate the document.
        
        Args:
            docx_path: Path to DOCX file to update
        """
        try:
            logger.info(f"Updating fields and repaginating: {docx_path}")
            
            # Open document
            doc = self._word_app.Documents.Open(docx_path)
            
            try:
                # Update all fields
                doc.Fields.Update()
                
                # Repaginate document
                doc.Repaginate()
                
                # Save changes
                doc.Save()
                
                logger.info("Fields updated and document repaginated")
                
            finally:
                doc.Close(SaveChanges=True)
                
        except Exception as e:
            logger.error(f"Failed to update fields and repaginate: {e}")
            raise ValidationError(f"Failed to update fields and repaginate: {e}")
    
    def _parse_toc_entries(self, toc_text: str) -> List[tuple]:
        """
        Parse TOC entries from result text.
        
        Args:
            toc_text: TOC result text
            
        Returns:
            List of (text, level) tuples
        """
        entries = []
        
        try:
            lines = toc_text.split('\n')
            for line_num, line in enumerate(lines):
                original_line = line
                line = line.strip()
                if not line:
                    continue
                
                # Count leading whitespace in original line to determine level
                leading_spaces = len(original_line) - len(original_line.lstrip())
                level = 1
                if leading_spaces >= 4:
                    level = 3
                elif leading_spaces >= 2:
                    level = 2
                
                # Remove page numbers (assume they're at the end after tab)
                parts = line.split('\t')
                if len(parts) > 1:
                    text = parts[0].strip()
                else:
                    text = line.strip()
                
                entries.append((text, level))
                
        except Exception as e:
            logger.warning(f"Failed to parse TOC entries: {e}")
        
        return entries
    
    def _text_matches_approximately(self, text1: str, text2: str) -> bool:
        """
        Check if two text strings match approximately (allowing for minor differences).
        
        Args:
            text1: First text string
            text2: Second text string
            
        Returns:
            True if texts match approximately
        """
        # Normalize texts
        norm1 = text1.strip().lower()
        norm2 = text2.strip().lower()
        
        # Exact match
        if norm1 == norm2:
            return True
        
        # Check if one contains the other (for cases where TOC might have truncated text)
        if norm1 in norm2 or norm2 in norm1:
            return True
        
        # Check similarity (simple approach - could be enhanced with fuzzy matching)
        if len(norm1) > 0 and len(norm2) > 0:
            # Calculate simple similarity ratio
            common_chars = sum(1 for c in norm1 if c in norm2)
            similarity = common_chars / max(len(norm1), len(norm2))
            return similarity > 0.8
        
        return False