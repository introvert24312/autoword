"""
Advanced validation and quality assurance module for AutoWord vNext.

This module implements comprehensive document integrity verification, style consistency
checking, cross-reference validation, accessibility compliance checking, and formatting
quality metrics beyond the basic assertions in DocumentValidator.
"""

import os
import logging
from typing import List, Dict, Any, Optional, Set, Tuple
from pathlib import Path
from collections import defaultdict, Counter
import re

import pythoncom
import win32com.client as win32

from ..models import (
    StructureV1, ValidationResult, StyleDefinition, FontSpec, ParagraphSpec,
    HeadingReference, FieldReference, TableSkeleton, CrossReference,
    LineSpacingMode, StyleType
)
from ..exceptions import ValidationError
from ..extractor.document_extractor import DocumentExtractor


logger = logging.getLogger(__name__)


class QualityMetrics:
    """Quality metrics for document formatting assessment."""
    
    def __init__(self):
        self.style_consistency_score: float = 0.0
        self.cross_reference_integrity_score: float = 0.0
        self.accessibility_score: float = 0.0
        self.formatting_quality_score: float = 0.0
        self.overall_score: float = 0.0
        
        # Detailed metrics
        self.inconsistent_styles: List[str] = []
        self.broken_cross_references: List[str] = []
        self.accessibility_issues: List[str] = []
        self.formatting_issues: List[str] = []
        
        # Statistics
        self.total_styles_checked: int = 0
        self.total_cross_references_checked: int = 0
        self.total_accessibility_checks: int = 0
        self.total_formatting_checks: int = 0


class AdvancedValidator:
    """Advanced document validation and quality assurance system."""
    
    def __init__(self, visible: bool = False):
        """
        Initialize advanced validator.
        
        Args:
            visible: Whether to show Word application window during validation
        """
        self.visible = visible
        self._word_app = None
        self._com_initialized = False
        
        # Configuration for validation rules
        self.min_heading_hierarchy_score = 0.8
        self.max_style_variations_per_type = 3
        self.min_accessibility_score = 0.7
        self.min_cross_reference_integrity = 0.9
    
    def __enter__(self):
        """Enter context manager for COM resource management."""
        try:
            # Initialize COM
            pythoncom.CoInitialize()
            self._com_initialized = True
            logger.debug("COM initialized for advanced validation")
            
            # Create Word application instance
            self._word_app = win32.gencache.EnsureDispatch('Word.Application')
            self._word_app.Visible = self.visible
            self._word_app.DisplayAlerts = 0  # Disable alerts
            self._word_app.ScreenUpdating = False  # Improve performance
            
            logger.info(f"Word application started for advanced validation (visible={self.visible})")
            return self
            
        except Exception as e:
            self._cleanup()
            raise ValidationError(f"Failed to initialize Word COM for advanced validation: {e}")
    
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
                logger.info("Word application closed after advanced validation")
        except Exception as e:
            logger.warning(f"Error during Word cleanup: {e}")
        
        finally:
            if self._com_initialized:
                try:
                    pythoncom.CoUninitialize()
                    self._com_initialized = False
                    logger.debug("COM uninitialized after advanced validation")
                except Exception as e:
                    logger.warning(f"Error during COM cleanup: {e}")
    
    def validate_document_integrity(self, structure: StructureV1, docx_path: str) -> ValidationResult:
        """
        Perform comprehensive document integrity verification beyond basic assertions.
        
        Args:
            structure: Document structure to validate
            docx_path: Path to DOCX file for deep validation
            
        Returns:
            ValidationResult: Comprehensive validation result
        """
        errors = []
        warnings = []
        
        try:
            logger.info(f"Performing advanced document integrity validation: {docx_path}")
            
            # 1. Validate document structure integrity
            structure_errors = self._validate_structure_integrity(structure)
            errors.extend(structure_errors)
            
            # 2. Validate heading hierarchy
            hierarchy_errors = self._validate_heading_hierarchy(structure)
            errors.extend(hierarchy_errors)
            
            # 3. Validate paragraph consistency
            paragraph_errors = self._validate_paragraph_consistency(structure)
            errors.extend(paragraph_errors)
            
            # 4. Validate table integrity
            table_errors = self._validate_table_integrity(structure)
            errors.extend(table_errors)
            
            # 5. Validate field integrity
            field_errors = self._validate_field_integrity(structure)
            errors.extend(field_errors)
            
            # 6. Deep validation with Word COM if available
            if self._word_app and os.path.exists(docx_path):
                com_errors, com_warnings = self._validate_with_word_com(docx_path)
                errors.extend(com_errors)
                warnings.extend(com_warnings)
            
            is_valid = len(errors) == 0
            logger.info(f"Document integrity validation completed: {len(errors)} errors, {len(warnings)} warnings")
            
            return ValidationResult(
                is_valid=is_valid,
                errors=errors,
                warnings=warnings
            )
            
        except Exception as e:
            logger.error(f"Error during document integrity validation: {e}")
            return ValidationResult(
                is_valid=False,
                errors=[f"Document integrity validation failed: {e}"],
                warnings=warnings
            )
    
    def check_style_consistency(self, structure: StructureV1) -> ValidationResult:
        """
        Perform comprehensive style consistency checking across entire document.
        
        Args:
            structure: Document structure to validate
            
        Returns:
            ValidationResult: Style consistency validation result
        """
        errors = []
        warnings = []
        
        try:
            logger.info("Performing style consistency validation")
            
            # 1. Check for style definition consistency
            style_def_errors = self._check_style_definitions(structure)
            errors.extend(style_def_errors)
            
            # 2. Check for style usage consistency
            style_usage_errors = self._check_style_usage_consistency(structure)
            errors.extend(style_usage_errors)
            
            # 3. Check for font consistency within styles
            font_consistency_errors = self._check_font_consistency(structure)
            errors.extend(font_consistency_errors)
            
            # 4. Check for paragraph formatting consistency
            para_consistency_errors = self._check_paragraph_formatting_consistency(structure)
            errors.extend(para_consistency_errors)
            
            # 5. Check for heading style consistency
            heading_style_errors = self._check_heading_style_consistency(structure)
            errors.extend(heading_style_errors)
            
            # 6. Generate style consistency warnings
            style_warnings = self._generate_style_warnings(structure)
            warnings.extend(style_warnings)
            
            is_valid = len(errors) == 0
            logger.info(f"Style consistency validation completed: {len(errors)} errors, {len(warnings)} warnings")
            
            return ValidationResult(
                is_valid=is_valid,
                errors=errors,
                warnings=warnings
            )
            
        except Exception as e:
            logger.error(f"Error during style consistency validation: {e}")
            return ValidationResult(
                is_valid=False,
                errors=[f"Style consistency validation failed: {e}"],
                warnings=warnings
            )
    
    def validate_cross_references(self, structure: StructureV1, docx_path: str) -> ValidationResult:
        """
        Validate cross-references and provide repair recommendations.
        
        Args:
            structure: Document structure to validate
            docx_path: Path to DOCX file for cross-reference validation
            
        Returns:
            ValidationResult: Cross-reference validation result
        """
        errors = []
        warnings = []
        
        try:
            logger.info("Performing cross-reference validation")
            
            # 1. Validate heading references
            heading_ref_errors = self._validate_heading_references(structure)
            errors.extend(heading_ref_errors)
            
            # 2. Validate table references
            table_ref_errors = self._validate_table_references(structure)
            errors.extend(table_ref_errors)
            
            # 3. Validate figure references (if any)
            figure_ref_errors = self._validate_figure_references(structure)
            errors.extend(figure_ref_errors)
            
            # 4. Validate TOC references
            toc_ref_errors = self._validate_toc_references(structure)
            errors.extend(toc_ref_errors)
            
            # 5. Check for broken field references
            field_ref_errors = self._validate_field_references(structure)
            errors.extend(field_ref_errors)
            
            # 6. Generate repair recommendations
            repair_warnings = self._generate_cross_reference_repair_recommendations(structure)
            warnings.extend(repair_warnings)
            
            is_valid = len(errors) == 0
            logger.info(f"Cross-reference validation completed: {len(errors)} errors, {len(warnings)} warnings")
            
            return ValidationResult(
                is_valid=is_valid,
                errors=errors,
                warnings=warnings
            )
            
        except Exception as e:
            logger.error(f"Error during cross-reference validation: {e}")
            return ValidationResult(
                is_valid=False,
                errors=[f"Cross-reference validation failed: {e}"],
                warnings=warnings
            )
    
    def check_accessibility_compliance(self, structure: StructureV1) -> ValidationResult:
        """
        Check document accessibility compliance.
        
        Args:
            structure: Document structure to validate
            
        Returns:
            ValidationResult: Accessibility compliance result
        """
        errors = []
        warnings = []
        
        try:
            logger.info("Performing accessibility compliance validation")
            
            # 1. Check heading structure for accessibility
            heading_accessibility_errors = self._check_heading_accessibility(structure)
            errors.extend(heading_accessibility_errors)
            
            # 2. Check table accessibility
            table_accessibility_errors = self._check_table_accessibility(structure)
            errors.extend(table_accessibility_errors)
            
            # 3. Check color contrast and font readability
            readability_errors = self._check_readability_compliance(structure)
            errors.extend(readability_errors)
            
            # 4. Check document structure for screen readers
            structure_accessibility_errors = self._check_structure_accessibility(structure)
            errors.extend(structure_accessibility_errors)
            
            # 5. Check for alternative text and descriptions
            alt_text_warnings = self._check_alternative_text_compliance(structure)
            warnings.extend(alt_text_warnings)
            
            # 6. Generate accessibility improvement recommendations
            accessibility_warnings = self._generate_accessibility_recommendations(structure)
            warnings.extend(accessibility_warnings)
            
            is_valid = len(errors) == 0
            logger.info(f"Accessibility compliance validation completed: {len(errors)} errors, {len(warnings)} warnings")
            
            return ValidationResult(
                is_valid=is_valid,
                errors=errors,
                warnings=warnings
            )
            
        except Exception as e:
            logger.error(f"Error during accessibility compliance validation: {e}")
            return ValidationResult(
                is_valid=False,
                errors=[f"Accessibility compliance validation failed: {e}"],
                warnings=warnings
            )
    
    def generate_quality_metrics(self, structure: StructureV1, docx_path: str) -> QualityMetrics:
        """
        Generate comprehensive formatting quality metrics and reporting.
        
        Args:
            structure: Document structure to analyze
            docx_path: Path to DOCX file for analysis
            
        Returns:
            QualityMetrics: Comprehensive quality assessment
        """
        metrics = QualityMetrics()
        
        try:
            logger.info("Generating document quality metrics")
            
            # 1. Calculate style consistency score
            metrics.style_consistency_score = self._calculate_style_consistency_score(structure)
            metrics.total_styles_checked = len(structure.styles)
            
            # 2. Calculate cross-reference integrity score
            metrics.cross_reference_integrity_score = self._calculate_cross_reference_score(structure)
            metrics.total_cross_references_checked = len(structure.fields)
            
            # 3. Calculate accessibility score
            metrics.accessibility_score = self._calculate_accessibility_score(structure)
            metrics.total_accessibility_checks = self._count_accessibility_checks(structure)
            
            # 4. Calculate formatting quality score
            metrics.formatting_quality_score = self._calculate_formatting_quality_score(structure)
            metrics.total_formatting_checks = self._count_formatting_checks(structure)
            
            # 5. Calculate overall score
            metrics.overall_score = (
                metrics.style_consistency_score * 0.3 +
                metrics.cross_reference_integrity_score * 0.2 +
                metrics.accessibility_score * 0.25 +
                metrics.formatting_quality_score * 0.25
            )
            
            # 6. Collect detailed issues
            metrics.inconsistent_styles = self._collect_inconsistent_styles(structure)
            metrics.broken_cross_references = self._collect_broken_cross_references(structure)
            metrics.accessibility_issues = self._collect_accessibility_issues(structure)
            metrics.formatting_issues = self._collect_formatting_issues(structure)
            
            logger.info(f"Quality metrics generated: overall score {metrics.overall_score:.2f}")
            
        except Exception as e:
            logger.error(f"Error generating quality metrics: {e}")
            metrics.overall_score = 0.0
            metrics.formatting_issues.append(f"Quality metrics generation failed: {e}")
        
        return metrics 
   
    # Private methods for document integrity validation
    
    def _validate_structure_integrity(self, structure: StructureV1) -> List[str]:
        """Validate basic document structure integrity."""
        errors = []
        
        # Check for required metadata
        if not structure.metadata:
            errors.append("Document integrity error: Missing metadata")
        
        # Check paragraph indexing consistency
        paragraph_indexes = [p.index for p in structure.paragraphs]
        if paragraph_indexes != sorted(paragraph_indexes):
            errors.append("Document integrity error: Paragraph indexes are not sequential")
        
        # Check for duplicate paragraph indexes
        if len(paragraph_indexes) != len(set(paragraph_indexes)):
            errors.append("Document integrity error: Duplicate paragraph indexes found")
        
        # Validate heading references point to valid paragraphs
        max_paragraph_index = max(paragraph_indexes) if paragraph_indexes else -1
        for heading in structure.headings:
            if heading.paragraph_index > max_paragraph_index:
                errors.append(
                    f"Document integrity error: Heading '{heading.text}' references "
                    f"invalid paragraph index {heading.paragraph_index}"
                )
        
        return errors
    
    def _validate_heading_hierarchy(self, structure: StructureV1) -> List[str]:
        """Validate heading hierarchy structure."""
        errors = []
        
        if not structure.headings:
            return errors
        
        # Check for proper heading level progression
        prev_level = 0
        for i, heading in enumerate(structure.headings):
            if heading.level > prev_level + 1:
                errors.append(
                    f"Heading hierarchy error: Heading '{heading.text}' at level {heading.level} "
                    f"skips levels (previous was level {prev_level})"
                )
            prev_level = heading.level
        
        # Check for empty headings
        for heading in structure.headings:
            if not heading.text.strip():
                errors.append(
                    f"Heading hierarchy error: Empty heading at paragraph {heading.paragraph_index}"
                )
        
        # Check for duplicate headings at same level
        level_headings = defaultdict(list)
        for heading in structure.headings:
            level_headings[heading.level].append(heading.text.strip().lower())
        
        for level, texts in level_headings.items():
            duplicates = [text for text, count in Counter(texts).items() if count > 1]
            if duplicates:
                errors.append(
                    f"Heading hierarchy error: Duplicate headings at level {level}: {duplicates}"
                )
        
        return errors
    
    def _validate_paragraph_consistency(self, structure: StructureV1) -> List[str]:
        """Validate paragraph consistency."""
        errors = []
        
        # Check for paragraphs with missing or invalid styles
        for paragraph in structure.paragraphs:
            if paragraph.style_name:
                # Check if style exists in document styles
                style_exists = any(s.name == paragraph.style_name for s in structure.styles)
                if not style_exists:
                    errors.append(
                        f"Paragraph consistency error: Paragraph {paragraph.index} "
                        f"references undefined style '{paragraph.style_name}'"
                    )
        
        # Check for inconsistent heading markers
        for paragraph in structure.paragraphs:
            if paragraph.is_heading and not paragraph.heading_level:
                errors.append(
                    f"Paragraph consistency error: Paragraph {paragraph.index} "
                    f"marked as heading but has no heading level"
                )
            elif not paragraph.is_heading and paragraph.heading_level:
                errors.append(
                    f"Paragraph consistency error: Paragraph {paragraph.index} "
                    f"has heading level but not marked as heading"
                )
        
        return errors
    
    def _validate_table_integrity(self, structure: StructureV1) -> List[str]:
        """Validate table structure integrity."""
        errors = []
        
        for table in structure.tables:
            # Check for valid table dimensions
            if table.rows <= 0 or table.columns <= 0:
                errors.append(
                    f"Table integrity error: Table at paragraph {table.paragraph_index} "
                    f"has invalid dimensions: {table.rows}x{table.columns}"
                )
            
            # Check cell references if provided
            if table.cell_references:
                max_paragraph_index = max([p.index for p in structure.paragraphs]) if structure.paragraphs else -1
                for cell_ref in table.cell_references:
                    if cell_ref > max_paragraph_index:
                        errors.append(
                            f"Table integrity error: Table at paragraph {table.paragraph_index} "
                            f"references invalid cell paragraph {cell_ref}"
                        )
            
            # Validate cell paragraph mapping
            if table.cell_paragraph_map:
                expected_cells = table.rows * table.columns
                actual_cells = len(table.cell_paragraph_map)
                if actual_cells != expected_cells:
                    errors.append(
                        f"Table integrity error: Table at paragraph {table.paragraph_index} "
                        f"has {actual_cells} cell mappings but should have {expected_cells}"
                    )
        
        return errors
    
    def _validate_field_integrity(self, structure: StructureV1) -> List[str]:
        """Validate field integrity."""
        errors = []
        
        for field in structure.fields:
            # Check for valid field codes
            if not field.field_code or not field.field_code.strip():
                errors.append(
                    f"Field integrity error: Field at paragraph {field.paragraph_index} "
                    f"has empty field code"
                )
            
            # Check for common field code patterns
            field_code = field.field_code.upper() if field.field_code else ""
            if "TOC" in field_code and not field.result_text:
                errors.append(
                    f"Field integrity error: TOC field at paragraph {field.paragraph_index} "
                    f"has no result text"
                )
            
            # Check for page number fields
            if any(keyword in field_code for keyword in ["PAGE", "NUMPAGES"]):
                if field.result_text and not field.result_text.strip().isdigit():
                    errors.append(
                        f"Field integrity error: Page field at paragraph {field.paragraph_index} "
                        f"has non-numeric result: '{field.result_text}'"
                    )
        
        return errors
    
    def _validate_with_word_com(self, docx_path: str) -> Tuple[List[str], List[str]]:
        """Perform deep validation using Word COM."""
        errors = []
        warnings = []
        
        try:
            doc = self._word_app.Documents.Open(docx_path)
            
            try:
                # Check for spelling and grammar errors
                if doc.SpellingErrors.Count > 0:
                    warnings.append(f"Document has {doc.SpellingErrors.Count} spelling errors")
                
                if doc.GrammarErrors.Count > 0:
                    warnings.append(f"Document has {doc.GrammarErrors.Count} grammar errors")
                
                # Check for broken links
                for hyperlink in doc.Hyperlinks:
                    if not hyperlink.Address and not hyperlink.SubAddress:
                        warnings.append(f"Broken hyperlink found: '{hyperlink.TextToDisplay}'")
                
                # Check for unresolved cross-references
                for field in doc.Fields:
                    if "REF" in field.Code.Text.upper():
                        if "Error!" in field.Result.Text:
                            errors.append(f"Broken cross-reference: {field.Code.Text}")
                
                # Check document protection
                if doc.ProtectionType != 0:  # wdNoProtection
                    warnings.append("Document has protection enabled")
                
                # Check for track changes
                if doc.TrackRevisions:
                    warnings.append("Document has track changes enabled")
                
                # Check for comments
                if doc.Comments.Count > 0:
                    warnings.append(f"Document has {doc.Comments.Count} comments")
                
            finally:
                doc.Close(SaveChanges=False)
                
        except Exception as e:
            errors.append(f"Word COM validation error: {e}")
        
        return errors, warnings  
  
    # Private methods for style consistency validation
    
    def _check_style_definitions(self, structure: StructureV1) -> List[str]:
        """Check for consistent style definitions."""
        errors = []
        
        # Check for duplicate style names
        style_names = [s.name for s in structure.styles]
        duplicates = [name for name, count in Counter(style_names).items() if count > 1]
        if duplicates:
            errors.append(f"Style definition error: Duplicate style names found: {duplicates}")
        
        # Check for incomplete style definitions
        for style in structure.styles:
            if style.type == StyleType.PARAGRAPH and not style.paragraph:
                errors.append(f"Style definition error: Paragraph style '{style.name}' missing paragraph spec")
            
            if not style.font:
                errors.append(f"Style definition error: Style '{style.name}' missing font specification")
        
        return errors
    
    def _check_style_usage_consistency(self, structure: StructureV1) -> List[str]:
        """Check for consistent style usage across document."""
        errors = []
        
        # Collect style usage statistics
        style_usage = Counter()
        for paragraph in structure.paragraphs:
            if paragraph.style_name:
                style_usage[paragraph.style_name] += 1
        
        # Check for unused styles
        defined_styles = {s.name for s in structure.styles}
        used_styles = set(style_usage.keys())
        unused_styles = defined_styles - used_styles
        
        if unused_styles:
            errors.append(f"Style usage error: Unused styles found: {list(unused_styles)}")
        
        # Check for undefined styles in use
        undefined_styles = used_styles - defined_styles
        if undefined_styles:
            errors.append(f"Style usage error: Undefined styles in use: {list(undefined_styles)}")
        
        return errors
    
    def _check_font_consistency(self, structure: StructureV1) -> List[str]:
        """Check for font consistency within styles."""
        errors = []
        
        # Group styles by type and check font consistency
        style_groups = defaultdict(list)
        for style in structure.styles:
            if style.font:
                style_groups[style.type].append(style)
        
        # Check for inconsistent font sizes within similar styles
        for style_type, styles in style_groups.items():
            if style_type == StyleType.PARAGRAPH:
                font_sizes = [s.font.size_pt for s in styles if s.font and s.font.size_pt]
                if len(set(font_sizes)) > self.max_style_variations_per_type:
                    errors.append(
                        f"Font consistency error: Too many font size variations in {style_type} styles: {set(font_sizes)}"
                    )
        
        return errors
    
    def _check_paragraph_formatting_consistency(self, structure: StructureV1) -> List[str]:
        """Check for paragraph formatting consistency."""
        errors = []
        
        # Check line spacing consistency
        line_spacings = defaultdict(list)
        for style in structure.styles:
            if style.paragraph and style.paragraph.line_spacing_value:
                line_spacings[style.type].append(style.paragraph.line_spacing_value)
        
        for style_type, spacings in line_spacings.items():
            if len(set(spacings)) > self.max_style_variations_per_type:
                errors.append(
                    f"Paragraph formatting error: Too many line spacing variations in {style_type} styles: {set(spacings)}"
                )
        
        return errors
    
    def _check_heading_style_consistency(self, structure: StructureV1) -> List[str]:
        """Check for heading style consistency."""
        errors = []
        
        # Check that all headings have consistent styling
        heading_styles = {}
        for heading in structure.headings:
            if heading.style_name:
                if heading.level not in heading_styles:
                    heading_styles[heading.level] = heading.style_name
                elif heading_styles[heading.level] != heading.style_name:
                    errors.append(
                        f"Heading style consistency error: Level {heading.level} headings use multiple styles: "
                        f"'{heading_styles[heading.level]}' and '{heading.style_name}'"
                    )
        
        return errors
    
    def _generate_style_warnings(self, structure: StructureV1) -> List[str]:
        """Generate style-related warnings."""
        warnings = []
        
        # Check for potential style improvements
        for style in structure.styles:
            if style.font and style.font.size_pt:
                if style.font.size_pt < 10:
                    warnings.append(f"Style warning: Style '{style.name}' has small font size ({style.font.size_pt}pt)")
                elif style.font.size_pt > 18:
                    warnings.append(f"Style warning: Style '{style.name}' has large font size ({style.font.size_pt}pt)")
        
        return warnings
    
    # Private methods for cross-reference validation
    
    def _validate_heading_references(self, structure: StructureV1) -> List[str]:
        """Validate heading cross-references."""
        errors = []
        
        # Create heading lookup by text
        heading_texts = {h.text.strip().lower(): h for h in structure.headings}
        
        # Check field references to headings
        for field in structure.fields:
            if field.field_code and "REF" in field.field_code.upper():
                # Extract referenced heading text (simplified parsing)
                ref_match = re.search(r'REF\s+([^\\]+)', field.field_code, re.IGNORECASE)
                if ref_match:
                    ref_text = ref_match.group(1).strip().strip('"').lower()
                    if ref_text not in heading_texts:
                        errors.append(
                            f"Cross-reference error: Heading reference '{ref_text}' not found "
                            f"(field at paragraph {field.paragraph_index})"
                        )
        
        return errors
    
    def _validate_table_references(self, structure: StructureV1) -> List[str]:
        """Validate table cross-references."""
        errors = []
        
        # Create table lookup
        table_positions = {t.paragraph_index: t for t in structure.tables}
        
        # Check for table references in fields
        for field in structure.fields:
            if field.field_code and any(keyword in field.field_code.upper() for keyword in ["TABLE", "TAB"]):
                # This is a simplified check - in practice, would need more sophisticated parsing
                if "Error!" in (field.result_text or ""):
                    errors.append(
                        f"Cross-reference error: Broken table reference at paragraph {field.paragraph_index}"
                    )
        
        return errors
    
    def _validate_figure_references(self, structure: StructureV1) -> List[str]:
        """Validate figure cross-references."""
        errors = []
        
        # Check for figure references in fields
        for field in structure.fields:
            if field.field_code and any(keyword in field.field_code.upper() for keyword in ["FIGURE", "FIG"]):
                if "Error!" in (field.result_text or ""):
                    errors.append(
                        f"Cross-reference error: Broken figure reference at paragraph {field.paragraph_index}"
                    )
        
        return errors
    
    def _validate_toc_references(self, structure: StructureV1) -> List[str]:
        """Validate TOC cross-references."""
        errors = []
        
        # Find TOC fields
        toc_fields = [f for f in structure.fields if f.field_code and "TOC" in f.field_code.upper()]
        
        for toc_field in toc_fields:
            if not toc_field.result_text or not toc_field.result_text.strip():
                errors.append(
                    f"Cross-reference error: Empty TOC at paragraph {toc_field.paragraph_index}"
                )
            elif "Error!" in toc_field.result_text:
                errors.append(
                    f"Cross-reference error: Broken TOC at paragraph {toc_field.paragraph_index}"
                )
        
        return errors
    
    def _validate_field_references(self, structure: StructureV1) -> List[str]:
        """Validate field cross-references."""
        errors = []
        
        for field in structure.fields:
            # Check for error indicators in field results
            if field.result_text and "Error!" in field.result_text:
                errors.append(
                    f"Cross-reference error: Field error at paragraph {field.paragraph_index}: {field.result_text}"
                )
            
            # Check for missing field results where expected
            if field.field_code:
                field_code_upper = field.field_code.upper()
                if any(keyword in field_code_upper for keyword in ["REF", "PAGEREF", "HYPERLINK"]):
                    if not field.result_text or not field.result_text.strip():
                        errors.append(
                            f"Cross-reference error: Missing field result at paragraph {field.paragraph_index}"
                        )
        
        return errors
    
    def _generate_cross_reference_repair_recommendations(self, structure: StructureV1) -> List[str]:
        """Generate cross-reference repair recommendations."""
        warnings = []
        
        # Suggest updating all fields
        if structure.fields:
            warnings.append("Recommendation: Update all fields (Ctrl+A, F9) to refresh cross-references")
        
        # Check for potential orphaned references
        ref_fields = [f for f in structure.fields if f.field_code and "REF" in f.field_code.upper()]
        if ref_fields:
            warnings.append(f"Found {len(ref_fields)} cross-references - verify all targets exist")
        
        return warnings
    
    # Private methods for accessibility validation
    
    def _check_heading_accessibility(self, structure: StructureV1) -> List[str]:
        """Check heading structure for accessibility compliance."""
        errors = []
        
        if not structure.headings:
            errors.append("Accessibility error: Document has no headings for navigation")
            return errors
        
        # Check for proper heading hierarchy (no skipped levels)
        prev_level = 0
        for heading in structure.headings:
            if heading.level > prev_level + 1:
                errors.append(
                    f"Accessibility error: Heading level skipped - '{heading.text}' "
                    f"is level {heading.level} but previous was level {prev_level}"
                )
            prev_level = heading.level
        
        # Check for meaningful heading text
        for heading in structure.headings:
            if len(heading.text.strip()) < 3:
                errors.append(
                    f"Accessibility error: Heading text too short for screen readers: '{heading.text}'"
                )
        
        return errors
    
    def _check_table_accessibility(self, structure: StructureV1) -> List[str]:
        """Check table accessibility compliance."""
        errors = []
        
        for table in structure.tables:
            # Check for header rows
            if not table.has_header:
                errors.append(
                    f"Accessibility error: Table at paragraph {table.paragraph_index} "
                    f"should have header row for screen readers"
                )
            
            # Check table size for accessibility
            if table.rows > 20 or table.columns > 10:
                errors.append(
                    f"Accessibility warning: Large table at paragraph {table.paragraph_index} "
                    f"({table.rows}x{table.columns}) may be difficult to navigate"
                )
        
        return errors
    
    def _check_readability_compliance(self, structure: StructureV1) -> List[str]:
        """Check color contrast and font readability."""
        errors = []
        
        for style in structure.styles:
            if style.font:
                # Check font size for readability
                if style.font.size_pt and style.font.size_pt < 9:
                    errors.append(
                        f"Accessibility error: Style '{style.name}' font size ({style.font.size_pt}pt) "
                        f"is too small for accessibility"
                    )
                
                # Check for color-only information (simplified check)
                if style.font.color_hex and style.font.color_hex.lower() in ["#ff0000", "#00ff00"]:
                    errors.append(
                        f"Accessibility warning: Style '{style.name}' uses color that may not be "
                        f"distinguishable for colorblind users"
                    )
        
        return errors
    
    def _check_structure_accessibility(self, structure: StructureV1) -> List[str]:
        """Check document structure for screen reader accessibility."""
        errors = []
        
        # Check for logical reading order
        if structure.paragraphs:
            # Ensure paragraphs are in sequential order
            indexes = [p.index for p in structure.paragraphs]
            if indexes != sorted(indexes):
                errors.append("Accessibility error: Paragraph order is not sequential for screen readers")
        
        # Check for proper use of lists (would need more detailed structure info)
        # This is a placeholder for more sophisticated list structure checking
        
        return errors
    
    def _check_alternative_text_compliance(self, structure: StructureV1) -> List[str]:
        """Check for alternative text and descriptions."""
        warnings = []
        
        # This would require access to embedded objects and images
        # For now, provide general guidance
        if structure.tables:
            warnings.append("Accessibility recommendation: Ensure all images have alternative text")
        
        return warnings
    
    def _generate_accessibility_recommendations(self, structure: StructureV1) -> List[str]:
        """Generate accessibility improvement recommendations."""
        warnings = []
        
        # General accessibility recommendations
        warnings.append("Accessibility recommendation: Use built-in heading styles for proper structure")
        warnings.append("Accessibility recommendation: Ensure sufficient color contrast (4.5:1 minimum)")
        warnings.append("Accessibility recommendation: Use descriptive link text instead of 'click here'")
        
        if structure.tables:
            warnings.append("Accessibility recommendation: Use table headers and captions")
        
        return warnings
    
    # Private methods for quality metrics calculation
    
    def _calculate_style_consistency_score(self, structure: StructureV1) -> float:
        """Calculate style consistency score (0.0 to 1.0)."""
        if not structure.styles:
            return 0.0
        
        score = 1.0
        
        # Penalize for style definition issues
        style_names = [s.name for s in structure.styles]
        duplicates = len(style_names) - len(set(style_names))
        score -= (duplicates * 0.1)
        
        # Penalize for incomplete style definitions
        incomplete_styles = sum(1 for s in structure.styles if not s.font or not s.paragraph)
        score -= (incomplete_styles / len(structure.styles)) * 0.3
        
        # Penalize for excessive style variations
        font_sizes = [s.font.size_pt for s in structure.styles if s.font and s.font.size_pt]
        if len(set(font_sizes)) > self.max_style_variations_per_type:
            score -= 0.2
        
        return max(0.0, min(1.0, score))
    
    def _calculate_cross_reference_score(self, structure: StructureV1) -> float:
        """Calculate cross-reference integrity score (0.0 to 1.0)."""
        if not structure.fields:
            return 1.0  # No cross-references to break
        
        total_fields = len(structure.fields)
        broken_fields = 0
        
        for field in structure.fields:
            if field.result_text and "Error!" in field.result_text:
                broken_fields += 1
            elif not field.result_text and field.field_code:
                # Check if this type of field should have result text
                field_code_upper = field.field_code.upper()
                if any(keyword in field_code_upper for keyword in ["REF", "PAGEREF", "TOC"]):
                    broken_fields += 1
        
        return 1.0 - (broken_fields / total_fields)
    
    def _calculate_accessibility_score(self, structure: StructureV1) -> float:
        """Calculate accessibility compliance score (0.0 to 1.0)."""
        score = 1.0
        
        # Check heading structure
        if not structure.headings:
            score -= 0.3
        else:
            # Check for proper heading hierarchy
            prev_level = 0
            hierarchy_violations = 0
            for heading in structure.headings:
                if heading.level > prev_level + 1:
                    hierarchy_violations += 1
                prev_level = heading.level
            
            if hierarchy_violations > 0:
                score -= (hierarchy_violations / len(structure.headings)) * 0.2
        
        # Check font sizes
        small_fonts = 0
        total_styles = len(structure.styles)
        if total_styles > 0:
            for style in structure.styles:
                if style.font and style.font.size_pt and style.font.size_pt < 9:
                    small_fonts += 1
            score -= (small_fonts / total_styles) * 0.2
        
        # Check table accessibility
        if structure.tables:
            tables_without_headers = sum(1 for t in structure.tables if not t.has_header)
            score -= (tables_without_headers / len(structure.tables)) * 0.1
        
        return max(0.0, min(1.0, score))
    
    def _calculate_formatting_quality_score(self, structure: StructureV1) -> float:
        """Calculate formatting quality score (0.0 to 1.0)."""
        score = 1.0
        
        # Check for consistent paragraph formatting
        if structure.styles:
            line_spacings = [s.paragraph.line_spacing_value for s in structure.styles 
                           if s.paragraph and s.paragraph.line_spacing_value]
            if len(set(line_spacings)) > self.max_style_variations_per_type:
                score -= 0.2
        
        # Check for proper document structure
        if not structure.headings:
            score -= 0.3
        
        # Check for reasonable paragraph count
        if len(structure.paragraphs) < 5:
            score -= 0.1
        
        return max(0.0, min(1.0, score))
    
    def _count_accessibility_checks(self, structure: StructureV1) -> int:
        """Count total accessibility checks performed."""
        checks = 0
        checks += len(structure.headings)  # Heading checks
        checks += len(structure.tables)   # Table checks
        checks += len(structure.styles)   # Style checks
        checks += 1  # Structure check
        return checks
    
    def _count_formatting_checks(self, structure: StructureV1) -> int:
        """Count total formatting checks performed."""
        checks = 0
        checks += len(structure.styles)     # Style checks
        checks += len(structure.paragraphs) # Paragraph checks
        checks += len(structure.headings)   # Heading checks
        checks += 1  # Overall structure check
        return checks
    
    def _collect_inconsistent_styles(self, structure: StructureV1) -> List[str]:
        """Collect list of inconsistent styles."""
        issues = []
        
        # Find duplicate style names
        style_names = [s.name for s in structure.styles]
        duplicates = [name for name, count in Counter(style_names).items() if count > 1]
        issues.extend([f"Duplicate style: {name}" for name in duplicates])
        
        # Find incomplete styles
        for style in structure.styles:
            if not style.font:
                issues.append(f"Missing font spec: {style.name}")
            if style.type == StyleType.PARAGRAPH and not style.paragraph:
                issues.append(f"Missing paragraph spec: {style.name}")
        
        return issues
    
    def _collect_broken_cross_references(self, structure: StructureV1) -> List[str]:
        """Collect list of broken cross-references."""
        issues = []
        
        for field in structure.fields:
            if field.result_text and "Error!" in field.result_text:
                issues.append(f"Broken field at paragraph {field.paragraph_index}: {field.field_code}")
        
        return issues
    
    def _collect_accessibility_issues(self, structure: StructureV1) -> List[str]:
        """Collect list of accessibility issues."""
        issues = []
        
        # Check heading hierarchy
        if not structure.headings:
            issues.append("No headings for navigation")
        
        # Check small fonts
        for style in structure.styles:
            if style.font and style.font.size_pt and style.font.size_pt < 9:
                issues.append(f"Small font in style: {style.name} ({style.font.size_pt}pt)")
        
        # Check tables without headers
        for table in structure.tables:
            if not table.has_header:
                issues.append(f"Table without header at paragraph {table.paragraph_index}")
        
        return issues
    
    def _collect_formatting_issues(self, structure: StructureV1) -> List[str]:
        """Collect list of formatting issues."""
        issues = []
        
        # Check for excessive style variations
        font_sizes = [s.font.size_pt for s in structure.styles if s.font and s.font.size_pt]
        if len(set(font_sizes)) > self.max_style_variations_per_type:
            issues.append(f"Too many font size variations: {sorted(set(font_sizes))}")
        
        # Check line spacing variations
        line_spacings = [s.paragraph.line_spacing_value for s in structure.styles 
                        if s.paragraph and s.paragraph.line_spacing_value]
        if len(set(line_spacings)) > self.max_style_variations_per_type:
            issues.append(f"Too many line spacing variations: {sorted(set(line_spacings))}")
        
        return issues