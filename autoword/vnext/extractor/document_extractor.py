"""
Document extractor for converting DOCX to structured JSON representations.

This module implements the Extractor component of the AutoWord vNext pipeline,
responsible for converting DOCX files to structured JSON with zero information loss.
"""

import os
import logging
import zipfile
import xml.etree.ElementTree as ET
from typing import Tuple, List, Dict, Optional, Any
from datetime import datetime
from pathlib import Path

import pythoncom
import win32com.client as win32
from win32com.client import constants as win32_constants

from ..models import (
    StructureV1, InventoryFullV1, DocumentMetadata, StyleDefinition, 
    ParagraphSkeleton, HeadingReference, FieldReference, TableSkeleton,
    FontSpec, ParagraphSpec, MediaReference, ContentControlReference,
    FormulaReference, ChartReference, FootnoteReference, EndnoteReference,
    CrossReference, StyleType, LineSpacingMode
)
from ..exceptions import ExtractionError


logger = logging.getLogger(__name__)


class DocumentExtractor:
    """Extract document structure and inventory with zero information loss."""
    
    def __init__(self, visible: bool = False):
        """
        Initialize document extractor.
        
        Args:
            visible: Whether to show Word application window
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
            logger.debug("COM initialized for extraction")
            
            # Create Word application instance
            self._word_app = win32.gencache.EnsureDispatch('Word.Application')
            self._word_app.Visible = self.visible
            self._word_app.DisplayAlerts = 0  # Disable alerts
            self._word_app.ScreenUpdating = False  # Improve performance
            
            logger.info(f"Word application started for extraction (visible={self.visible})")
            return self
            
        except Exception as e:
            self._cleanup()
            raise ExtractionError(f"Failed to initialize Word COM for extraction: {e}")
    
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
                logger.info("Word application closed after extraction")
        except Exception as e:
            logger.warning(f"Error during Word cleanup: {e}")
        
        finally:
            if self._com_initialized:
                try:
                    pythoncom.CoUninitialize()
                    self._com_initialized = False
                    logger.debug("COM uninitialized after extraction")
                except Exception as e:
                    logger.warning(f"Error during COM cleanup: {e}")
    
    def extract_structure(self, docx_path: str) -> StructureV1:
        """
        Extract document skeleton and metadata.
        
        Args:
            docx_path: Path to DOCX file
            
        Returns:
            StructureV1: Document structure with skeleton data
            
        Raises:
            ExtractionError: If extraction fails
        """
        if not os.path.exists(docx_path):
            raise ExtractionError(f"DOCX file not found: {docx_path}", docx_path=docx_path)
        
        try:
            logger.info(f"Extracting structure from: {docx_path}")
            
            # Open document
            doc = self._word_app.Documents.Open(docx_path, ReadOnly=True)
            
            try:
                # Extract metadata
                metadata = self._extract_metadata(doc)
                
                # Extract styles
                styles = self._extract_styles(doc)
                
                # Extract paragraphs with skeleton data
                paragraphs = self._extract_paragraphs(doc)
                
                # Extract headings
                headings = self._extract_headings(doc, paragraphs)
                
                # Extract fields
                fields = self._extract_fields(doc)
                
                # Extract tables
                tables = self._extract_tables(doc)
                
                # Create structure object
                structure = StructureV1(
                    metadata=metadata,
                    styles=styles,
                    paragraphs=paragraphs,
                    headings=headings,
                    fields=fields,
                    tables=tables
                )
                
                logger.info(f"Structure extracted: {len(paragraphs)} paragraphs, {len(headings)} headings, {len(styles)} styles")
                return structure
                
            finally:
                doc.Close(SaveChanges=0)
                
        except Exception as e:
            if isinstance(e, ExtractionError):
                raise
            raise ExtractionError(
                f"Failed to extract structure: {e}",
                docx_path=docx_path,
                extraction_stage="structure"
            )
    
    def extract_inventory(self, docx_path: str) -> InventoryFullV1:
        """
        Extract complete inventory including OOXML fragments.
        
        Args:
            docx_path: Path to DOCX file
            
        Returns:
            InventoryFullV1: Complete inventory with OOXML fragments
            
        Raises:
            ExtractionError: If extraction fails
        """
        if not os.path.exists(docx_path):
            raise ExtractionError(f"DOCX file not found: {docx_path}", docx_path=docx_path, extraction_stage="inventory")
        
        try:
            logger.info(f"Extracting inventory from: {docx_path}")
            
            # Extract OOXML fragments by parsing the DOCX as ZIP
            ooxml_fragments = self._extract_ooxml_fragments(docx_path)
            
            # Extract media indexes
            media_indexes = self._extract_media_indexes(docx_path)
            
            # Open document for COM-based extraction
            doc = self._word_app.Documents.Open(docx_path, ReadOnly=True)
            
            try:
                # Extract content controls
                content_controls = self._extract_content_controls(doc)
                
                # Extract formulas
                formulas = self._extract_formulas(doc)
                
                # Extract charts
                charts = self._extract_charts(doc)
                
                # Extract footnotes
                footnotes = self._extract_footnotes(doc)
                
                # Extract endnotes
                endnotes = self._extract_endnotes(doc)
                
                # Extract cross-references
                cross_references = self._extract_cross_references(doc)
                
                # Create inventory object
                inventory = InventoryFullV1(
                    ooxml_fragments=ooxml_fragments,
                    media_indexes=media_indexes,
                    content_controls=content_controls,
                    formulas=formulas,
                    charts=charts,
                    footnotes=footnotes,
                    endnotes=endnotes,
                    cross_references=cross_references
                )
                
                logger.info(f"Inventory extracted: {len(ooxml_fragments)} OOXML fragments, {len(media_indexes)} media files, "
                          f"{len(content_controls)} content controls, {len(formulas)} formulas, {len(charts)} charts, "
                          f"{len(footnotes)} footnotes, {len(endnotes)} endnotes, {len(cross_references)} cross-references")
                return inventory
                
            finally:
                doc.Close(SaveChanges=0)
                
        except Exception as e:
            if isinstance(e, ExtractionError):
                raise
            raise ExtractionError(
                f"Failed to extract inventory: {e}",
                docx_path=docx_path,
                extraction_stage="inventory"
            )
    
    def process_document(self, docx_path: str) -> Tuple[StructureV1, InventoryFullV1]:
        """
        Main extraction pipeline.
        
        Args:
            docx_path: Path to DOCX file
            
        Returns:
            Tuple of StructureV1 and InventoryFullV1
            
        Raises:
            ExtractionError: If extraction fails
        """
        logger.info(f"Processing document: {docx_path}")
        
        try:
            # Extract structure and inventory
            structure = self.extract_structure(docx_path)
            inventory = self.extract_inventory(docx_path)
            
            logger.info(f"Document processing completed: {docx_path}")
            return structure, inventory
            
        except Exception as e:
            if isinstance(e, ExtractionError):
                raise
            raise ExtractionError(
                f"Failed to process document: {e}",
                docx_path=docx_path
            )
    
    def _extract_metadata(self, doc) -> DocumentMetadata:
        """Extract document metadata from Word COM object."""
        try:
            # Get built-in document properties
            props = doc.BuiltInDocumentProperties
            
            # Extract basic metadata
            title = None
            author = None
            creation_time = None
            modified_time = None
            
            try:
                title = props("Title").Value if props("Title").Value else None
            except:
                pass
            
            try:
                author = props("Author").Value if props("Author").Value else None
            except:
                pass
            
            try:
                creation_time = props("Creation Date").Value
                if creation_time:
                    creation_time = datetime.fromisoformat(str(creation_time).replace('Z', '+00:00'))
            except:
                pass
            
            try:
                modified_time = props("Last Save Time").Value
                if modified_time:
                    modified_time = datetime.fromisoformat(str(modified_time).replace('Z', '+00:00'))
            except:
                pass
            
            # Get Word version
            word_version = None
            try:
                word_version = self._word_app.Version
            except:
                pass
            
            # Get document statistics
            page_count = None
            paragraph_count = None
            word_count = None
            
            try:
                page_count = doc.ComputeStatistics(2)  # wdStatisticPages
            except:
                pass
            
            try:
                paragraph_count = doc.ComputeStatistics(1)  # wdStatisticParagraphs
            except:
                pass
            
            try:
                word_count = doc.ComputeStatistics(0)  # wdStatisticWords
            except:
                pass
            
            return DocumentMetadata(
                title=title,
                author=author,
                creation_time=creation_time,
                modified_time=modified_time,
                word_version=word_version,
                page_count=page_count,
                paragraph_count=paragraph_count,
                word_count=word_count
            )
            
        except Exception as e:
            logger.warning(f"Failed to extract some metadata: {e}")
            return DocumentMetadata()
    
    def _extract_styles(self, doc) -> List[StyleDefinition]:
        """Extract style definitions from document."""
        styles = []
        
        try:
            for style in doc.Styles:
                try:
                    # Skip built-in styles that are not commonly used
                    if style.BuiltIn and style.InUse == False:
                        continue
                    
                    # Determine style type
                    style_type = StyleType.PARAGRAPH
                    try:
                        if style.Type == 2:  # wdStyleTypeCharacter
                            style_type = StyleType.CHARACTER
                        elif style.Type == 3:  # wdStyleTypeTable
                            style_type = StyleType.TABLE
                        elif style.Type == 4:  # wdStyleTypeLinked
                            style_type = StyleType.LINKED
                    except:
                        pass
                    
                    # Extract font specification
                    font_spec = None
                    try:
                        font = style.Font
                        font_spec = FontSpec(
                            east_asian=getattr(font, 'NameFarEast', None),
                            latin=getattr(font, 'Name', None),
                            size_pt=int(font.Size) if font.Size and font.Size > 0 else None,
                            bold=font.Bold if font.Bold != -9999999 else None,
                            italic=font.Italic if font.Italic != -9999999 else None,
                            color_hex=self._rgb_to_hex(font.Color) if font.Color != -9999999 else None
                        )
                    except:
                        pass
                    
                    # Extract paragraph specification
                    paragraph_spec = None
                    try:
                        if style.Type in [1, 4]:  # wdStyleTypeParagraph, wdStyleTypeLinked
                            para_format = style.ParagraphFormat
                            
                            # Determine line spacing mode
                            line_spacing_mode = None
                            line_spacing_value = None
                            
                            if para_format.LineSpacingRule == 0:  # wdLineSpaceSingle
                                line_spacing_mode = LineSpacingMode.SINGLE
                                line_spacing_value = 1.0
                            elif para_format.LineSpacingRule == 5:  # wdLineSpaceMultiple
                                line_spacing_mode = LineSpacingMode.MULTIPLE
                                line_spacing_value = para_format.LineSpacing
                            elif para_format.LineSpacingRule == 4:  # wdLineSpaceExactly
                                line_spacing_mode = LineSpacingMode.EXACTLY
                                line_spacing_value = para_format.LineSpacing
                            
                            paragraph_spec = ParagraphSpec(
                                line_spacing_mode=line_spacing_mode,
                                line_spacing_value=line_spacing_value,
                                space_before_pt=para_format.SpaceBefore if para_format.SpaceBefore > 0 else None,
                                space_after_pt=para_format.SpaceAfter if para_format.SpaceAfter > 0 else None,
                                indent_left_pt=para_format.LeftIndent if para_format.LeftIndent != 0 else None,
                                indent_right_pt=para_format.RightIndent if para_format.RightIndent != 0 else None,
                                indent_first_line_pt=para_format.FirstLineIndent if para_format.FirstLineIndent != 0 else None
                            )
                    except:
                        pass
                    
                    # Get base style information
                    based_on = None
                    next_style = None
                    try:
                        if hasattr(style, 'BaseStyle') and style.BaseStyle:
                            base_style_name = getattr(style.BaseStyle, 'NameLocal', None)
                            if base_style_name and isinstance(base_style_name, str):
                                based_on = base_style_name
                    except:
                        pass
                    
                    try:
                        if hasattr(style, 'NextParagraphStyle') and style.NextParagraphStyle:
                            next_style_name = getattr(style.NextParagraphStyle, 'NameLocal', None)
                            if next_style_name and isinstance(next_style_name, str):
                                next_style = next_style_name
                    except:
                        pass
                    
                    style_def = StyleDefinition(
                        name=style.NameLocal,
                        type=style_type,
                        font=font_spec,
                        paragraph=paragraph_spec,
                        based_on=based_on,
                        next_style=next_style
                    )
                    
                    styles.append(style_def)
                    
                except Exception as e:
                    logger.warning(f"Failed to extract style '{getattr(style, 'NameLocal', 'unknown')}': {e}")
                    continue
            
        except Exception as e:
            logger.warning(f"Failed to extract styles: {e}")
        
        return styles
    
    def _extract_paragraphs(self, doc) -> List[ParagraphSkeleton]:
        """Extract paragraph skeletons with preview text only."""
        paragraphs = []
        
        try:
            for i, para in enumerate(doc.Paragraphs):
                try:
                    # Get paragraph text and truncate to 120 characters
                    text = para.Range.Text.strip()
                    preview_text = text[:120] if len(text) > 120 else text
                    
                    # Get style name
                    style_name = None
                    try:
                        style_name = para.Style.NameLocal
                    except:
                        pass
                    
                    # Check if it's a heading
                    is_heading = False
                    heading_level = None
                    
                    try:
                        if para.OutlineLevel != 10:  # wdOutlineLevelBodyText
                            is_heading = True
                            heading_level = para.OutlineLevel
                    except:
                        pass
                    
                    # Alternative heading detection by style name
                    if not is_heading and style_name:
                        if any(heading_name in style_name.lower() for heading_name in ['heading', '标题']):
                            is_heading = True
                            # Try to extract level from style name
                            for level in range(1, 10):
                                if str(level) in style_name:
                                    heading_level = level
                                    break
                    
                    paragraph = ParagraphSkeleton(
                        index=i,
                        style_name=style_name,
                        preview_text=preview_text,
                        is_heading=is_heading,
                        heading_level=heading_level
                    )
                    
                    paragraphs.append(paragraph)
                    
                except Exception as e:
                    logger.warning(f"Failed to extract paragraph {i}: {e}")
                    continue
            
        except Exception as e:
            logger.warning(f"Failed to extract paragraphs: {e}")
        
        return paragraphs
    
    def _extract_headings(self, doc, paragraphs: List[ParagraphSkeleton]) -> List[HeadingReference]:
        """Extract heading references from paragraphs."""
        headings = []
        
        try:
            for para_skeleton in paragraphs:
                if para_skeleton.is_heading and para_skeleton.heading_level:
                    heading = HeadingReference(
                        paragraph_index=para_skeleton.index,
                        level=para_skeleton.heading_level,
                        text=para_skeleton.preview_text,
                        style_name=para_skeleton.style_name
                    )
                    headings.append(heading)
            
        except Exception as e:
            logger.warning(f"Failed to extract headings: {e}")
        
        return headings
    
    def _extract_fields(self, doc) -> List[FieldReference]:
        """Extract field references (TOC, page numbers, etc.)."""
        fields = []
        
        try:
            for field in doc.Fields:
                try:
                    # Find the paragraph containing this field
                    paragraph_index = 0
                    try:
                        # Get the paragraph containing the field
                        field_para = field.Range.Paragraphs(1)
                        # Find the index of this paragraph in the document
                        for i, para in enumerate(doc.Paragraphs):
                            if para.Range.Start == field_para.Range.Start:
                                paragraph_index = i
                                break
                    except:
                        pass
                    
                    # Get field type
                    field_type = "unknown"
                    try:
                        field_type = field.Type
                        # Convert numeric type to string if possible
                        if hasattr(win32_constants, f'wdField{field_type}'):
                            field_type = f'wdField{field_type}'
                    except:
                        pass
                    
                    # Get field code and result
                    field_code = None
                    result_text = None
                    
                    try:
                        field_code = field.Code.Text.strip()
                    except:
                        pass
                    
                    try:
                        result_text = field.Result.Text.strip()
                    except:
                        pass
                    
                    field_ref = FieldReference(
                        paragraph_index=paragraph_index,
                        field_type=str(field_type),
                        field_code=field_code,
                        result_text=result_text
                    )
                    
                    fields.append(field_ref)
                    
                except Exception as e:
                    logger.warning(f"Failed to extract field: {e}")
                    continue
            
        except Exception as e:
            logger.warning(f"Failed to extract fields: {e}")
        
        return fields
    
    def _extract_tables(self, doc) -> List[TableSkeleton]:
        """Extract table skeletons with basic structure."""
        tables = []
        
        try:
            for table in doc.Tables:
                try:
                    # Find the paragraph containing this table
                    paragraph_index = 0
                    try:
                        table_range = table.Range
                        for i, para in enumerate(doc.Paragraphs):
                            if para.Range.Start >= table_range.Start and para.Range.End <= table_range.End:
                                paragraph_index = i
                                break
                    except:
                        pass
                    
                    # Get table dimensions
                    rows = table.Rows.Count
                    columns = table.Columns.Count
                    
                    # Check if table has header
                    has_header = False
                    try:
                        if table.Rows.Count > 0:
                            first_row = table.Rows(1)
                            # Simple heuristic: if first row has different formatting, it's likely a header
                            has_header = True  # Simplified assumption
                    except:
                        pass
                    
                    # Extract cell references (paragraph indexes of cells) with row/column mapping
                    cell_references = []
                    cell_paragraph_map = {}  # Maps (row, col) to list of paragraph indexes
                    
                    try:
                        for row_idx in range(1, table.Rows.Count + 1):
                            for col_idx in range(1, table.Columns.Count + 1):
                                try:
                                    cell = table.Cell(row_idx, col_idx)
                                    cell_paragraphs = []
                                    
                                    # Find all paragraphs in this cell
                                    for para in cell.Range.Paragraphs:
                                        for i, doc_para in enumerate(doc.Paragraphs):
                                            if (doc_para.Range.Start <= para.Range.Start and 
                                                doc_para.Range.End >= para.Range.End):
                                                cell_paragraphs.append(i)
                                                if i not in cell_references:
                                                    cell_references.append(i)
                                                break
                                    
                                    if cell_paragraphs:
                                        cell_paragraph_map[f"{row_idx},{col_idx}"] = cell_paragraphs
                                        
                                except Exception as cell_e:
                                    logger.warning(f"Failed to extract cell ({row_idx}, {col_idx}): {cell_e}")
                                    continue
                    except Exception as e:
                        logger.warning(f"Failed to extract table cell references: {e}")
                        pass
                    
                    table_skeleton = TableSkeleton(
                        paragraph_index=paragraph_index,
                        rows=rows,
                        columns=columns,
                        has_header=has_header,
                        cell_references=cell_references if cell_references else None,
                        cell_paragraph_map=cell_paragraph_map if cell_paragraph_map else None
                    )
                    
                    tables.append(table_skeleton)
                    
                except Exception as e:
                    logger.warning(f"Failed to extract table: {e}")
                    continue
            
        except Exception as e:
            logger.warning(f"Failed to extract tables: {e}")
        
        return tables
    
    def _extract_ooxml_fragments(self, docx_path: str) -> Dict[str, str]:
        """Extract OOXML fragments by parsing DOCX as ZIP with enhanced object handling."""
        ooxml_fragments = {}
        
        try:
            with zipfile.ZipFile(docx_path, 'r') as zip_file:
                # Extract key OOXML files
                key_files = [
                    'word/document.xml',
                    'word/styles.xml',
                    'word/numbering.xml',
                    'word/settings.xml',
                    'word/fontTable.xml',
                    'word/theme/theme1.xml',
                    'docProps/core.xml',
                    'docProps/app.xml',
                    'word/footnotes.xml',
                    'word/endnotes.xml'
                ]
                
                for file_path in key_files:
                    try:
                        if file_path in zip_file.namelist():
                            content = zip_file.read(file_path).decode('utf-8')
                            ooxml_fragments[file_path] = content
                    except Exception as e:
                        logger.warning(f"Failed to extract OOXML fragment {file_path}: {e}")
                
                # Extract relationship files
                for file_name in zip_file.namelist():
                    if file_name.endswith('.rels'):
                        try:
                            content = zip_file.read(file_name).decode('utf-8')
                            ooxml_fragments[file_name] = content
                        except Exception as e:
                            logger.warning(f"Failed to extract relationship file {file_name}: {e}")
                
                # Extract chart files
                for file_name in zip_file.namelist():
                    if file_name.startswith('word/charts/'):
                        try:
                            if file_name.endswith('.xml'):
                                content = zip_file.read(file_name).decode('utf-8')
                                ooxml_fragments[file_name] = content
                            else:
                                # For binary chart data, store as base64
                                import base64
                                content = base64.b64encode(zip_file.read(file_name)).decode('utf-8')
                                ooxml_fragments[file_name] = f"base64:{content}"
                        except Exception as e:
                            logger.warning(f"Failed to extract chart file {file_name}: {e}")
                
                # Extract drawing files (SmartArt, shapes, etc.)
                for file_name in zip_file.namelist():
                    if file_name.startswith('word/drawings/'):
                        try:
                            if file_name.endswith('.xml'):
                                content = zip_file.read(file_name).decode('utf-8')
                                ooxml_fragments[file_name] = content
                            else:
                                # For binary drawing data, store as base64
                                import base64
                                content = base64.b64encode(zip_file.read(file_name)).decode('utf-8')
                                ooxml_fragments[file_name] = f"base64:{content}"
                        except Exception as e:
                            logger.warning(f"Failed to extract drawing file {file_name}: {e}")
                
                # Extract embedded objects
                for file_name in zip_file.namelist():
                    if file_name.startswith('word/embeddings/'):
                        try:
                            # Embedded objects are typically binary, store as base64
                            import base64
                            content = base64.b64encode(zip_file.read(file_name)).decode('utf-8')
                            ooxml_fragments[file_name] = f"base64:{content}"
                        except Exception as e:
                            logger.warning(f"Failed to extract embedded object {file_name}: {e}")
                
                # Extract ActiveX controls
                for file_name in zip_file.namelist():
                    if file_name.startswith('word/activeX/'):
                        try:
                            if file_name.endswith('.xml'):
                                content = zip_file.read(file_name).decode('utf-8')
                                ooxml_fragments[file_name] = content
                            else:
                                # For binary ActiveX data, store as base64
                                import base64
                                content = base64.b64encode(zip_file.read(file_name)).decode('utf-8')
                                ooxml_fragments[file_name] = f"base64:{content}"
                        except Exception as e:
                            logger.warning(f"Failed to extract ActiveX file {file_name}: {e}")
                
                # Extract custom XML parts (often used for content controls)
                for file_name in zip_file.namelist():
                    if file_name.startswith('customXml/'):
                        try:
                            if file_name.endswith('.xml'):
                                content = zip_file.read(file_name).decode('utf-8')
                                ooxml_fragments[file_name] = content
                        except Exception as e:
                            logger.warning(f"Failed to extract custom XML {file_name}: {e}")
                
                # Extract VBA project if present
                for file_name in zip_file.namelist():
                    if file_name.startswith('word/vbaProject'):
                        try:
                            # VBA projects are binary, store as base64
                            import base64
                            content = base64.b64encode(zip_file.read(file_name)).decode('utf-8')
                            ooxml_fragments[file_name] = f"base64:{content}"
                        except Exception as e:
                            logger.warning(f"Failed to extract VBA project {file_name}: {e}")
            
        except Exception as e:
            logger.warning(f"Failed to extract OOXML fragments: {e}")
        
        return ooxml_fragments
    
    def _extract_media_indexes(self, docx_path: str) -> Dict[str, MediaReference]:
        """Extract media file indexes from DOCX."""
        media_indexes = {}
        
        try:
            with zipfile.ZipFile(docx_path, 'r') as zip_file:
                # Find media files
                for file_name in zip_file.namelist():
                    if file_name.startswith('word/media/'):
                        try:
                            file_info = zip_file.getinfo(file_name)
                            
                            # Extract file extension
                            file_extension = Path(file_name).suffix.lower()
                            
                            # Determine content type
                            content_type = "application/octet-stream"
                            if file_extension in ['.jpg', '.jpeg']:
                                content_type = "image/jpeg"
                            elif file_extension == '.png':
                                content_type = "image/png"
                            elif file_extension == '.gif':
                                content_type = "image/gif"
                            elif file_extension == '.bmp':
                                content_type = "image/bmp"
                            elif file_extension == '.wmf':
                                content_type = "image/wmf"
                            elif file_extension == '.emf':
                                content_type = "image/emf"
                            
                            media_ref = MediaReference(
                                media_id=Path(file_name).stem,
                                content_type=content_type,
                                file_extension=file_extension,
                                size_bytes=file_info.file_size
                            )
                            
                            media_indexes[file_name] = media_ref
                            
                        except Exception as e:
                            logger.warning(f"Failed to extract media reference {file_name}: {e}")
            
        except Exception as e:
            logger.warning(f"Failed to extract media indexes: {e}")
        
        return media_indexes
    
    def _extract_content_controls(self, doc) -> List[ContentControlReference]:
        """Extract content control references."""
        content_controls = []
        
        try:
            for cc in doc.ContentControls:
                try:
                    # Find the paragraph containing this content control
                    paragraph_index = 0
                    try:
                        cc_range = cc.Range
                        for i, para in enumerate(doc.Paragraphs):
                            if para.Range.Start <= cc_range.Start <= para.Range.End:
                                paragraph_index = i
                                break
                    except:
                        pass
                    
                    # Get content control properties
                    control_id = str(cc.ID) if hasattr(cc, 'ID') else f"cc_{len(content_controls)}"
                    control_type = str(cc.Type) if hasattr(cc, 'Type') else "unknown"
                    title = cc.Title if hasattr(cc, 'Title') and cc.Title else None
                    tag = cc.Tag if hasattr(cc, 'Tag') and cc.Tag else None
                    
                    cc_ref = ContentControlReference(
                        paragraph_index=paragraph_index,
                        control_id=control_id,
                        control_type=control_type,
                        title=title,
                        tag=tag
                    )
                    
                    content_controls.append(cc_ref)
                    
                except Exception as e:
                    logger.warning(f"Failed to extract content control: {e}")
                    continue
            
        except Exception as e:
            logger.warning(f"Failed to extract content controls: {e}")
        
        return content_controls
    
    def _extract_formulas(self, doc) -> List[FormulaReference]:
        """Extract formula references (equations, etc.) with enhanced detection."""
        formulas = []
        
        try:
            # Look for equation objects in inline shapes
            for shape in doc.InlineShapes:
                try:
                    # Check for various formula/equation types
                    is_formula = False
                    formula_type = "unknown"
                    
                    if shape.Type == 8:  # wdInlineShapeOLEControlObject
                        formula_type = "ole_object"
                        is_formula = True
                    elif shape.Type == 7:  # wdInlineShapeEmbeddedOLEObject
                        formula_type = "embedded_ole"
                        is_formula = True
                    elif shape.Type == 13:  # wdInlineShapeOLEEmbedded
                        formula_type = "ole_embedded"
                        is_formula = True
                    
                    if is_formula:
                        # Find the paragraph containing this shape
                        paragraph_index = 0
                        try:
                            shape_range = shape.Range
                            for i, para in enumerate(doc.Paragraphs):
                                if para.Range.Start <= shape_range.Start <= para.Range.End:
                                    paragraph_index = i
                                    break
                        except:
                            pass
                        
                        # Enhanced formula type detection
                        formula_id = f"formula_{len(formulas)}"
                        
                        try:
                            if hasattr(shape.OLEFormat, 'ClassType'):
                                class_type = shape.OLEFormat.ClassType.lower()
                                if 'equation' in class_type or 'math' in class_type:
                                    formula_type = "equation"
                                    formula_id = f"eq_{len(formulas)}"
                                elif 'excel' in class_type:
                                    formula_type = "excel_object"
                                    formula_id = f"excel_{len(formulas)}"
                                elif 'visio' in class_type:
                                    formula_type = "visio_object"
                                    formula_id = f"visio_{len(formulas)}"
                        except:
                            pass
                        
                        # Try to extract LaTeX or MathML if available
                        latex_code = None
                        try:
                            # This would require specialized extraction based on the OLE object type
                            # For now, we'll leave it as None but the structure is ready
                            pass
                        except:
                            pass
                        
                        formula_ref = FormulaReference(
                            paragraph_index=paragraph_index,
                            formula_id=formula_id,
                            formula_type=formula_type,
                            latex_code=latex_code
                        )
                        
                        formulas.append(formula_ref)
                        
                except Exception as e:
                    logger.warning(f"Failed to extract formula from inline shape: {e}")
                    continue
            
            # Also look for math zones (Word 2007+ native equations)
            try:
                # Look for OMath objects in the document
                for range_obj in doc.Range().OMaths:
                    try:
                        # Find the paragraph containing this math object
                        paragraph_index = 0
                        try:
                            math_range = range_obj.Range
                            for i, para in enumerate(doc.Paragraphs):
                                if para.Range.Start <= math_range.Start <= para.Range.End:
                                    paragraph_index = i
                                    break
                        except:
                            pass
                        
                        formula_id = f"omath_{len(formulas)}"
                        
                        # Try to extract LaTeX representation
                        latex_code = None
                        try:
                            # Word's OMath can be converted to LaTeX with additional processing
                            # For now, we'll store the raw text
                            latex_code = range_obj.Range.Text
                        except:
                            pass
                        
                        formula_ref = FormulaReference(
                            paragraph_index=paragraph_index,
                            formula_id=formula_id,
                            formula_type="omath",
                            latex_code=latex_code
                        )
                        
                        formulas.append(formula_ref)
                        
                    except Exception as e:
                        logger.warning(f"Failed to extract OMath formula: {e}")
                        continue
                        
            except Exception as e:
                logger.debug(f"OMath extraction not available or failed: {e}")
            
        except Exception as e:
            logger.warning(f"Failed to extract formulas: {e}")
        
        return formulas
    
    def _extract_charts(self, doc) -> List[ChartReference]:
        """Extract chart/SmartArt/OLE object references with enhanced detection."""
        charts = []
        
        try:
            # Look for various chart and graphic objects in inline shapes
            for shape in doc.InlineShapes:
                try:
                    is_chart_object = False
                    chart_type = "unknown"
                    chart_id = f"object_{len(charts)}"
                    title = None
                    
                    # Enhanced object type detection
                    if shape.Type == 12:  # wdInlineShapeChart
                        is_chart_object = True
                        chart_type = "chart"
                        chart_id = f"chart_{len(charts)}"
                        try:
                            if hasattr(shape.Chart, 'ChartTitle') and shape.Chart.ChartTitle:
                                title = shape.Chart.ChartTitle.Text
                        except:
                            pass
                            
                    elif shape.Type == 15:  # wdInlineShapeSmartArt
                        is_chart_object = True
                        chart_type = "smartart"
                        chart_id = f"smartart_{len(charts)}"
                        
                    elif shape.Type == 1:  # wdInlineShapePicture
                        is_chart_object = True
                        chart_type = "picture"
                        chart_id = f"picture_{len(charts)}"
                        
                    elif shape.Type == 3:  # wdInlineShapeLinkedPicture
                        is_chart_object = True
                        chart_type = "linked_picture"
                        chart_id = f"linked_pic_{len(charts)}"
                        
                    elif shape.Type == 5:  # wdInlineShapeLinkedOLEObject
                        is_chart_object = True
                        chart_type = "linked_ole"
                        chart_id = f"linked_ole_{len(charts)}"
                        
                    elif shape.Type == 7:  # wdInlineShapeEmbeddedOLEObject
                        is_chart_object = True
                        chart_type = "embedded_ole"
                        chart_id = f"embedded_ole_{len(charts)}"
                        
                    elif shape.Type == 8:  # wdInlineShapeOLEControlObject
                        is_chart_object = True
                        chart_type = "ole_control"
                        chart_id = f"ole_control_{len(charts)}"
                        
                    elif shape.Type == 14:  # wdInlineShapeWebVideo
                        is_chart_object = True
                        chart_type = "web_video"
                        chart_id = f"video_{len(charts)}"
                    
                    if is_chart_object:
                        # Find the paragraph containing this shape
                        paragraph_index = 0
                        try:
                            shape_range = shape.Range
                            for i, para in enumerate(doc.Paragraphs):
                                if para.Range.Start <= shape_range.Start <= para.Range.End:
                                    paragraph_index = i
                                    break
                        except:
                            pass
                        
                        # Try to extract additional metadata
                        try:
                            if hasattr(shape, 'AlternativeText') and shape.AlternativeText:
                                if not title and isinstance(shape.AlternativeText, str):
                                    title = shape.AlternativeText
                        except:
                            pass
                        
                        chart_ref = ChartReference(
                            paragraph_index=paragraph_index,
                            chart_id=chart_id,
                            chart_type=chart_type,
                            title=title
                        )
                        
                        charts.append(chart_ref)
                        
                except Exception as e:
                    logger.warning(f"Failed to extract chart/object from inline shape: {e}")
                    continue
            
            # Also look for shapes in the main document (not inline)
            try:
                for shape in doc.Shapes:
                    try:
                        is_chart_object = False
                        chart_type = "shape"
                        chart_id = f"shape_{len(charts)}"
                        title = None
                        
                        # Determine shape type
                        if hasattr(shape, 'Type'):
                            if shape.Type == 3:  # msoChart
                                is_chart_object = True
                                chart_type = "chart_shape"
                                chart_id = f"chart_shape_{len(charts)}"
                            elif shape.Type == 15:  # msoSmartArt
                                is_chart_object = True
                                chart_type = "smartart_shape"
                                chart_id = f"smartart_shape_{len(charts)}"
                            elif shape.Type == 13:  # msoPicture
                                is_chart_object = True
                                chart_type = "picture_shape"
                                chart_id = f"picture_shape_{len(charts)}"
                            elif shape.Type == 7:  # msoOLEControlObject
                                is_chart_object = True
                                chart_type = "ole_shape"
                                chart_id = f"ole_shape_{len(charts)}"
                        
                        if is_chart_object:
                            # For shapes, finding the exact paragraph is more complex
                            # We'll use the anchor range
                            paragraph_index = 0
                            try:
                                if hasattr(shape, 'Anchor'):
                                    anchor_range = shape.Anchor
                                    for i, para in enumerate(doc.Paragraphs):
                                        if para.Range.Start <= anchor_range.Start <= para.Range.End:
                                            paragraph_index = i
                                            break
                            except:
                                pass
                            
                            # Try to get title/name
                            try:
                                if hasattr(shape, 'Name') and shape.Name:
                                    title = shape.Name
                                elif hasattr(shape, 'AlternativeText') and shape.AlternativeText:
                                    title = shape.AlternativeText
                            except:
                                pass
                            
                            chart_ref = ChartReference(
                                paragraph_index=paragraph_index,
                                chart_id=chart_id,
                                chart_type=chart_type,
                                title=title
                            )
                            
                            charts.append(chart_ref)
                            
                    except Exception as e:
                        logger.warning(f"Failed to extract shape object: {e}")
                        continue
                        
            except Exception as e:
                logger.debug(f"Shape extraction not available or failed: {e}")
            
        except Exception as e:
            logger.warning(f"Failed to extract charts: {e}")
        
        return charts
    
    def _extract_footnotes(self, doc) -> List[FootnoteReference]:
        """Extract footnote references with preservation of relationships."""
        footnotes = []
        
        try:
            for footnote in doc.Footnotes:
                try:
                    # Find the paragraph containing the footnote reference
                    paragraph_index = 0
                    try:
                        footnote_range = footnote.Reference
                        for i, para in enumerate(doc.Paragraphs):
                            if para.Range.Start <= footnote_range.Start <= para.Range.End:
                                paragraph_index = i
                                break
                    except:
                        pass
                    
                    # Get footnote properties
                    footnote_id = f"footnote_{footnote.Index}"
                    reference_mark = str(footnote.Index)
                    
                    # Get footnote text preview (first 120 characters)
                    text_preview = None
                    try:
                        footnote_text = footnote.Range.Text.strip()
                        if footnote_text:
                            text_preview = footnote_text[:120] if len(footnote_text) > 120 else footnote_text
                    except:
                        pass
                    
                    footnote_ref = FootnoteReference(
                        paragraph_index=paragraph_index,
                        footnote_id=footnote_id,
                        reference_mark=reference_mark,
                        text_preview=text_preview
                    )
                    
                    footnotes.append(footnote_ref)
                    
                except Exception as e:
                    logger.warning(f"Failed to extract footnote: {e}")
                    continue
            
        except Exception as e:
            logger.warning(f"Failed to extract footnotes: {e}")
        
        return footnotes
    
    def _extract_endnotes(self, doc) -> List[EndnoteReference]:
        """Extract endnote references with preservation of relationships."""
        endnotes = []
        
        try:
            for endnote in doc.Endnotes:
                try:
                    # Find the paragraph containing the endnote reference
                    paragraph_index = 0
                    try:
                        endnote_range = endnote.Reference
                        for i, para in enumerate(doc.Paragraphs):
                            if para.Range.Start <= endnote_range.Start <= para.Range.End:
                                paragraph_index = i
                                break
                    except:
                        pass
                    
                    # Get endnote properties
                    endnote_id = f"endnote_{endnote.Index}"
                    reference_mark = str(endnote.Index)
                    
                    # Get endnote text preview (first 120 characters)
                    text_preview = None
                    try:
                        endnote_text = endnote.Range.Text.strip()
                        if endnote_text:
                            text_preview = endnote_text[:120] if len(endnote_text) > 120 else endnote_text
                    except:
                        pass
                    
                    endnote_ref = EndnoteReference(
                        paragraph_index=paragraph_index,
                        endnote_id=endnote_id,
                        reference_mark=reference_mark,
                        text_preview=text_preview
                    )
                    
                    endnotes.append(endnote_ref)
                    
                except Exception as e:
                    logger.warning(f"Failed to extract endnote: {e}")
                    continue
            
        except Exception as e:
            logger.warning(f"Failed to extract endnotes: {e}")
        
        return endnotes
    
    def _extract_cross_references(self, doc) -> List[CrossReference]:
        """Extract cross-reference identification and relationship mapping."""
        cross_references = []
        
        try:
            # Look for cross-reference fields in the document
            for field in doc.Fields:
                try:
                    if field.Type == 37:  # wdFieldRef (cross-reference field)
                        # Find the paragraph containing this field
                        source_paragraph_index = 0
                        try:
                            field_range = field.Range
                            for i, para in enumerate(doc.Paragraphs):
                                if para.Range.Start <= field_range.Start <= para.Range.End:
                                    source_paragraph_index = i
                                    break
                        except:
                            pass
                        
                        # Parse the field code to understand the reference
                        field_code = None
                        reference_type = "unknown"
                        reference_text = ""
                        target_id = None
                        target_paragraph_index = None
                        
                        try:
                            field_code = field.Code.Text.strip()
                            reference_text = field.Result.Text.strip()
                            
                            # Parse field code to determine reference type
                            if 'REF' in field_code.upper():
                                if '_Ref' in field_code:
                                    reference_type = "bookmark"
                                    # Extract bookmark name
                                    parts = field_code.split()
                                    for part in parts:
                                        if '_Ref' in part:
                                            target_id = part
                                            break
                                elif '_Toc' in field_code:
                                    reference_type = "toc_entry"
                                    # Extract TOC entry reference
                                    parts = field_code.split()
                                    for part in parts:
                                        if '_Toc' in part:
                                            target_id = part
                                            break
                                else:
                                    reference_type = "heading"
                            
                        except:
                            pass
                        
                        # Try to find the target paragraph
                        if target_id:
                            try:
                                # Look for bookmarks with matching names
                                for bookmark in doc.Bookmarks:
                                    if bookmark.Name == target_id:
                                        bookmark_range = bookmark.Range
                                        for i, para in enumerate(doc.Paragraphs):
                                            if para.Range.Start <= bookmark_range.Start <= para.Range.End:
                                                target_paragraph_index = i
                                                break
                                        break
                            except:
                                pass
                        
                        cross_ref = CrossReference(
                            source_paragraph_index=source_paragraph_index,
                            target_paragraph_index=target_paragraph_index,
                            reference_type=reference_type,
                            reference_text=reference_text,
                            target_id=target_id
                        )
                        
                        cross_references.append(cross_ref)
                        
                except Exception as e:
                    logger.warning(f"Failed to extract cross-reference field: {e}")
                    continue
            
            # Also look for hyperlinks as a form of cross-reference
            try:
                for hyperlink in doc.Hyperlinks:
                    try:
                        # Find the paragraph containing this hyperlink
                        source_paragraph_index = 0
                        try:
                            hyperlink_range = hyperlink.Range
                            for i, para in enumerate(doc.Paragraphs):
                                if para.Range.Start <= hyperlink_range.Start <= para.Range.End:
                                    source_paragraph_index = i
                                    break
                        except:
                            pass
                        
                        # Get hyperlink properties
                        reference_text = hyperlink.TextToDisplay if hyperlink.TextToDisplay else ""
                        target_id = hyperlink.Address if hyperlink.Address else hyperlink.SubAddress
                        reference_type = "hyperlink"
                        
                        # Check if it's an internal reference
                        if hyperlink.SubAddress:
                            reference_type = "internal_link"
                            target_id = hyperlink.SubAddress
                            
                            # Try to find the target paragraph for internal links
                            target_paragraph_index = None
                            try:
                                # Look for bookmarks or headings with matching names
                                for bookmark in doc.Bookmarks:
                                    if bookmark.Name == target_id:
                                        bookmark_range = bookmark.Range
                                        for i, para in enumerate(doc.Paragraphs):
                                            if para.Range.Start <= bookmark_range.Start <= para.Range.End:
                                                target_paragraph_index = i
                                                break
                                        break
                            except:
                                pass
                        else:
                            target_paragraph_index = None
                        
                        cross_ref = CrossReference(
                            source_paragraph_index=source_paragraph_index,
                            target_paragraph_index=target_paragraph_index,
                            reference_type=reference_type,
                            reference_text=reference_text,
                            target_id=target_id
                        )
                        
                        cross_references.append(cross_ref)
                        
                    except Exception as e:
                        logger.warning(f"Failed to extract hyperlink cross-reference: {e}")
                        continue
                        
            except Exception as e:
                logger.debug(f"Hyperlink extraction not available or failed: {e}")
            
        except Exception as e:
            logger.warning(f"Failed to extract cross-references: {e}")
        
        return cross_references
    
    def _rgb_to_hex(self, rgb_color) -> Optional[str]:
        """Convert RGB color to hex format."""
        try:
            if rgb_color == -9999999:  # Word's "automatic" color
                return None
            
            # Extract RGB components
            r = rgb_color & 0xFF
            g = (rgb_color >> 8) & 0xFF
            b = (rgb_color >> 16) & 0xFF
            
            return f"#{r:02X}{g:02X}{b:02X}"
        except:
            return None