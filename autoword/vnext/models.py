"""
Core data models for AutoWord vNext with Pydantic validation.

This module defines the structured data models used throughout the vNext pipeline:
- StructureV1: Document skeleton and metadata
- PlanV1: LLM-generated execution plans
- InventoryFullV1: Complete inventory with OOXML fragments
"""

from typing import Dict, List, Optional, Union, Any, Literal
from datetime import datetime
from pydantic import BaseModel, Field, field_validator, ConfigDict
from enum import Enum


class SchemaVersion(str, Enum):
    """Schema version enumeration for data model versioning."""
    STRUCTURE_V1 = "structure.v1"
    PLAN_V1 = "plan.v1"
    INVENTORY_FULL_V1 = "inventory.full.v1"


class LineSpacingMode(str, Enum):
    """Line spacing mode enumeration."""
    SINGLE = "SINGLE"
    MULTIPLE = "MULTIPLE"
    EXACTLY = "EXACTLY"


class MatchMode(str, Enum):
    """Text matching mode enumeration."""
    EXACT = "EXACT"
    CONTAINS = "CONTAINS"
    REGEX = "REGEX"


class StyleType(str, Enum):
    """Style type enumeration."""
    PARAGRAPH = "paragraph"
    CHARACTER = "character"
    TABLE = "table"
    LINKED = "linked"


# Core Data Models

class DocumentMetadata(BaseModel):
    """Document metadata extracted from DOCX properties."""
    title: Optional[str] = None
    author: Optional[str] = None
    creation_time: Optional[datetime] = None
    modified_time: Optional[datetime] = None
    word_version: Optional[str] = None
    page_count: Optional[int] = None
    paragraph_count: Optional[int] = None
    word_count: Optional[int] = None


class FontSpec(BaseModel):
    """Font specification for styles."""
    east_asian: Optional[str] = None
    latin: Optional[str] = None
    size_pt: Optional[int] = Field(None, ge=1, le=72)
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    color_hex: Optional[str] = Field(None, pattern=r"^#[0-9A-Fa-f]{6}$")


class ParagraphSpec(BaseModel):
    """Paragraph formatting specification."""
    line_spacing_mode: Optional[LineSpacingMode] = None
    line_spacing_value: Optional[float] = Field(None, ge=0.1, le=10.0)
    space_before_pt: Optional[float] = Field(None, ge=0, le=1000)
    space_after_pt: Optional[float] = Field(None, ge=0, le=1000)
    indent_left_pt: Optional[float] = Field(None, ge=0, le=1000)
    indent_right_pt: Optional[float] = Field(None, ge=0, le=1000)
    indent_first_line_pt: Optional[float] = Field(None, ge=-1000, le=1000)


class StyleDefinition(BaseModel):
    """Style definition with font and paragraph specifications."""
    name: str = Field(..., min_length=1, max_length=255)
    type: StyleType
    font: Optional[FontSpec] = None
    paragraph: Optional[ParagraphSpec] = None
    based_on: Optional[str] = None
    next_style: Optional[str] = None


class ParagraphSkeleton(BaseModel):
    """Paragraph skeleton with preview text only (≤120 chars)."""
    index: int = Field(..., ge=0)
    style_name: Optional[str] = None
    preview_text: str = Field(..., max_length=120)
    is_heading: bool = False
    heading_level: Optional[int] = Field(None, ge=1, le=9)
    
    @field_validator('preview_text')
    @classmethod
    def truncate_preview(cls, v):
        """Ensure preview text is truncated to 120 characters."""
        return v[:120] if len(v) > 120 else v


class HeadingReference(BaseModel):
    """Heading reference with level and text."""
    paragraph_index: int = Field(..., ge=0)
    level: int = Field(..., ge=1, le=9)
    text: str = Field(..., min_length=1, max_length=255)
    style_name: Optional[str] = None


class FieldReference(BaseModel):
    """Field reference (TOC, page numbers, etc.)."""
    paragraph_index: int = Field(..., ge=0)
    field_type: str = Field(..., min_length=1)
    field_code: Optional[str] = None
    result_text: Optional[str] = None


class TableSkeleton(BaseModel):
    """Table skeleton with basic structure."""
    paragraph_index: int = Field(..., ge=0)
    rows: int = Field(..., ge=1)
    columns: int = Field(..., ge=1)
    has_header: bool = False
    cell_references: Optional[List[int]] = None  # Paragraph indexes of cells
    cell_paragraph_map: Optional[Dict[str, List[int]]] = None  # Maps "row,col" to paragraph indexes


class StructureV1(BaseModel):
    """Document structure v1 schema - skeleton and metadata only."""
    schema_version: Literal["structure.v1"] = "structure.v1"
    metadata: DocumentMetadata
    styles: List[StyleDefinition] = Field(default_factory=list)
    paragraphs: List[ParagraphSkeleton] = Field(default_factory=list)
    headings: List[HeadingReference] = Field(default_factory=list)
    fields: List[FieldReference] = Field(default_factory=list)
    tables: List[TableSkeleton] = Field(default_factory=list)
    
    model_config = ConfigDict(use_enum_values=True, validate_assignment=True)


# Atomic Operations for Plan V1

class AtomicOperation(BaseModel):
    """Base class for atomic operations."""
    operation_type: str = Field(..., min_length=1)
    
    model_config = ConfigDict(extra="forbid")  # Reject unknown fields


class DeleteSectionByHeading(AtomicOperation):
    """Delete section by heading text operation."""
    operation_type: Literal["delete_section_by_heading"] = "delete_section_by_heading"
    heading_text: str = Field(..., min_length=1, max_length=255)
    level: int = Field(..., ge=1, le=9)
    match: MatchMode = MatchMode.EXACT
    case_sensitive: bool = False
    occurrence_index: Optional[int] = Field(None, ge=1)


class UpdateToc(AtomicOperation):
    """Update table of contents operation."""
    operation_type: Literal["update_toc"] = "update_toc"


class DeleteToc(AtomicOperation):
    """Delete table of contents operation."""
    operation_type: Literal["delete_toc"] = "delete_toc"
    mode: str = Field(default="all", pattern=r"^(all|first|last)$")


class SetStyleRule(AtomicOperation):
    """Set style rule operation."""
    operation_type: Literal["set_style_rule"] = "set_style_rule"
    target_style_name: str = Field(..., min_length=1, max_length=255)
    font: Optional[FontSpec] = None
    paragraph: Optional[ParagraphSpec] = None


class ReassignParagraphsToStyle(AtomicOperation):
    """Reassign paragraphs to style operation."""
    operation_type: Literal["reassign_paragraphs_to_style"] = "reassign_paragraphs_to_style"
    selector: Dict[str, Any] = Field(..., min_length=1)
    target_style_name: str = Field(..., min_length=1, max_length=255)
    clear_direct_formatting: bool = False


class ClearDirectFormatting(AtomicOperation):
    """Clear direct formatting operation."""
    operation_type: Literal["clear_direct_formatting"] = "clear_direct_formatting"
    scope: str = Field(..., pattern=r"^(document|selection|range)$")
    range_spec: Optional[Dict[str, Any]] = None
    authorization_required: Literal[True] = True


# Union type for all atomic operations
AtomicOperationUnion = Union[
    DeleteSectionByHeading,
    UpdateToc,
    DeleteToc,
    SetStyleRule,
    ReassignParagraphsToStyle,
    ClearDirectFormatting
]


class PlanV1(BaseModel):
    """Execution plan v1 schema with whitelisted atomic operations."""
    schema_version: Literal["plan.v1"] = "plan.v1"
    ops: List[AtomicOperationUnion] = Field(default_factory=list)
    
    model_config = ConfigDict(use_enum_values=True, validate_assignment=True)


# Inventory Models

class MediaReference(BaseModel):
    """Media file reference in inventory."""
    media_id: str = Field(..., min_length=1)
    content_type: str = Field(..., min_length=1)
    file_extension: Optional[str] = None
    size_bytes: Optional[int] = Field(None, ge=0)


class ContentControlReference(BaseModel):
    """Content control reference."""
    paragraph_index: int = Field(..., ge=0)
    control_id: str = Field(..., min_length=1)
    control_type: str = Field(..., min_length=1)
    title: Optional[str] = None
    tag: Optional[str] = None


class FormulaReference(BaseModel):
    """Formula reference (equations, etc.)."""
    paragraph_index: int = Field(..., ge=0)
    formula_id: str = Field(..., min_length=1)
    formula_type: str = Field(..., min_length=1)
    latex_code: Optional[str] = None


class ChartReference(BaseModel):
    """Chart/SmartArt reference."""
    paragraph_index: int = Field(..., ge=0)
    chart_id: str = Field(..., min_length=1)
    chart_type: str = Field(..., min_length=1)
    title: Optional[str] = None


class FootnoteReference(BaseModel):
    """Footnote reference."""
    paragraph_index: int = Field(..., ge=0)
    footnote_id: str = Field(..., min_length=1)
    reference_mark: str = Field(..., min_length=1)
    text_preview: Optional[str] = Field(None, max_length=120)


class EndnoteReference(BaseModel):
    """Endnote reference."""
    paragraph_index: int = Field(..., ge=0)
    endnote_id: str = Field(..., min_length=1)
    reference_mark: str = Field(..., min_length=1)
    text_preview: Optional[str] = Field(None, max_length=120)


class CrossReference(BaseModel):
    """Cross-reference relationship mapping."""
    source_paragraph_index: int = Field(..., ge=0)
    target_paragraph_index: Optional[int] = Field(None, ge=0)
    reference_type: str = Field(..., min_length=1)  # heading, figure, table, etc.
    reference_text: str = Field(..., min_length=1)
    target_id: Optional[str] = None


class InventoryFullV1(BaseModel):
    """Complete inventory v1 schema with OOXML fragments."""
    schema_version: Literal["inventory.full.v1"] = "inventory.full.v1"
    ooxml_fragments: Dict[str, str] = Field(default_factory=dict)
    media_indexes: Dict[str, MediaReference] = Field(default_factory=dict)
    content_controls: List[ContentControlReference] = Field(default_factory=list)
    formulas: List[FormulaReference] = Field(default_factory=list)
    charts: List[ChartReference] = Field(default_factory=list)
    footnotes: List[FootnoteReference] = Field(default_factory=list)
    endnotes: List[EndnoteReference] = Field(default_factory=list)
    cross_references: List[CrossReference] = Field(default_factory=list)
    
    model_config = ConfigDict(use_enum_values=True, validate_assignment=True)


# Validation and Processing Results

class ValidationResult(BaseModel):
    """Validation result with detailed error information."""
    is_valid: bool
    errors: List[str] = Field(default_factory=list)
    warnings: List[str] = Field(default_factory=list)


class OperationResult(BaseModel):
    """Result of atomic operation execution."""
    success: bool
    operation_type: str
    message: Optional[str] = None
    warnings: List[str] = Field(default_factory=list)


class ProcessingResult(BaseModel):
    """Final processing result."""
    status: str = Field(..., pattern=r"^(SUCCESS|ROLLBACK|FAILED_VALIDATION|INVALID_PLAN)$")
    message: Optional[str] = None
    errors: List[str] = Field(default_factory=list)
    warnings: List[str] = Field(default_factory=list)
    audit_directory: Optional[str] = None


class DiffReport(BaseModel):
    """Structural difference report between before/after."""
    added_paragraphs: List[int] = Field(default_factory=list)
    removed_paragraphs: List[int] = Field(default_factory=list)
    modified_paragraphs: List[int] = Field(default_factory=list)
    style_changes: Dict[str, Dict[str, Any]] = Field(default_factory=dict)
    heading_changes: List[Dict[str, Any]] = Field(default_factory=list)
    summary: str = ""