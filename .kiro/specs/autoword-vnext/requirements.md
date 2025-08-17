# Requirements Document

## Introduction

AutoWord vNext transforms Word document processing into a structured "Extract→Plan→Execute→Validate→Audit" closed loop system. The system achieves >99% stability and reproducibility for document modifications based on annotations, ensuring no valid information is lost through a rigorous five-module pipeline with strict JSON schemas and atomic operations.

## Requirements

### Requirement 1

**User Story:** As a document processor, I want to extract complete document structure and inventory, so that I can create a comprehensive foundation for planning modifications.

#### Acceptance Criteria

1. WHEN processing a DOCX file THEN the system SHALL generate structure.json(v1) containing document skeleton and metadata
2. WHEN extracting structure THEN the system SHALL include styles, paragraphs, headings, fields, and tables with precise indexing
3. WHEN processing documents THEN the system SHALL generate inventory.full.json(v1) containing all sensitive/large objects as OOXML fragments
4. WHEN extracting content THEN the system SHALL preserve zero information loss through the inventory system
5. WHEN processing documents THEN the system SHALL export only skeleton and preview data, not full text content

### Requirement 2

**User Story:** As a document processor, I want LLM-generated execution plans in strict JSON format, so that I can ensure predictable and safe document modifications.

#### Acceptance Criteria

1. WHEN receiving structure.json(v1) and user intent THEN the system SHALL generate plan.json(v1) through LLM processing
2. WHEN LLM generates plans THEN the system SHALL output only valid JSON matching the plan.v1 schema
3. WHEN validating plans THEN the system SHALL reject any operations not in the predefined whitelist
4. WHEN LLM processing occurs THEN the system SHALL never directly generate DOCX/OOXML content
5. WHEN plan validation fails THEN the system SHALL return INVALID_PLAN status and reject execution

### Requirement 3

**User Story:** As a document processor, I want atomic operation execution through Word COM, so that I can apply modifications safely and predictably.

#### Acceptance Criteria

1. WHEN executing plans THEN the system SHALL perform only whitelisted atomic operations
2. WHEN executing delete_section_by_heading THEN the system SHALL remove content from matching heading to next same-level heading
3. WHEN executing set_style_rule THEN the system SHALL modify style objects with both EastAsian and Latin font settings
4. WHEN executing reassign_paragraphs_to_style THEN the system SHALL reassign paragraph instances to target styles
5. WHEN executing update_toc THEN the system SHALL refresh all TOC fields and update page numbers
6. WHEN executing clear_direct_formatting THEN the system SHALL require explicit authorization and apply to specified scope

### Requirement 4

**User Story:** As a document processor, I want comprehensive validation and rollback capabilities, so that I can ensure document integrity and recover from failures.

#### Acceptance Criteria

1. WHEN modifications complete THEN the system SHALL generate structure.json'(v1) for comparison validation
2. WHEN validating results THEN the system SHALL verify chapter assertions (no "摘要/参考文献" at level 1)
3. WHEN validating styles THEN the system SHALL confirm H1/H2/Normal styles match exact specifications
4. WHEN validating TOC THEN the system SHALL ensure TOC items and page numbers match heading tree
5. WHEN validation fails THEN the system SHALL rollback to original DOCX and mark FAILED_VALIDATION
6. WHEN processing completes THEN the system SHALL force Fields.Update() and Repaginate()

### Requirement 5

**User Story:** As a document processor, I want complete audit trails and snapshots, so that I can track all changes and maintain accountability.

#### Acceptance Criteria

1. WHEN processing documents THEN the system SHALL create timestamped run directories with fixed file names
2. WHEN processing occurs THEN the system SHALL save before/after DOCX files and structure JSON files
3. WHEN generating audit trails THEN the system SHALL create plan.json, diff.report.json, and warnings.log
4. WHEN completing processing THEN the system SHALL write result.status.txt with SUCCESS/ROLLBACK/FAILED_VALIDATION
5. WHEN logging warnings THEN the system SHALL record font fallbacks, NOOP operations, and other issues

### Requirement 6

**User Story:** As a document processor, I want localization support with font fallback, so that I can handle Chinese documents reliably across different systems.

#### Acceptance Criteria

1. WHEN processing documents THEN the system SHALL map style aliases (Heading 1 ↔ 标题 1, Normal ↔ 正文)
2. WHEN setting fonts THEN the system SHALL attempt English names first, then Chinese names as fallback
3. WHEN applying East Asian fonts THEN the system SHALL use font fallback table (楷体→楷体_GB2312→STKaiti)
4. WHEN fonts are unavailable THEN the system SHALL log to warnings.log but continue processing
5. WHEN font fallback occurs THEN the system SHALL use first available font from fallback chain

### Requirement 7

**User Story:** As a document processor, I want strict JSON schema validation, so that I can ensure data integrity and prevent malformed operations.

#### Acceptance Criteria

1. WHEN generating structure.json(v1) THEN the system SHALL include schema_version, metadata, styles, paragraphs, headings, fields, and tables
2. WHEN creating plan.json(v1) THEN the system SHALL validate against schema with only whitelisted operations
3. WHEN validating schemas THEN the system SHALL reject any JSON with unknown fields or invalid values
4. WHEN processing inventory THEN the system SHALL store OOXML fragments and media indexes in inventory.full.json(v1)
5. WHEN schema validation fails THEN the system SHALL halt processing and return detailed error messages

### Requirement 8

**User Story:** As a document processor, I want comprehensive error handling and recovery, so that I can maintain system stability under all conditions.

#### Acceptance Criteria

1. WHEN any exception occurs THEN the system SHALL stop processing, rollback changes, and write FAILED_VALIDATION
2. WHEN operations fail to find targets THEN the system SHALL log NOOP to warnings.log and continue
3. WHEN unauthorized operations are attempted THEN the system SHALL reject and log security violations
4. WHEN processing documents with revisions THEN the system SHALL handle according to configured revision strategy
5. WHEN system errors occur THEN the system SHALL provide detailed error context for troubleshooting

### Requirement 9

**User Story:** As a document processor, I want the system to handle complex document elements, so that I can process real-world documents without data loss.

#### Acceptance Criteria

1. WHEN processing documents with tables THEN the system SHALL locate and delete sections containing table headings
2. WHEN handling formulas and content controls THEN the system SHALL preserve objects through inventory system
3. WHEN processing documents with charts/SmartArt THEN the system SHALL store as OOXML/binary in inventory.full
4. WHEN encountering footnotes/endnotes THEN the system SHALL maintain references and content integrity
5. WHEN processing cross-references THEN the system SHALL identify and preserve reference relationships

### Requirement 10

**User Story:** As a document processor, I want runtime constraints and security controls, so that I can prevent unauthorized modifications and maintain document safety.

#### Acceptance Criteria

1. WHEN receiving plans THEN the system SHALL reject any operations not in the predefined whitelist
2. WHEN modifying documents THEN the system SHALL never use string replacement on content or TOC display text
3. WHEN applying changes THEN the system SHALL complete all modifications through Word object layer only
4. WHEN processing LLM output THEN the system SHALL validate JSON schema and reject non-compliant plans
5. WHEN handling user input THEN the system SHALL sanitize and validate all parameters before execution