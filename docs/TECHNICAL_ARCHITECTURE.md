# AutoWord vNext - Technical Architecture Documentation

## Overview

AutoWord vNext is a structured document processing pipeline that transforms Word document modifications into a predictable, reproducible system through a five-module architecture: Extract→Plan→Execute→Validate→Audit.

## Core Architecture Principles

### 1. Zero Information Loss
- Complete document structure captured in `structure.v1.json`
- Sensitive/large objects preserved in `inventory.full.v1.json`
- OOXML fragments stored for complex elements

### 2. Strict JSON Schema Validation
- All data interchange uses versioned JSON schemas
- LLM outputs validated against `plan.v1.json` schema
- Whitelist operation enforcement prevents unauthorized modifications

### 3. Atomic Operations Only
- All document modifications through predefined atomic operations
- Word COM object layer modifications only
- No direct text manipulation or string replacement

### 4. Complete Auditability
- Timestamped run directories with before/after snapshots
- Comprehensive diff reports and operation logs
- Status tracking: SUCCESS/ROLLBACK/FAILED_VALIDATION

## Module Architecture

```
┌─────────────────┐    ┌─────────────────┐    ┌─────────────────┐
│   Extractor     │───►│    Planner      │───►│   Executor      │
│                 │    │                 │    │                 │
│ DOCX →          │    │ structure.json  │    │ plan.json →     │
│ structure.json  │    │ + user_intent → │    │ modified DOCX   │
│ inventory.json  │    │ plan.json       │    │                 │
└─────────────────┘    └─────────────────┘    └─────────────────┘
                                │                        │
                       ┌─────────────────┐    ┌─────────────────┐
                       │   LLM Service   │    │   Validator     │
                       │                 │    │                 │
                       │ • JSON Schema   │    │ structure.json' │
                       │ • Whitelist     │    │ → assertions    │
                       │ • Constraints   │    │ → rollback      │
                       └─────────────────┘    └─────────────────┘
                                                        │
                                              ┌─────────────────┐
                                              │    Auditor      │
                                              │                 │
                                              │ • Snapshots     │
                                              │ • Diff Reports  │
                                              │ • Status Logs   │
                                              └─────────────────┘
```

## Module Specifications

### 1. Extractor Module (`autoword.vnext.extractor`)

**Purpose**: Convert DOCX to structured JSON with zero information loss.

**Key Classes**:
- `DocumentExtractor`: Main extraction orchestrator
- `StructureExtractor`: Document skeleton extraction
- `InventoryExtractor`: Complete object inventory

**Data Flow**:
```
DOCX File → Word COM → Structure Analysis → JSON Serialization
                    → Inventory Collection → OOXML Storage
```

**Output Schemas**:
- `structure.v1.json`: Document skeleton and metadata
- `inventory.full.v1.json`: Complete object inventory

### 2. Planner Module (`autoword.vnext.planner`)

**Purpose**: Generate execution plans through LLM with strict validation.

**Key Classes**:
- `DocumentPlanner`: LLM integration and plan generation
- `SchemaValidator`: JSON schema validation
- `WhitelistValidator`: Operation whitelist enforcement

**LLM Integration**:
- Supports GPT-4, Claude, and other JSON-capable models
- Strict prompt engineering for JSON-only output
- Schema validation gateway prevents malformed plans

**Security Controls**:
- Whitelist operation enforcement
- JSON schema validation
- No DOCX/OOXML generation allowed

### 3. Executor Module (`autoword.vnext.executor`)

**Purpose**: Execute atomic operations through Word COM.

**Key Classes**:
- `DocumentExecutor`: Main execution orchestrator
- `AtomicOperations`: Individual operation implementations
- `LocalizationManager`: Style aliases and font fallbacks

**Atomic Operations**:
- `delete_section_by_heading`: Section removal by heading match
- `update_toc`: TOC refresh and page number update
- `delete_toc`: TOC removal with mode selection
- `set_style_rule`: Style object modification
- `reassign_paragraphs_to_style`: Paragraph style reassignment
- `clear_direct_formatting`: Direct formatting removal

### 4. Validator Module (`autoword.vnext.validator`)

**Purpose**: Validate modifications with rollback capability.

**Key Classes**:
- `DocumentValidator`: Main validation orchestrator
- `AssertionChecker`: Specific assertion implementations
- `RollbackManager`: Document restoration

**Validation Assertions**:
- Chapter assertions: No "摘要/参考文献" at level 1
- Style assertions: H1/H2/Normal specifications
- TOC assertions: Consistency with heading tree
- Pagination assertions: Field updates and repagination

### 5. Auditor Module (`autoword.vnext.auditor`)

**Purpose**: Create complete audit trails.

**Key Classes**:
- `DocumentAuditor`: Main audit orchestrator
- `SnapshotManager`: Before/after file management
- `DiffGenerator`: Structural difference analysis

**Audit Outputs**:
- Timestamped run directories
- Before/after DOCX snapshots
- Structure JSON comparisons
- Diff reports and operation logs

## Data Models and Schemas

### Structure V1 Schema

```json
{
  "schema_version": "structure.v1",
  "metadata": {
    "title": "string",
    "author": "string",
    "created_time": "ISO8601",
    "modified_time": "ISO8601",
    "word_version": "string"
  },
  "styles": [
    {
      "name": "string",
      "type": "paragraph|character|table|linked",
      "font": {
        "east_asian": "string",
        "latin": "string",
        "size_pt": "number",
        "bold": "boolean",
        "italic": "boolean",
        "color_hex": "string"
      },
      "paragraph": {
        "line_spacing_mode": "SINGLE|MULTIPLE|EXACTLY",
        "line_spacing_value": "number"
      }
    }
  ],
  "paragraphs": [
    {
      "index": "number",
      "style_name": "string",
      "preview_text": "string (≤120 chars)",
      "is_heading": "boolean"
    }
  ],
  "headings": [
    {
      "text": "string",
      "level": "number (1-9)",
      "paragraph_index": "number",
      "occurrence_index": "number"
    }
  ],
  "fields": [
    {
      "type": "TOC|PAGE|DATE|etc",
      "paragraph_index": "number",
      "field_code": "string"
    }
  ],
  "tables": [
    {
      "paragraph_index": "number",
      "rows": "number",
      "columns": "number",
      "cell_references": ["array of paragraph indexes"]
    }
  ]
}
```

### Plan V1 Schema

```json
{
  "schema_version": "plan.v1",
  "ops": [
    {
      "op_type": "delete_section_by_heading",
      "heading_text": "string",
      "level": "number (1-9)",
      "match": "EXACT|CONTAINS|REGEX",
      "case_sensitive": "boolean",
      "occurrence_index": "number (optional)"
    },
    {
      "op_type": "set_style_rule",
      "target_style_name": "string",
      "font": {
        "east_asian": "string",
        "latin": "string",
        "size_pt": "number",
        "bold": "boolean"
      },
      "paragraph": {
        "line_spacing_mode": "string",
        "line_spacing_value": "number"
      }
    }
  ]
}
```

## Error Handling and Recovery

### Exception Hierarchy

```python
class VNextError(Exception):
    """Base exception for vNext pipeline"""

class ExtractionError(VNextError):
    """Errors during document extraction"""

class PlanningError(VNextError):
    """Errors during plan generation"""

class ExecutionError(VNextError):
    """Errors during plan execution"""

class ValidationError(VNextError):
    """Errors during validation"""
    def __init__(self, assertion_failures: List[str], rollback_path: str):
        self.assertion_failures = assertion_failures
        self.rollback_path = rollback_path
```

### Rollback Strategy

1. **Automatic Rollback**: Any exception triggers rollback to original DOCX
2. **Validation Rollback**: Failed assertions trigger rollback with detailed error report
3. **Status Tracking**: All outcomes logged as SUCCESS/ROLLBACK/FAILED_VALIDATION

## Performance Considerations

### Memory Management
- Stream processing for large documents
- Lazy loading of inventory objects
- Efficient JSON serialization
- Proper COM object disposal

### Optimization Strategies
- Batch similar atomic operations
- Cache document structure during session
- Minimize Word COM round-trips
- Efficient schema validation

### Scalability Limits
- Single document processing model
- Memory constraints for very large documents
- LLM context limits
- COM single-threading requirements

## Security Model

### LLM Interface Security
- JSON schema gateway validation
- Whitelist operation enforcement
- No DOCX/OOXML generation allowed
- Input sanitization and validation

### Word COM Security
- Object-layer modifications only
- No string replacement on content/TOC text
- Proper COM resource management
- Exception handling with rollback

### Audit Trail Security
- Complete before/after snapshots
- Immutable timestamped directories
- Detailed operation logging
- Status tracking integrity

## Integration Points

### LLM Service Integration
- Configurable LLM providers (OpenAI, Anthropic, etc.)
- Prompt template management
- Response validation and retry logic
- Rate limiting and error handling

### Word COM Integration
- Version compatibility (Word 2016+)
- COM object lifecycle management
- Error handling and recovery
- Resource cleanup and disposal

### File System Integration
- Timestamped directory creation
- Atomic file operations
- Backup and snapshot management
- Cross-platform path handling

## Testing Architecture

### Unit Testing
- Mock Word COM objects
- Schema validation testing
- Atomic operation isolation
- Error scenario coverage

### Integration Testing
- Real Word document processing
- LLM integration testing
- End-to-end pipeline validation
- Performance benchmarking

### Validation Testing
- All 7 DoD scenarios
- Complex document handling
- Error recovery testing
- Rollback verification

This technical architecture provides the foundation for understanding, maintaining, and extending the AutoWord vNext system.