# AutoWord vNext - Design Document

## Overview

AutoWord vNext implements a structured "Extract→Plan→Execute→Validate→Audit" closed loop system that transforms Word document processing into a predictable, reproducible pipeline. The system achieves >99% stability through strict JSON schemas, atomic operations, and comprehensive validation layers.

The core architectural principle is "zero information loss with complete auditability" - every document modification is planned, executed through whitelisted operations, validated against assertions, and fully audited with before/after snapshots.

## Architecture

### High-Level Pipeline Architecture

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

### Module Boundaries and Data Flow

```
Input: DOCX File
    ↓
[Extractor] → structure.v1.json + inventory.full.v1.json
    ↓
[Planner] + User Intent → plan.v1.json (LLM generated)
    ↓
[Executor] + Original DOCX → Modified DOCX (atomic operations only)
    ↓
[Validator] → structure'.v1.json + validation assertions
    ↓
[Auditor] → Complete audit trail in timestamped directory
```

## Components and Interfaces

### 1. Extractor Module

**Purpose**: Convert DOCX to structured JSON representations with zero information loss.

```python
class DocumentExtractor:
    def extract_structure(self, docx_path: str) -> StructureV1:
        """Extract document skeleton and metadata"""
        
    def extract_inventory(self, docx_path: str) -> InventoryFullV1:
        """Extract complete inventory including OOXML fragments"""
        
    def process_document(self, docx_path: str) -> Tuple[StructureV1, InventoryFullV1]:
        """Main extraction pipeline"""
```

**Key Responsibilities**:
- Parse DOCX structure (styles, paragraphs, headings, fields, tables)
- Generate structure.v1.json with skeleton data only
- Store sensitive/large objects in inventory.full.v1.json
- Maintain precise indexing for cross-references

### 2. Planner Module

**Purpose**: Generate execution plans through LLM with strict JSON schema validation.

```python
class DocumentPlanner:
    def generate_plan(self, structure: StructureV1, user_intent: str) -> PlanV1:
        """Generate plan through LLM with schema validation"""
        
    def validate_plan_schema(self, plan_json: dict) -> ValidationResult:
        """Validate against plan.v1 schema"""
        
    def check_whitelist_compliance(self, plan: PlanV1) -> ValidationResult:
        """Ensure only whitelisted operations"""
```

**Key Responsibilities**:
- Interface with LLM services (GPT/Claude)
- Enforce JSON schema validation
- Apply whitelist operation constraints
- Reject invalid plans with INVALID_PLAN status

### 3. Executor Module

**Purpose**: Execute atomic operations through Word COM with strict safety controls.

```python
class DocumentExecutor:
    def execute_plan(self, plan: PlanV1, docx_path: str) -> str:
        """Execute complete plan and return modified DOCX path"""
        
    def execute_operation(self, op: AtomicOperation, doc: WordDocument) -> OperationResult:
        """Execute single atomic operation"""
        
    def apply_localization_fallbacks(self, doc: WordDocument):
        """Apply style aliases and font fallbacks"""
```

**Atomic Operations Interface**:
```python
# Whitelisted operations only
def delete_section_by_heading(heading_text: str, level: int, match: str, case_sensitive: bool)
def update_toc() -> None
def delete_toc(mode: str) -> None
def set_style_rule(target_style: str, font_east_asian: str, font_latin: str, font_size_pt: int, font_bold: bool, line_spacing_mode: str, line_spacing_value: float)
def reassign_paragraphs_to_style(selector: dict, target_style: str, clear_direct_formatting: bool)
def clear_direct_formatting(scope: str, range_spec: dict) -> None
```

### 4. Validator Module

**Purpose**: Validate document modifications against strict assertions with rollback capability.

```python
class DocumentValidator:
    def validate_modifications(self, original_structure: StructureV1, modified_docx: str) -> ValidationResult:
        """Validate all assertions and generate comparison structure"""
        
    def check_chapter_assertions(self, structure: StructureV1) -> List[str]:
        """Verify no 摘要/参考文献 at level 1"""
        
    def check_style_assertions(self, structure: StructureV1) -> List[str]:
        """Verify H1/H2/Normal style specifications"""
        
    def check_toc_assertions(self, structure: StructureV1) -> List[str]:
        """Verify TOC consistency with heading tree"""
```

**Validation Assertions**:
- Chapter assertions: No "摘要/参考文献" at level 1
- Style assertions: H1 (楷体, 12pt, bold, 2.0 spacing), H2 (宋体, 12pt, bold, 2.0), Normal (宋体, 12pt, 2.0)
- TOC assertions: Items and page numbers match heading tree
- Pagination assertions: Fields updated, metadata.modified_time changed

### 5. Auditor Module

**Purpose**: Create complete audit trails with timestamped snapshots.

```python
class DocumentAuditor:
    def create_audit_directory(self) -> str:
        """Create timestamped run directory"""
        
    def save_snapshots(self, before_docx: str, after_docx: str, before_structure: StructureV1, after_structure: StructureV1, plan: PlanV1):
        """Save all required audit files"""
        
    def generate_diff_report(self, before_structure: StructureV1, after_structure: StructureV1) -> DiffReport:
        """Generate structural difference report"""
        
    def write_status(self, status: str, details: str):
        """Write final status (SUCCESS/ROLLBACK/FAILED_VALIDATION)"""
```

## Data Models

### Structure V1 Schema

```python
@dataclass
class StructureV1:
    schema_version: str = "structure.v1"
    metadata: DocumentMetadata
    styles: List[StyleDefinition]
    paragraphs: List[ParagraphSkeleton]  # Preview only, ≤120 chars
    headings: List[HeadingReference]
    fields: List[FieldReference]
    tables: List[TableSkeleton]  # Optional skeleton

@dataclass
class StyleDefinition:
    name: str
    type: str  # paragraph, character, table, linked
    font: FontSpec
    paragraph: ParagraphSpec

@dataclass
class FontSpec:
    east_asian: str
    latin: str
    size_pt: int
    bold: bool
    italic: bool
    color_hex: str

@dataclass
class ParagraphSpec:
    line_spacing_mode: str  # SINGLE, MULTIPLE, EXACTLY
    line_spacing_value: float
```

### Plan V1 Schema

```python
@dataclass
class PlanV1:
    schema_version: str = "plan.v1"
    ops: List[AtomicOperation]

@dataclass
class AtomicOperation:
    # Only whitelisted operation types allowed
    # Schema validation rejects unknown keys
    pass

# Example operations:
@dataclass
class DeleteSectionByHeading(AtomicOperation):
    heading_text: str
    level: int  # 1-9
    match: str  # EXACT, CONTAINS, REGEX
    case_sensitive: bool = False
    occurrence_index: Optional[int] = None

@dataclass
class SetStyleRule(AtomicOperation):
    target_style_name: str
    font: FontSpec
    paragraph: ParagraphSpec
```

### Inventory Full V1 Schema

```python
@dataclass
class InventoryFullV1:
    schema_version: str = "inventory.full.v1"
    ooxml_fragments: Dict[str, str]  # Object ID → OOXML
    media_indexes: Dict[str, MediaReference]
    content_controls: List[ContentControlReference]
    formulas: List[FormulaReference]
    charts: List[ChartReference]
```

## Error Handling

### Validation and Rollback Strategy

```python
class ValidationError(Exception):
    def __init__(self, assertion_failures: List[str], rollback_path: str):
        self.assertion_failures = assertion_failures
        self.rollback_path = rollback_path

class ExecutionPipeline:
    def process_document(self, docx_path: str, user_intent: str) -> ProcessingResult:
        try:
            # Extract
            structure, inventory = self.extractor.process_document(docx_path)
            
            # Plan
            plan = self.planner.generate_plan(structure, user_intent)
            
            # Execute
            modified_docx = self.executor.execute_plan(plan, docx_path)
            
            # Validate
            validation_result = self.validator.validate_modifications(structure, modified_docx)
            
            if not validation_result.is_valid:
                # Rollback
                self._rollback_to_original(docx_path)
                return ProcessingResult(status="FAILED_VALIDATION", errors=validation_result.errors)
            
            # Audit
            self.auditor.save_complete_audit_trail(...)
            
            return ProcessingResult(status="SUCCESS")
            
        except Exception as e:
            self._rollback_to_original(docx_path)
            return ProcessingResult(status="ROLLBACK", error=str(e))
```

### Localization and Fallback Handling

```python
class LocalizationManager:
    STYLE_ALIASES = {
        "Heading 1": "标题 1",
        "Heading 2": "标题 2", 
        "Normal": "正文"
    }
    
    FONT_FALLBACKS = {
        "楷体": ["楷体", "楷体_GB2312", "STKaiti"],
        "宋体": ["宋体", "SimSun"],
        "黑体": ["黑体", "SimHei"]
    }
    
    def resolve_style_name(self, style_name: str, doc: WordDocument) -> str:
        """Try English name first, then Chinese fallback"""
        
    def resolve_font_name(self, font_name: str, doc: WordDocument) -> str:
        """Apply font fallback chain, log to warnings.log"""
```

## Testing Strategy

### Unit Testing
- **Schema Validation**: Test all JSON schemas with valid/invalid inputs
- **Atomic Operations**: Mock Word COM for isolated operation testing
- **Localization**: Test style aliases and font fallbacks
- **Error Handling**: Test rollback scenarios and validation failures

### Integration Testing
- **LLM Integration**: Test plan generation with sample structures
- **Word COM Integration**: Test atomic operations with real documents
- **Pipeline Integration**: Test complete extract→plan→execute→validate→audit flow

### End-to-End Testing
- **Validation Test Cases**: All 7 DoD scenarios must pass
- **Complex Documents**: Test with formulas, content controls, charts
- **Error Scenarios**: Test rollback and recovery paths
- **Performance Testing**: Large documents and complex plans

### Test Data Strategy
```
tests/
├── fixtures/
│   ├── sample_documents/
│   │   ├── normal_paper.docx
│   │   ├── no_toc_document.docx
│   │   ├── duplicate_headings.docx
│   │   ├── headings_in_tables.docx
│   │   ├── missing_fonts.docx
│   │   ├── complex_objects.docx
│   │   └── revision_tracking.docx
│   ├── expected_structures/
│   ├── expected_plans/
│   └── expected_outputs/
├── unit/
├── integration/
└── e2e/
```

## Security and Runtime Constraints

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
- Status tracking (SUCCESS/ROLLBACK/FAILED_VALIDATION)

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
- LLM context limits (handled through chunking if needed)
- COM single-threading requirements

This design provides a robust, auditable foundation for AutoWord vNext that ensures predictable document processing with complete traceability and zero information loss.