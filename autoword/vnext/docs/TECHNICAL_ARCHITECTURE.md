# AutoWord vNext Technical Architecture

## Overview

AutoWord vNext implements a structured "Extract→Plan→Execute→Validate→Audit" closed loop system that transforms Word document processing into a predictable, reproducible pipeline. The system achieves >99% stability through strict JSON schemas, atomic operations, and comprehensive validation layers.

## Architectural Principles

### 1. Zero Information Loss
- Complete document inventory with OOXML fragment preservation
- Skeleton-only structure extraction (no full content)
- Comprehensive media and object reference tracking
- Reversible transformations with complete audit trails

### 2. Strict Schema Validation
- JSON Schema gateway for all data interchange
- Whitelist-only operation enforcement
- LLM output validation and rejection
- Runtime constraint checking

### 3. Atomic Operation Model
- Whitelisted operations through Word COM automation
- Object-layer modifications only (no string replacement)
- Transactional execution with rollback capability
- Comprehensive error handling and recovery

### 4. Complete Auditability
- Timestamped execution directories
- Before/after snapshots for all modifications
- Detailed operation logging and diff reports
- Status tracking (SUCCESS/ROLLBACK/FAILED_VALIDATION)

## System Architecture

### High-Level Pipeline Flow

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

### Module Architecture

#### 1. Extractor Module (`autoword.vnext.extractor`)

**Purpose**: Convert DOCX to structured JSON representations with zero information loss.

**Components**:
- `DocumentExtractor`: Main extraction orchestrator
- `StructureExtractor`: Document skeleton extraction
- `InventoryExtractor`: OOXML fragment and media extraction
- `MetadataExtractor`: Document metadata extraction

**Key Classes**:
```python
class DocumentExtractor:
    def extract_structure(self, docx_path: str) -> StructureV1
    def extract_inventory(self, docx_path: str) -> InventoryFullV1
    def process_document(self, docx_path: str) -> Tuple[StructureV1, InventoryFullV1]

class StructureExtractor:
    def extract_styles(self, doc: WordDocument) -> List[StyleDefinition]
    def extract_paragraphs(self, doc: WordDocument) -> List[ParagraphSkeleton]
    def extract_headings(self, doc: WordDocument) -> List[HeadingReference]
    def extract_fields(self, doc: WordDocument) -> List[FieldReference]
    def extract_tables(self, doc: WordDocument) -> List[TableSkeleton]
```

**Data Flow**:
```
DOCX File → Word COM → Document Object Model → JSON Extraction → 
structure.v1.json + inventory.full.v1.json
```

#### 2. Planner Module (`autoword.vnext.planner`)

**Purpose**: Generate execution plans through LLM with strict JSON schema validation.

**Components**:
- `DocumentPlanner`: Main planning orchestrator
- `LLMClient`: LLM service integration (GPT/Claude)
- `PlanValidator`: Schema and whitelist validation
- `PromptBuilder`: Structured prompt generation

**Key Classes**:
```python
class DocumentPlanner:
    def generate_plan(self, structure: StructureV1, user_intent: str) -> PlanV1
    def validate_plan_schema(self, plan_json: dict) -> ValidationResult
    def check_whitelist_compliance(self, plan: PlanV1) -> ValidationResult

class LLMClient:
    def generate_plan(self, prompt: str, model: str) -> str
    def validate_json_response(self, response: str) -> dict

class PlanValidator:
    def validate_schema(self, plan_data: dict) -> ValidationResult
    def check_operation_whitelist(self, operations: List[dict]) -> ValidationResult
```

**Security Model**:
- JSON Schema gateway validation
- Whitelist operation enforcement
- No DOCX/OOXML generation allowed
- Input sanitization and validation

#### 3. Executor Module (`autoword.vnext.executor`)

**Purpose**: Execute atomic operations through Word COM with strict safety controls.

**Components**:
- `DocumentExecutor`: Main execution orchestrator
- `AtomicOperations`: Whitelisted operation implementations
- `LocalizationManager`: Style aliases and font fallbacks
- `COMManager`: Word COM automation wrapper

**Atomic Operations**:
```python
class AtomicOperations:
    def delete_section_by_heading(self, heading_text: str, level: int, 
                                 match: str, case_sensitive: bool) -> OperationResult
    def update_toc(self) -> OperationResult
    def delete_toc(self, mode: str) -> OperationResult
    def set_style_rule(self, target_style: str, font_spec: FontSpec, 
                      paragraph_spec: ParagraphSpec) -> OperationResult
    def reassign_paragraphs_to_style(self, selector: dict, target_style: str, 
                                   clear_direct_formatting: bool) -> OperationResult
    def clear_direct_formatting(self, scope: str, range_spec: dict) -> OperationResult
```

**Safety Mechanisms**:
- Object-layer modifications only
- No string replacement on content/TOC text
- Proper COM resource management
- Exception handling with rollback

#### 4. Validator Module (`autoword.vnext.validator`)

**Purpose**: Validate document modifications against strict assertions with rollback capability.

**Components**:
- `DocumentValidator`: Main validation orchestrator
- `AssertionChecker`: Validation rule implementations
- `ComparisonEngine`: Before/after structure comparison
- `RollbackManager`: Automatic rollback on failure

**Validation Assertions**:
```python
class AssertionChecker:
    def check_chapter_assertions(self, structure: StructureV1) -> List[str]
    def check_style_assertions(self, structure: StructureV1) -> List[str]
    def check_toc_assertions(self, structure: StructureV1) -> List[str]
    def check_pagination_assertions(self, before: StructureV1, 
                                  after: StructureV1) -> List[str]
```

**Validation Rules**:
- Chapter assertions: No "摘要/参考文献" at level 1
- Style assertions: H1/H2/Normal style specifications
- TOC assertions: Consistency with heading tree
- Pagination assertions: Fields updated, metadata changed

#### 5. Auditor Module (`autoword.vnext.auditor`)

**Purpose**: Create complete audit trails with timestamped snapshots.

**Components**:
- `DocumentAuditor`: Main audit orchestrator
- `SnapshotManager`: Before/after file management
- `DiffGenerator`: Structural difference reporting
- `StatusTracker`: Execution status logging

**Audit Trail Structure**:
```
audit_runs/run_YYYYMMDD_HHMMSS_XXXXX/
├── before.docx                    # Original document
├── after.docx                     # Modified document
├── before_structure.v1.json       # Original structure
├── after_structure.v1.json        # Modified structure
├── plan.v1.json                   # Executed plan
├── diff.report.json               # Structural differences
├── warnings.log                   # Processing warnings
└── result.status.txt              # Final status
```

## Data Models and Schemas

### Core Data Models

#### StructureV1
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
```

#### PlanV1
```python
@dataclass
class PlanV1:
    schema_version: str = "plan.v1"
    ops: List[AtomicOperation]
```

#### InventoryFullV1
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

### Schema Validation System

**JSON Schema Files**:
- `schemas/structure.v1.json`: Structure validation schema
- `schemas/plan.v1.json`: Plan validation schema
- `schemas/inventory.full.v1.json`: Inventory validation schema

**Validation Pipeline**:
```python
class SchemaValidator:
    def validate_structure(self, data: dict) -> ValidationResult
    def validate_plan(self, data: dict) -> ValidationResult
    def validate_inventory(self, data: dict) -> ValidationResult
    def validate_against_schema(self, data: dict, schema_path: str) -> ValidationResult
```

## Error Handling and Recovery

### Exception Hierarchy

```python
class VNextException(Exception):
    """Base exception for vNext pipeline"""

class ExtractionError(VNextException):
    """Errors during document extraction"""

class PlanningError(VNextException):
    """Errors during plan generation"""

class ExecutionError(VNextException):
    """Errors during plan execution"""

class ValidationError(VNextException):
    """Errors during document validation"""

class AuditError(VNextException):
    """Errors during audit trail creation"""
```

### Recovery Mechanisms

#### Automatic Rollback
```python
class ExecutionPipeline:
    def process_document(self, docx_path: str, user_intent: str) -> ProcessingResult:
        try:
            # Pipeline execution
            result = self._execute_pipeline(docx_path, user_intent)
            return result
        except Exception as e:
            # Automatic rollback
            self._rollback_to_original(docx_path)
            return ProcessingResult(status="ROLLBACK", error=str(e))
```

#### Validation Failure Handling
```python
def validate_and_rollback(self, original_structure: StructureV1, 
                         modified_docx: str) -> ValidationResult:
    validation_result = self.validator.validate_modifications(
        original_structure, modified_docx)
    
    if not validation_result.is_valid:
        self._rollback_to_original(modified_docx)
        return ValidationResult(status="FAILED_VALIDATION", 
                              errors=validation_result.errors)
    
    return validation_result
```

## Performance Architecture

### Memory Management
- Stream processing for large documents
- Lazy loading of inventory objects
- Efficient JSON serialization
- Proper COM object disposal

### Optimization Strategies
- Document structure caching during session
- Batch similar atomic operations
- Minimize Word COM round-trips
- Efficient schema validation

### Performance Monitoring
```python
class PerformanceMonitor:
    def track_operation_time(self, operation: str, duration: float)
    def track_memory_usage(self, stage: str, memory_mb: float)
    def track_document_size(self, docx_path: str, size_mb: float)
    def generate_performance_report(self) -> PerformanceReport
```

## Security Architecture

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
- Status tracking with integrity checks

## Deployment Architecture

### Package Structure
```
autoword-vnext/
├── autoword/vnext/              # Core pipeline modules
├── config/                      # Configuration files
├── schemas/                     # JSON schema definitions
├── test_data/                   # DoD validation scenarios
├── docs/                        # Documentation
├── run_vnext.py                 # Cross-platform entry point
├── DEPLOYMENT_MANIFEST.json     # Package manifest
└── requirements-deployment.txt  # Production dependencies
```

### Configuration Management
- Hierarchical configuration system
- Environment-specific overrides
- Runtime parameter validation
- Configuration file validation

### Monitoring and Logging
- Structured logging with multiple levels
- Performance metrics collection
- Error tracking and reporting
- Audit trail management

This technical architecture provides a robust, scalable foundation for AutoWord vNext that ensures predictable document processing with complete traceability and zero information loss.