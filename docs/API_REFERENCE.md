# AutoWord vNext - API Reference

## Overview

This document provides complete API reference for all public interfaces and data models in AutoWord vNext.

## Core Pipeline API

### VNextPipeline

Main pipeline orchestrator for document processing.

```python
class VNextPipeline:
    def __init__(self, config: Optional[VNextConfig] = None):
        """Initialize pipeline with optional configuration"""
        
    def process_document(self, docx_path: str, user_intent: str) -> ProcessingResult:
        """Process document with user intent"""
        
    def generate_plan(self, docx_path: str, user_intent: str) -> PlanResult:
        """Generate execution plan without execution (dry run)"""
        
    def validate_document(self, docx_path: str) -> ValidationResult:
        """Validate document structure and integrity"""
```

#### Methods

##### process_document

```python
def process_document(self, docx_path: str, user_intent: str) -> ProcessingResult:
    """
    Process document through complete pipeline.
    
    Args:
        docx_path: Path to input DOCX file
        user_intent: Natural language description of desired changes
        
    Returns:
        ProcessingResult with status, output path, and details
        
    Raises:
        VNextError: Pipeline processing error
        FileNotFoundError: Input file not found
        PermissionError: File access denied
    """
```

**Example**:
```python
pipeline = VNextPipeline()
result = pipeline.process_document(
    "document.docx",
    "Delete abstract section and update table of contents"
)

if result.status == "SUCCESS":
    print(f"Output: {result.output_path}")
    print(f"Audit: {result.audit_directory}")
else:
    print(f"Error: {result.error}")
```

##### generate_plan

```python
def generate_plan(self, docx_path: str, user_intent: str) -> PlanResult:
    """
    Generate execution plan without executing (dry run).
    
    Args:
        docx_path: Path to input DOCX file
        user_intent: Natural language description of desired changes
        
    Returns:
        PlanResult with generated plan and validation status
        
    Raises:
        ExtractionError: Document structure extraction failed
        PlanningError: Plan generation failed
    """
```

**Example**:
```python
result = pipeline.generate_plan("document.docx", "Remove all appendices")

if result.is_valid:
    print(f"Operations: {len(result.plan.ops)}")
    for op in result.plan.ops:
        print(f"- {op.op_type}: {op}")
else:
    print(f"Invalid plan: {result.errors}")
```

## Module APIs

### Extractor Module

#### DocumentExtractor

```python
class DocumentExtractor:
    def __init__(self, config: Optional[ExtractorConfig] = None):
        """Initialize extractor with optional configuration"""
        
    def extract_structure(self, docx_path: str) -> StructureV1:
        """Extract document structure to JSON schema"""
        
    def extract_inventory(self, docx_path: str) -> InventoryFullV1:
        """Extract complete document inventory"""
        
    def process_document(self, docx_path: str) -> Tuple[StructureV1, InventoryFullV1]:
        """Extract both structure and inventory"""
```

**Example**:
```python
extractor = DocumentExtractor()
structure, inventory = extractor.process_document("document.docx")

print(f"Paragraphs: {len(structure.paragraphs)}")
print(f"Headings: {len(structure.headings)}")
print(f"Styles: {len(structure.styles)}")
```

### Planner Module

#### DocumentPlanner

```python
class DocumentPlanner:
    def __init__(self, llm_config: LLMConfig):
        """Initialize planner with LLM configuration"""
        
    def generate_plan(self, structure: StructureV1, user_intent: str) -> PlanV1:
        """Generate execution plan through LLM"""
        
    def validate_plan(self, plan: PlanV1) -> ValidationResult:
        """Validate plan against schema and whitelist"""
```

**Example**:
```python
from autoword.vnext.planner import DocumentPlanner, LLMConfig

llm_config = LLMConfig(
    provider="openai",
    model="gpt-4",
    api_key="sk-..."
)

planner = DocumentPlanner(llm_config)
plan = planner.generate_plan(structure, "Delete abstract section")
```

### Executor Module

#### DocumentExecutor

```python
class DocumentExecutor:
    def __init__(self, config: Optional[ExecutorConfig] = None):
        """Initialize executor with optional configuration"""
        
    def execute_plan(self, plan: PlanV1, docx_path: str) -> str:
        """Execute complete plan and return output path"""
        
    def execute_operation(self, operation: AtomicOperation, doc: WordDocument) -> OperationResult:
        """Execute single atomic operation"""
```

**Example**:
```python
executor = DocumentExecutor()
output_path = executor.execute_plan(plan, "input.docx")
print(f"Modified document: {output_path}")
```

### Validator Module

#### DocumentValidator

```python
class DocumentValidator:
    def __init__(self, config: Optional[ValidatorConfig] = None):
        """Initialize validator with optional configuration"""
        
    def validate_modifications(self, original_structure: StructureV1, modified_docx: str) -> ValidationResult:
        """Validate document modifications"""
        
    def check_assertions(self, structure: StructureV1) -> List[AssertionResult]:
        """Check all validation assertions"""
```

**Example**:
```python
validator = DocumentValidator()
result = validator.validate_modifications(original_structure, "modified.docx")

if not result.is_valid:
    print("Validation failures:")
    for error in result.errors:
        print(f"- {error}")
```

### Auditor Module

#### DocumentAuditor

```python
class DocumentAuditor:
    def __init__(self, config: Optional[AuditorConfig] = None):
        """Initialize auditor with optional configuration"""
        
    def create_audit_directory(self) -> str:
        """Create timestamped audit directory"""
        
    def save_complete_audit(self, before_docx: str, after_docx: str, 
                          before_structure: StructureV1, after_structure: StructureV1,
                          plan: PlanV1, status: str) -> str:
        """Save complete audit trail"""
```

## Data Models

### Core Data Types

#### ProcessingResult

```python
@dataclass
class ProcessingResult:
    status: str  # SUCCESS, ROLLBACK, FAILED_VALIDATION, INVALID_PLAN
    output_path: Optional[str] = None
    audit_directory: Optional[str] = None
    error: Optional[str] = None
    validation_errors: Optional[List[str]] = None
    plan_errors: Optional[List[str]] = None
    processing_time: Optional[float] = None
```

#### PlanResult

```python
@dataclass
class PlanResult:
    plan: Optional[PlanV1] = None
    is_valid: bool = False
    errors: Optional[List[str]] = None
    warnings: Optional[List[str]] = None
```

#### ValidationResult

```python
@dataclass
class ValidationResult:
    is_valid: bool
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    assertion_results: List[AssertionResult] = field(default_factory=list)
```

### Schema Data Models

#### StructureV1

```python
@dataclass
class StructureV1:
    schema_version: str = "structure.v1"
    metadata: DocumentMetadata
    styles: List[StyleDefinition]
    paragraphs: List[ParagraphSkeleton]
    headings: List[HeadingReference]
    fields: List[FieldReference]
    tables: List[TableSkeleton]
```

#### DocumentMetadata

```python
@dataclass
class DocumentMetadata:
    title: str
    author: str
    created_time: str  # ISO8601
    modified_time: str  # ISO8601
    word_version: str
    page_count: int
    word_count: int
    character_count: int
```

#### StyleDefinition

```python
@dataclass
class StyleDefinition:
    name: str
    type: str  # paragraph, character, table, linked
    font: FontSpec
    paragraph: ParagraphSpec
    is_builtin: bool
    is_modified: bool
```

#### FontSpec

```python
@dataclass
class FontSpec:
    east_asian: str
    latin: str
    size_pt: int
    bold: bool
    italic: bool
    underline: bool
    color_hex: str
```

#### ParagraphSpec

```python
@dataclass
class ParagraphSpec:
    line_spacing_mode: str  # SINGLE, MULTIPLE, EXACTLY
    line_spacing_value: float
    space_before_pt: int
    space_after_pt: int
    alignment: str  # LEFT, CENTER, RIGHT, JUSTIFY
    indent_left_pt: int
    indent_right_pt: int
    indent_first_line_pt: int
```

#### ParagraphSkeleton

```python
@dataclass
class ParagraphSkeleton:
    index: int
    style_name: str
    preview_text: str  # ≤120 characters
    is_heading: bool
    heading_level: Optional[int] = None
    has_direct_formatting: bool = False
```

#### HeadingReference

```python
@dataclass
class HeadingReference:
    text: str
    level: int  # 1-9
    paragraph_index: int
    occurrence_index: int  # For duplicate headings
    style_name: str
```

#### FieldReference

```python
@dataclass
class FieldReference:
    type: str  # TOC, PAGE, DATE, etc.
    paragraph_index: int
    field_code: str
    result_text: str
    is_locked: bool
```

#### TableSkeleton

```python
@dataclass
class TableSkeleton:
    paragraph_index: int
    rows: int
    columns: int
    cell_references: List[int]  # Paragraph indexes of cells
    has_header_row: bool
    style_name: Optional[str] = None
```

#### PlanV1

```python
@dataclass
class PlanV1:
    schema_version: str = "plan.v1"
    ops: List[AtomicOperation]
    metadata: PlanMetadata
```

#### PlanMetadata

```python
@dataclass
class PlanMetadata:
    generated_time: str  # ISO8601
    user_intent: str
    llm_model: str
    operation_count: int
```

### Atomic Operation Types

#### DeleteSectionByHeading

```python
@dataclass
class DeleteSectionByHeading:
    op_type: str = "delete_section_by_heading"
    heading_text: str
    level: int  # 1-9
    match: str  # EXACT, CONTAINS, REGEX
    case_sensitive: bool = False
    occurrence_index: Optional[int] = None
```

#### UpdateTOC

```python
@dataclass
class UpdateTOC:
    op_type: str = "update_toc"
```

#### DeleteTOC

```python
@dataclass
class DeleteTOC:
    op_type: str = "delete_toc"
    mode: str = "first"  # first, all, by_index
    toc_index: Optional[int] = None
```

#### SetStyleRule

```python
@dataclass
class SetStyleRule:
    op_type: str = "set_style_rule"
    target_style_name: str
    font: FontSpec
    paragraph: ParagraphSpec
```

#### ReassignParagraphsToStyle

```python
@dataclass
class ReassignParagraphsToStyle:
    op_type: str = "reassign_paragraphs_to_style"
    selector: ParagraphSelector
    target_style: str
    clear_direct_formatting: bool = False
```

#### ParagraphSelector

```python
@dataclass
class ParagraphSelector:
    current_style: Optional[str] = None
    text_contains: Optional[str] = None
    paragraph_indexes: Optional[List[int]] = None
    heading_level: Optional[int] = None
```

#### ClearDirectFormatting

```python
@dataclass
class ClearDirectFormatting:
    op_type: str = "clear_direct_formatting"
    scope: str  # document, selection, style
    range_spec: Optional[RangeSpec] = None
    authorization: str = "EXPLICIT_USER_REQUEST"
```

#### RangeSpec

```python
@dataclass
class RangeSpec:
    start_paragraph: Optional[int] = None
    end_paragraph: Optional[int] = None
    style_name: Optional[str] = None
```

## Configuration Classes

### VNextConfig

```python
@dataclass
class VNextConfig:
    llm: LLMConfig
    localization: LocalizationConfig
    validation: ValidationConfig
    audit: AuditConfig
    executor: ExecutorConfig
```

### LLMConfig

```python
@dataclass
class LLMConfig:
    provider: str  # openai, anthropic, azure
    model: str
    api_key: str
    base_url: Optional[str] = None
    temperature: float = 0.1
    max_tokens: int = 4000
    timeout: int = 30
    retry_attempts: int = 3
```

### LocalizationConfig

```python
@dataclass
class LocalizationConfig:
    language: str = "zh-CN"
    style_aliases: Dict[str, str] = field(default_factory=dict)
    font_fallbacks: Dict[str, List[str]] = field(default_factory=dict)
    enable_font_fallback: bool = True
    log_fallbacks: bool = True
```

### ValidationConfig

```python
@dataclass
class ValidationConfig:
    strict_mode: bool = True
    rollback_on_failure: bool = True
    chapter_assertions: bool = True
    style_assertions: bool = True
    toc_assertions: bool = True
    pagination_assertions: bool = True
    forbidden_level1_headings: List[str] = field(default_factory=lambda: ["摘要", "参考文献"])
```

### AuditConfig

```python
@dataclass
class AuditConfig:
    save_snapshots: bool = True
    generate_diff_reports: bool = True
    output_directory: str = "./audit_output"
    keep_audit_days: int = 30
    compress_old_audits: bool = True
```

### ExecutorConfig

```python
@dataclass
class ExecutorConfig:
    batch_operations: bool = True
    com_timeout: int = 30
    retry_com_errors: bool = True
    max_com_retries: int = 3
    enable_localization: bool = True
```

## Exception Classes

### Base Exceptions

```python
class VNextError(Exception):
    """Base exception for vNext pipeline"""
    pass

class ExtractionError(VNextError):
    """Errors during document extraction"""
    pass

class PlanningError(VNextError):
    """Errors during plan generation"""
    pass

class ExecutionError(VNextError):
    """Errors during plan execution"""
    pass

class ValidationError(VNextError):
    """Errors during validation"""
    def __init__(self, assertion_failures: List[str], rollback_path: str):
        self.assertion_failures = assertion_failures
        self.rollback_path = rollback_path
        super().__init__(f"Validation failed: {assertion_failures}")

class AuditError(VNextError):
    """Errors during audit trail generation"""
    pass
```

### Specific Exceptions

```python
class DocumentLockError(ExtractionError):
    """Document is locked or in use"""
    pass

class StructureParsingError(ExtractionError):
    """Failed to parse document structure"""
    pass

class LLMConnectionError(PlanningError):
    """Failed to connect to LLM service"""
    pass

class InvalidPlanError(PlanningError):
    """Generated plan failed validation"""
    pass

class COMAutomationError(ExecutionError):
    """Word COM automation failed"""
    pass

class OperationNotFoundError(ExecutionError):
    """Target for operation not found"""
    pass

class AssertionFailureError(ValidationError):
    """Document assertion failed"""
    pass

class RollbackError(ValidationError):
    """Failed to rollback document"""
    pass
```

## Utility Functions

### Schema Validation

```python
def validate_structure_schema(structure_json: dict) -> ValidationResult:
    """Validate structure JSON against schema"""
    
def validate_plan_schema(plan_json: dict) -> ValidationResult:
    """Validate plan JSON against schema"""
    
def validate_inventory_schema(inventory_json: dict) -> ValidationResult:
    """Validate inventory JSON against schema"""
```

### File Operations

```python
def create_timestamped_directory(base_path: str) -> str:
    """Create directory with timestamp"""
    
def backup_document(docx_path: str, backup_dir: str) -> str:
    """Create document backup"""
    
def cleanup_temp_files(temp_dir: str) -> None:
    """Clean up temporary files"""
```

### Logging Utilities

```python
def get_logger(name: str) -> logging.Logger:
    """Get configured logger instance"""
    
def setup_logging(level: str = "INFO", log_file: Optional[str] = None) -> None:
    """Setup logging configuration"""
```

## Usage Examples

### Basic Pipeline Usage

```python
from autoword.vnext import VNextPipeline

# Initialize with default configuration
pipeline = VNextPipeline()

# Process document
result = pipeline.process_document(
    "document.docx",
    "Delete abstract and references sections, update TOC"
)

# Check result
if result.status == "SUCCESS":
    print(f"Success: {result.output_path}")
    print(f"Audit: {result.audit_directory}")
else:
    print(f"Failed: {result.error}")
```

### Custom Configuration

```python
from autoword.vnext import VNextPipeline, VNextConfig, LLMConfig

# Custom configuration
config = VNextConfig(
    llm=LLMConfig(
        provider="openai",
        model="gpt-4",
        api_key="sk-...",
        temperature=0.1
    ),
    localization=LocalizationConfig(
        language="zh-CN",
        style_aliases={
            "Heading 1": "标题 1",
            "Normal": "正文"
        }
    )
)

pipeline = VNextPipeline(config)
```

### Error Handling

```python
from autoword.vnext import VNextPipeline, VNextError

pipeline = VNextPipeline()

try:
    result = pipeline.process_document("document.docx", user_intent)
    
    if result.status == "SUCCESS":
        print("Processing successful")
    elif result.status == "FAILED_VALIDATION":
        print(f"Validation failed: {result.validation_errors}")
    elif result.status == "INVALID_PLAN":
        print(f"Invalid plan: {result.plan_errors}")
        
except VNextError as e:
    print(f"Pipeline error: {e}")
except Exception as e:
    print(f"Unexpected error: {e}")
```

### Batch Processing

```python
import glob
from autoword.vnext import VNextPipeline

pipeline = VNextPipeline()
documents = glob.glob("*.docx")

for doc_path in documents:
    try:
        result = pipeline.process_document(
            doc_path,
            "Standardize formatting and update TOC"
        )
        print(f"{doc_path}: {result.status}")
    except Exception as e:
        print(f"{doc_path}: ERROR - {e}")
```

This API reference provides complete documentation for all public interfaces in AutoWord vNext. For implementation details and internal APIs, refer to the source code and technical architecture documentation.