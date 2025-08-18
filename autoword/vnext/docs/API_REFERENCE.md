# AutoWord vNext API Reference

## Overview

This document provides comprehensive API documentation for all public interfaces and data models in AutoWord vNext. The API is designed for programmatic integration and advanced usage scenarios.

## Core Pipeline API

### VNextPipeline

The main pipeline orchestrator that coordinates all five modules.

```python
from autoword.vnext.pipeline import VNextPipeline

class VNextPipeline:
    def __init__(self, config: Optional[Dict] = None)
    def process_document(self, docx_path: str, user_intent: str, 
                        audit_dir: Optional[str] = None) -> ProcessingResult
    def process_batch(self, input_dir: str, user_intent: str,
                     audit_dir: Optional[str] = None) -> BatchProcessingResult
    def dry_run(self, docx_path: str, user_intent: str) -> DryRunResult
```

#### Methods

##### `__init__(config: Optional[Dict] = None)`
Initialize the pipeline with optional configuration.

**Parameters**:
- `config` (dict, optional): Configuration dictionary overriding defaults

**Example**:
```python
config = {
    "model": "gpt4",
    "temperature": 0.1,
    "enable_audit_trail": True
}
pipeline = VNextPipeline(config)
```

##### `process_document(docx_path: str, user_intent: str, audit_dir: Optional[str] = None) -> ProcessingResult`
Process a single document through the complete pipeline.

**Parameters**:
- `docx_path` (str): Path to input DOCX file
- `user_intent` (str): Natural language description of desired changes
- `audit_dir` (str, optional): Base directory for audit trails

**Returns**: `ProcessingResult` object

**Example**:
```python
result = pipeline.process_document(
    "document.docx",
    "Remove abstract and references sections",
    audit_dir="./my_audits"
)

if result.status == "SUCCESS":
    print(f"Success: {result.output_path}")
    print(f"Audit: {result.audit_directory}")
else:
    print(f"Failed: {result.error}")
```

##### `process_batch(input_dir: str, user_intent: str, audit_dir: Optional[str] = None) -> BatchProcessingResult`
Process multiple documents in a directory.

**Parameters**:
- `input_dir` (str): Directory containing DOCX files
- `user_intent` (str): Intent to apply to all documents
- `audit_dir` (str, optional): Base directory for audit trails

**Returns**: `BatchProcessingResult` object

**Example**:
```python
batch_result = pipeline.process_batch(
    "./documents",
    "Apply standard formatting",
    audit_dir="./batch_audits"
)

print(f"Processed: {batch_result.total_processed}")
print(f"Successful: {batch_result.successful}")
print(f"Failed: {batch_result.failed}")
```

##### `dry_run(docx_path: str, user_intent: str) -> DryRunResult`
Generate execution plan without modifying the document.

**Parameters**:
- `docx_path` (str): Path to input DOCX file
- `user_intent` (str): Natural language description of desired changes

**Returns**: `DryRunResult` object

**Example**:
```python
dry_result = pipeline.dry_run("document.docx", "Update TOC")

if dry_result.plan_generated:
    print("Generated plan:")
    for op in dry_result.plan.ops:
        print(f"  - {op.operation}")
else:
    print(f"Plan generation failed: {dry_result.error}")
```

## Module APIs

### Extractor Module

Extract document structure and inventory from DOCX files.

```python
from autoword.vnext.extractor import DocumentExtractor

class DocumentExtractor:
    def __init__(self, config: Optional[Dict] = None)
    def extract_structure(self, docx_path: str) -> StructureV1
    def extract_inventory(self, docx_path: str) -> InventoryFullV1
    def process_document(self, docx_path: str) -> Tuple[StructureV1, InventoryFullV1]
```

#### Methods

##### `extract_structure(docx_path: str) -> StructureV1`
Extract document structure as JSON schema.

**Parameters**:
- `docx_path` (str): Path to DOCX file

**Returns**: `StructureV1` data model

**Example**:
```python
extractor = DocumentExtractor()
structure = extractor.extract_structure("document.docx")

print(f"Document has {len(structure.headings)} headings")
print(f"Styles: {[s.name for s in structure.styles]}")
```

##### `extract_inventory(docx_path: str) -> InventoryFullV1`
Extract complete document inventory including OOXML fragments.

**Parameters**:
- `docx_path` (str): Path to DOCX file

**Returns**: `InventoryFullV1` data model

**Example**:
```python
inventory = extractor.extract_inventory("document.docx")
print(f"OOXML fragments: {len(inventory.ooxml_fragments)}")
print(f"Media files: {len(inventory.media_indexes)}")
```

### Planner Module

Generate execution plans through LLM integration.

```python
from autoword.vnext.planner import DocumentPlanner

class DocumentPlanner:
    def __init__(self, config: Optional[Dict] = None)
    def generate_plan(self, structure: StructureV1, user_intent: str) -> PlanV1
    def validate_plan(self, plan_data: dict) -> ValidationResult
```

#### Methods

##### `generate_plan(structure: StructureV1, user_intent: str) -> PlanV1`
Generate execution plan using LLM.

**Parameters**:
- `structure` (StructureV1): Document structure from extractor
- `user_intent` (str): Natural language intent

**Returns**: `PlanV1` data model

**Raises**: `PlanningError` if plan generation fails

**Example**:
```python
planner = DocumentPlanner({"model": "gpt4"})
plan = planner.generate_plan(structure, "Remove abstract section")

print(f"Generated {len(plan.ops)} operations")
for op in plan.ops:
    print(f"  - {op.operation}")
```

### Executor Module

Execute atomic operations through Word COM automation.

```python
from autoword.vnext.executor import DocumentExecutor

class DocumentExecutor:
    def __init__(self, config: Optional[Dict] = None)
    def execute_plan(self, plan: PlanV1, docx_path: str) -> str
    def execute_operation(self, operation: AtomicOperation, doc: WordDocument) -> OperationResult
```

#### Methods

##### `execute_plan(plan: PlanV1, docx_path: str) -> str`
Execute complete plan on document.

**Parameters**:
- `plan` (PlanV1): Execution plan from planner
- `docx_path` (str): Path to DOCX file to modify

**Returns**: Path to modified document

**Raises**: `ExecutionError` if execution fails

**Example**:
```python
executor = DocumentExecutor()
modified_path = executor.execute_plan(plan, "document.docx")
print(f"Modified document saved to: {modified_path}")
```

### Validator Module

Validate document modifications against assertions.

```python
from autoword.vnext.validator import DocumentValidator

class DocumentValidator:
    def __init__(self, config: Optional[Dict] = None)
    def validate_modifications(self, original_structure: StructureV1, 
                             modified_docx: str) -> ValidationResult
    def check_assertions(self, structure: StructureV1) -> List[str]
```

#### Methods

##### `validate_modifications(original_structure: StructureV1, modified_docx: str) -> ValidationResult`
Validate document modifications.

**Parameters**:
- `original_structure` (StructureV1): Original document structure
- `modified_docx` (str): Path to modified document

**Returns**: `ValidationResult` object

**Example**:
```python
validator = DocumentValidator()
result = validator.validate_modifications(original_structure, "modified.docx")

if result.is_valid:
    print("Validation passed")
else:
    print(f"Validation failed: {result.errors}")
```

### Auditor Module

Create audit trails and manage snapshots.

```python
from autoword.vnext.auditor import DocumentAuditor

class DocumentAuditor:
    def __init__(self, config: Optional[Dict] = None)
    def create_audit_directory(self, base_dir: str) -> str
    def save_audit_trail(self, audit_dir: str, before_docx: str, after_docx: str,
                        before_structure: StructureV1, after_structure: StructureV1,
                        plan: PlanV1, status: str) -> None
```

#### Methods

##### `create_audit_directory(base_dir: str) -> str`
Create timestamped audit directory.

**Parameters**:
- `base_dir` (str): Base directory for audit trails

**Returns**: Path to created audit directory

**Example**:
```python
auditor = DocumentAuditor()
audit_dir = auditor.create_audit_directory("./audits")
print(f"Created audit directory: {audit_dir}")
```

## Data Models

### Core Data Models

#### ProcessingResult

Result of document processing operation.

```python
@dataclass
class ProcessingResult:
    status: str  # SUCCESS, FAILED_VALIDATION, ROLLBACK, INVALID_PLAN
    output_path: Optional[str] = None
    audit_directory: Optional[str] = None
    error: Optional[str] = None
    execution_time: Optional[float] = None
    operations_executed: Optional[int] = None
```

#### BatchProcessingResult

Result of batch processing operation.

```python
@dataclass
class BatchProcessingResult:
    total_processed: int
    successful: int
    failed: int
    results: List[ProcessingResult]
    execution_time: float
```

#### DryRunResult

Result of dry-run operation.

```python
@dataclass
class DryRunResult:
    plan_generated: bool
    plan: Optional[PlanV1] = None
    error: Optional[str] = None
    structure: Optional[StructureV1] = None
```

#### ValidationResult

Result of validation operation.

```python
@dataclass
class ValidationResult:
    is_valid: bool
    errors: List[str]
    warnings: List[str]
    assertions_checked: int
    assertions_passed: int
```

### Document Structure Models

#### StructureV1

Complete document structure representation.

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

Document metadata information.

```python
@dataclass
class DocumentMetadata:
    title: str
    author: str
    created_time: str
    modified_time: str
    word_version: str
    page_count: int
    paragraph_count: int
    word_count: int
```

#### StyleDefinition

Style definition with formatting specifications.

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

Font formatting specification.

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

Paragraph formatting specification.

```python
@dataclass
class ParagraphSpec:
    line_spacing_mode: str  # SINGLE, MULTIPLE, EXACTLY
    line_spacing_value: float
    space_before_pt: float
    space_after_pt: float
    alignment: str  # LEFT, CENTER, RIGHT, JUSTIFY
    indent_left_pt: float
    indent_right_pt: float
    indent_first_line_pt: float
```

#### HeadingReference

Heading reference with location information.

```python
@dataclass
class HeadingReference:
    text: str
    level: int  # 1-9
    style_name: str
    paragraph_index: int
    page_number: int
    in_table: bool
    table_index: Optional[int] = None
```

#### ParagraphSkeleton

Paragraph skeleton with preview text.

```python
@dataclass
class ParagraphSkeleton:
    index: int
    style_name: str
    preview_text: str  # ≤120 characters
    is_heading: bool
    heading_level: Optional[int] = None
    page_number: int
```

### Execution Plan Models

#### PlanV1

Execution plan with atomic operations.

```python
@dataclass
class PlanV1:
    schema_version: str = "plan.v1"
    ops: List[AtomicOperation]
```

#### AtomicOperation

Base class for atomic operations.

```python
@dataclass
class AtomicOperation:
    operation: str
    # Specific parameters defined in subclasses
```

#### DeleteSectionByHeading

Delete section operation.

```python
@dataclass
class DeleteSectionByHeading(AtomicOperation):
    operation: str = "delete_section_by_heading"
    heading_text: str
    level: int
    match: str  # EXACT, CONTAINS, REGEX
    case_sensitive: bool = False
    occurrence_index: Optional[int] = None
```

#### SetStyleRule

Style modification operation.

```python
@dataclass
class SetStyleRule(AtomicOperation):
    operation: str = "set_style_rule"
    target_style: str
    font_east_asian: Optional[str] = None
    font_latin: Optional[str] = None
    font_size_pt: Optional[int] = None
    font_bold: Optional[bool] = None
    line_spacing_mode: Optional[str] = None
    line_spacing_value: Optional[float] = None
```

### Inventory Models

#### InventoryFullV1

Complete document inventory.

```python
@dataclass
class InventoryFullV1:
    schema_version: str = "inventory.full.v1"
    ooxml_fragments: Dict[str, str]
    media_indexes: Dict[str, MediaReference]
    content_controls: List[ContentControlReference]
    formulas: List[FormulaReference]
    charts: List[ChartReference]
```

#### MediaReference

Media file reference.

```python
@dataclass
class MediaReference:
    media_id: str
    filename: str
    content_type: str
    size_bytes: int
    embedded: bool
```

## Configuration API

### Configuration Management

```python
from autoword.vnext.config import ConfigManager

class ConfigManager:
    @staticmethod
    def load_config(config_path: str) -> Dict
    @staticmethod
    def validate_config(config: Dict) -> ValidationResult
    @staticmethod
    def merge_configs(base: Dict, override: Dict) -> Dict
    @staticmethod
    def create_template(output_path: str) -> None
```

#### Methods

##### `load_config(config_path: str) -> Dict`
Load configuration from JSON file.

**Parameters**:
- `config_path` (str): Path to configuration file

**Returns**: Configuration dictionary

**Example**:
```python
config = ConfigManager.load_config("my-config.json")
pipeline = VNextPipeline(config)
```

##### `create_template(output_path: str) -> None`
Create configuration template file.

**Parameters**:
- `output_path` (str): Path for template file

**Example**:
```python
ConfigManager.create_template("template-config.json")
```

### Configuration Schema

```python
{
    "model": "gpt4",  # LLM model: gpt4, gpt35, claude37, claude3
    "temperature": 0.1,  # LLM temperature (0.0-1.0)
    "audit_dir": "./audit_trails",  # Base audit directory
    "visible": false,  # Show Word windows during processing
    "enable_audit_trail": true,  # Create audit trails
    "enable_rollback": true,  # Enable automatic rollback
    "max_execution_time_seconds": 300,  # Maximum execution time
    "monitoring_level": "detailed",  # basic, detailed, debug, performance
    "localization": {
        "style_aliases": {
            "Heading 1": "标题 1"
        },
        "font_fallbacks": {
            "楷体": ["楷体", "楷体_GB2312", "STKaiti"]
        }
    },
    "validation_rules": {
        "chapter_assertions": {
            "forbidden_level_1_headings": ["摘要", "参考文献"]
        },
        "style_assertions": {
            "H1": {
                "font_east_asian": "楷体",
                "font_size_pt": 12,
                "font_bold": true
            }
        }
    }
}
```

## Exception Handling

### Exception Hierarchy

```python
class VNextException(Exception):
    """Base exception for vNext pipeline"""
    pass

class ExtractionError(VNextException):
    """Errors during document extraction"""
    pass

class PlanningError(VNextException):
    """Errors during plan generation"""
    pass

class ExecutionError(VNextException):
    """Errors during plan execution"""
    pass

class ValidationError(VNextException):
    """Errors during document validation"""
    pass

class AuditError(VNextException):
    """Errors during audit trail creation"""
    pass

class ConfigurationError(VNextException):
    """Configuration-related errors"""
    pass
```

### Error Handling Patterns

```python
from autoword.vnext.exceptions import VNextException, PlanningError

try:
    result = pipeline.process_document("doc.docx", "Remove sections")
    if result.status == "SUCCESS":
        print(f"Success: {result.output_path}")
    else:
        print(f"Processing failed: {result.status}")
        if result.error:
            print(f"Error details: {result.error}")
            
except PlanningError as e:
    print(f"Plan generation failed: {e}")
    
except VNextException as e:
    print(f"Pipeline error: {e}")
    
except Exception as e:
    print(f"Unexpected error: {e}")
```

## Utility Functions

### Schema Validation

```python
from autoword.vnext.schema_validator import SchemaValidator

validator = SchemaValidator()

# Validate structure
result = validator.validate_structure(structure_data)

# Validate plan
result = validator.validate_plan(plan_data)

# Validate inventory
result = validator.validate_inventory(inventory_data)
```

### Localization Support

```python
from autoword.vnext.localization import LocalizationManager

localization = LocalizationManager()

# Resolve style name
style_name = localization.resolve_style_name("Heading 1", document)

# Resolve font name with fallback
font_name = localization.resolve_font_name("楷体", document)
```

### Performance Monitoring

```python
from autoword.vnext.monitoring import PerformanceMonitor

monitor = PerformanceMonitor()

# Track operation
with monitor.track_operation("extraction"):
    structure = extractor.extract_structure("doc.docx")

# Get performance report
report = monitor.generate_report()
print(f"Total time: {report.total_time}")
print(f"Memory peak: {report.peak_memory_mb}")
```

## Integration Examples

### Basic Integration

```python
from autoword.vnext.pipeline import VNextPipeline

# Initialize pipeline
pipeline = VNextPipeline({
    "model": "gpt4",
    "audit_dir": "./my_audits"
})

# Process document
result = pipeline.process_document(
    "input.docx",
    "Remove abstract and references, update TOC"
)

# Handle result
if result.status == "SUCCESS":
    print(f"Document processed: {result.output_path}")
    print(f"Audit trail: {result.audit_directory}")
else:
    print(f"Processing failed: {result.status}")
    if result.error:
        print(f"Error: {result.error}")
```

### Advanced Integration

```python
from autoword.vnext.pipeline import VNextPipeline
from autoword.vnext.extractor import DocumentExtractor
from autoword.vnext.planner import DocumentPlanner

# Custom configuration
config = {
    "model": "gpt4",
    "temperature": 0.05,
    "monitoring_level": "performance",
    "localization": {
        "font_fallbacks": {
            "CustomFont": ["CustomFont", "Fallback1", "Fallback2"]
        }
    }
}

# Initialize components
pipeline = VNextPipeline(config)
extractor = DocumentExtractor(config)
planner = DocumentPlanner(config)

# Extract structure first
structure = extractor.extract_structure("document.docx")
print(f"Document has {len(structure.headings)} headings")

# Generate plan
plan = planner.generate_plan(structure, "Apply formatting")
print(f"Generated {len(plan.ops)} operations")

# Process with pipeline
result = pipeline.process_document("document.docx", "Apply formatting")

# Analyze results
if result.status == "SUCCESS":
    print(f"Processing completed in {result.execution_time:.2f}s")
    print(f"Executed {result.operations_executed} operations")
```

### Batch Processing Integration

```python
import os
from pathlib import Path
from autoword.vnext.pipeline import VNextPipeline

def process_directory(input_dir: str, intent: str, output_dir: str):
    """Process all DOCX files in a directory"""
    
    pipeline = VNextPipeline({
        "model": "gpt4",
        "audit_dir": os.path.join(output_dir, "audits")
    })
    
    input_path = Path(input_dir)
    output_path = Path(output_dir)
    output_path.mkdir(exist_ok=True)
    
    results = []
    
    for docx_file in input_path.glob("*.docx"):
        print(f"Processing {docx_file.name}...")
        
        result = pipeline.process_document(str(docx_file), intent)
        results.append(result)
        
        if result.status == "SUCCESS":
            # Move processed file to output directory
            output_file = output_path / docx_file.name
            os.rename(result.output_path, str(output_file))
            print(f"  Success: {output_file}")
        else:
            print(f"  Failed: {result.status}")
    
    # Summary
    successful = sum(1 for r in results if r.status == "SUCCESS")
    print(f"\nProcessed {len(results)} files, {successful} successful")
    
    return results

# Usage
results = process_directory(
    "./input_documents",
    "Apply standard formatting and remove abstract sections",
    "./processed_documents"
)
```

This API reference provides comprehensive documentation for all public interfaces in AutoWord vNext, enabling effective programmatic integration and advanced usage scenarios.