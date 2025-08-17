# AutoWord vNext Pipeline

## Overview

AutoWord vNext is a structured document processing pipeline that transforms Word document modifications into a predictable, reproducible "Extract→Plan→Execute→Validate→Audit" workflow. The system achieves >99% stability through strict JSON schemas, atomic operations, and comprehensive validation layers.

## Key Features

- **Zero Information Loss**: Complete document inventory with OOXML fragment preservation
- **LLM-Driven Planning**: AI-generated execution plans with strict schema validation
- **Atomic Operations**: Whitelisted operations through Word COM automation
- **Comprehensive Validation**: Multi-layer assertions with automatic rollback
- **Complete Audit Trails**: Timestamped snapshots and diff reports
- **Localization Support**: Chinese/English style aliases and font fallbacks

## Architecture

```
Input DOCX → [Extractor] → structure.v1.json + inventory.full.v1.json
                ↓
User Intent → [Planner] → plan.v1.json (LLM generated)
                ↓
Original DOCX → [Executor] → Modified DOCX (atomic operations)
                ↓
[Validator] → Assertions + Rollback on failure
                ↓
[Auditor] → Complete audit trail with before/after snapshots
```

## Quick Start

### Installation

```bash
# Install dependencies
pip install -r requirements.txt

# Verify installation
python -m autoword.vnext --version
```

### Basic Usage

```bash
# Process a document
python -m autoword.vnext process document.docx "删除摘要和参考文献章节，更新目录"

# Run with specific configuration
python -m autoword.vnext process document.docx "用户意图" --config config/custom.json

# Batch processing
python -m autoword.vnext batch input_dir/ output_dir/ "标准化格式"
```

### Python API

```python
from autoword.vnext.pipeline import VNextPipeline

# Initialize pipeline
pipeline = VNextPipeline()

# Process document
result = pipeline.process_document(
    docx_path="document.docx",
    user_intent="删除摘要章节并更新目录"
)

# Check result
if result.status == "SUCCESS":
    print(f"Document processed successfully: {result.output_path}")
    print(f"Audit trail: {result.audit_directory}")
else:
    print(f"Processing failed: {result.error}")
```

## Core Components

### 1. Extractor Module
Converts DOCX to structured JSON with zero information loss.

**Output**: 
- `structure.v1.json` - Document skeleton and metadata
- `inventory.full.v1.json` - Complete OOXML fragments and media

### 2. Planner Module  
Generates execution plans through LLM with strict schema validation.

**Input**: structure.v1.json + user intent
**Output**: plan.v1.json with whitelisted atomic operations

### 3. Executor Module
Executes atomic operations through Word COM automation.

**Supported Operations**:
- `delete_section_by_heading` - Remove document sections
- `update_toc` / `delete_toc` - TOC management
- `set_style_rule` - Style object modification
- `reassign_paragraphs_to_style` - Paragraph style assignment
- `clear_direct_formatting` - Remove direct formatting

### 4. Validator Module
Validates modifications against strict assertions with rollback capability.

**Assertions**:
- Chapter assertions (no "摘要/参考文献" at level 1)
- Style assertions (H1/H2/Normal specifications)
- TOC assertions (consistency with heading tree)
- Pagination assertions (fields updated, metadata changed)

### 5. Auditor Module
Creates complete audit trails with timestamped snapshots.

**Audit Files**:
- `before.docx` / `after.docx` - Document snapshots
- `before_structure.v1.json` / `after_structure.v1.json` - Structure comparison
- `plan.v1.json` - Executed plan
- `diff.report.json` - Structural differences
- `warnings.log` - Processing warnings
- `result.status.txt` - Final status (SUCCESS/ROLLBACK/FAILED_VALIDATION)

## Configuration

### Localization Settings (`config/localization.json`)

```json
{
  "style_aliases": {
    "Heading 1": "标题 1",
    "Normal": "正文"
  },
  "font_fallbacks": {
    "楷体": ["楷体", "楷体_GB2312", "STKaiti"]
  }
}
```

### Pipeline Settings (`config/pipeline.json`)

```json
{
  "pipeline_settings": {
    "max_execution_time_seconds": 300,
    "enable_rollback": true,
    "enable_audit_trail": true
  }
}
```

### Validation Rules (`config/validation_rules.json`)

```json
{
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
```

## JSON Schemas

### Structure V1 Schema
Document structure representation with metadata, styles, paragraphs, headings, fields, and tables.

See: `schemas/structure.v1.md` for complete documentation.

### Plan V1 Schema  
LLM-generated execution plans with whitelisted atomic operations only.

See: `schemas/plan.v1.md` for complete documentation.

### Inventory Full V1 Schema
Complete document inventory with OOXML fragments and media references.

## Test Data

The `test_data/` directory contains 7 validation scenarios:

1. **Normal Paper Processing** - Standard workflow
2. **No-TOC Document** - NOOP operation handling  
3. **Duplicate Headings** - Occurrence index specification
4. **Headings in Tables** - Complex structure handling
5. **Missing Fonts** - Font fallback chains
6. **Complex Objects** - Formulas, charts, content controls
7. **Revision Tracking** - Track changes handling

## Error Handling

### Automatic Rollback
Any exception during processing triggers automatic rollback to original document with `FAILED_VALIDATION` status.

### Validation Failures
Failed assertions trigger rollback with detailed error reporting in audit trail.

### NOOP Operations
Operations that find no targets are logged as NOOP in `warnings.log` but don't cause failure.

## Security Features

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

## Performance Considerations

### System Requirements
- Windows 10/11 with Microsoft Word 2016+
- Python 3.8+ (recommended: 3.10+)
- 4GB+ RAM (8GB+ recommended for large documents)

### Optimization
- Efficient COM object management
- Document structure caching
- Memory-efficient JSON processing
- Configurable timeout settings

## Known Issues

See `KNOWN-ISSUES.md` for detailed documentation of edge cases, limitations, and workarounds.

## Development

### Running Tests

```bash
# Run all tests
python -m pytest tests/

# Run specific test module
python -m pytest tests/test_vnext_pipeline.py

# Run with coverage
python -m pytest --cov=autoword.vnext tests/
```

### Code Quality

```bash
# Format code
black autoword/vnext/

# Lint code  
flake8 autoword/vnext/

# Type checking
mypy autoword/vnext/
```

## Deployment

See `DEPLOYMENT.md` for complete deployment guide including:
- System requirements and installation
- Configuration management
- Security considerations
- Monitoring and maintenance
- Troubleshooting procedures

## License

[License information]

## Support

For technical support:
1. Check `KNOWN-ISSUES.md` for common problems
2. Review audit trails for detailed error information
3. Enable debug logging for troubleshooting
4. Provide system information and sample documents when reporting issues

## Version History

- **v1.0.0**: Initial release with complete pipeline implementation
- All 5 modules implemented and tested
- 7 DoD validation scenarios passing
- Complete configuration and deployment system