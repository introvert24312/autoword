# AutoWord vNext - User Guide

## Introduction

AutoWord vNext is an advanced document processing system that automates Word document modifications through a structured pipeline. This guide will help you understand how to use the system effectively and configure it for your needs.

## Quick Start

### Installation

1. **Prerequisites**:
   - Python 3.8+
   - Microsoft Word 2016 or later
   - Windows operating system (for COM integration)

2. **Install AutoWord vNext**:
   ```bash
   pip install autoword-vnext
   ```

3. **Verify Installation**:
   ```bash
   python -m autoword.vnext.cli --version
   ```

### Basic Usage

#### Command Line Interface

```bash
# Process a document with user intent
python -m autoword.vnext.cli process document.docx "Remove abstract and references sections, update table of contents"

# Dry run (generate plan without execution)
python -m autoword.vnext.cli plan document.docx "Update all heading styles to use 楷体 font"

# Batch processing
python -m autoword.vnext.cli batch *.docx "Standardize document formatting"
```

#### Python API

```python
from autoword.vnext import VNextPipeline

# Initialize pipeline
pipeline = VNextPipeline()

# Process document
result = pipeline.process_document(
    docx_path="document.docx",
    user_intent="Remove abstract and references, update TOC"
)

if result.status == "SUCCESS":
    print(f"Document processed successfully: {result.output_path}")
else:
    print(f"Processing failed: {result.error}")
```

## Understanding the Pipeline

### The Five-Stage Process

1. **Extract**: Document structure and inventory extraction
2. **Plan**: LLM-generated execution plan creation
3. **Execute**: Atomic operations execution
4. **Validate**: Document integrity verification
5. **Audit**: Complete audit trail generation

### Data Flow

```
Your Document (DOCX) 
    ↓
[Extract] → structure.json + inventory.json
    ↓
[Plan] + Your Intent → plan.json (AI-generated)
    ↓
[Execute] → Modified Document
    ↓
[Validate] → Integrity Check
    ↓
[Audit] → Complete Record
```

## Configuration

### Basic Configuration

Create a configuration file `vnext_config.json`:

```json
{
  "llm": {
    "provider": "openai",
    "model": "gpt-4",
    "api_key": "your-api-key",
    "temperature": 0.1
  },
  "localization": {
    "language": "zh-CN",
    "style_aliases": {
      "Heading 1": "标题 1",
      "Heading 2": "标题 2",
      "Normal": "正文"
    },
    "font_fallbacks": {
      "楷体": ["楷体", "楷体_GB2312", "STKaiti"],
      "宋体": ["宋体", "SimSun"],
      "黑体": ["黑体", "SimHei"]
    }
  },
  "validation": {
    "strict_mode": true,
    "rollback_on_failure": true
  },
  "audit": {
    "save_snapshots": true,
    "generate_diff_reports": true
  }
}
```

### LLM Provider Configuration

#### OpenAI Configuration
```json
{
  "llm": {
    "provider": "openai",
    "model": "gpt-4",
    "api_key": "sk-...",
    "base_url": "https://api.openai.com/v1",
    "temperature": 0.1,
    "max_tokens": 4000
  }
}
```

#### Anthropic Claude Configuration
```json
{
  "llm": {
    "provider": "anthropic",
    "model": "claude-3-sonnet-20240229",
    "api_key": "sk-ant-...",
    "temperature": 0.1,
    "max_tokens": 4000
  }
}
```

### Localization Settings

#### Chinese Document Processing
```json
{
  "localization": {
    "language": "zh-CN",
    "style_aliases": {
      "Heading 1": "标题 1",
      "Heading 2": "标题 2",
      "Heading 3": "标题 3",
      "Normal": "正文",
      "Title": "标题"
    },
    "font_fallbacks": {
      "楷体": ["楷体", "楷体_GB2312", "STKaiti", "KaiTi"],
      "宋体": ["宋体", "SimSun", "NSimSun"],
      "黑体": ["黑体", "SimHei", "Microsoft YaHei"],
      "仿宋": ["仿宋", "FangSong", "FangSong_GB2312"]
    }
  }
}
```

## Common Use Cases

### 1. Academic Paper Processing

**Scenario**: Remove abstract and references sections, update TOC

```bash
python -m autoword.vnext.cli process paper.docx "Delete the abstract section and references section, then update the table of contents"
```

**Expected Operations**:
- `delete_section_by_heading` for "摘要" or "Abstract"
- `delete_section_by_heading` for "参考文献" or "References"
- `update_toc` to refresh page numbers

### 2. Document Formatting Standardization

**Scenario**: Apply consistent heading styles

```python
from autoword.vnext import VNextPipeline

pipeline = VNextPipeline()
result = pipeline.process_document(
    "document.docx",
    "Set all Heading 1 styles to use 楷体 font, 12pt, bold, with 2.0 line spacing"
)
```

**Expected Operations**:
- `set_style_rule` for "Heading 1" style modification

### 3. Batch Document Processing

**Scenario**: Process multiple documents with same formatting

```bash
# Process all DOCX files in directory
python -m autoword.vnext.cli batch documents/*.docx "Standardize heading fonts and update all TOCs"
```

### 4. TOC Management

**Scenario**: Remove and recreate table of contents

```python
result = pipeline.process_document(
    "document.docx",
    "Delete the existing table of contents and create a new one"
)
```

**Expected Operations**:
- `delete_toc` with mode "all"
- `update_toc` to create new TOC

## Understanding Output

### Success Output

When processing succeeds, you'll find:

```
output/
├── run_20240117_143022/
│   ├── before.docx                 # Original document
│   ├── after.docx                  # Modified document
│   ├── structure.before.json       # Original structure
│   ├── structure.after.json        # Modified structure
│   ├── plan.json                   # Execution plan
│   ├── diff.report.json           # Structural differences
│   ├── warnings.log               # Warnings and NOOPs
│   └── result.status.txt          # Final status
```

### Status Codes

- **SUCCESS**: Document processed successfully
- **ROLLBACK**: Processing failed, document restored
- **FAILED_VALIDATION**: Validation failed, document restored
- **INVALID_PLAN**: LLM generated invalid plan

### Reading Audit Reports

#### Structure Comparison
```json
{
  "schema_version": "diff.v1",
  "changes": [
    {
      "type": "section_deleted",
      "heading": "摘要",
      "level": 1,
      "paragraph_range": [5, 12]
    },
    {
      "type": "toc_updated",
      "field_index": 2,
      "page_changes": {
        "Introduction": "3 → 2",
        "Conclusion": "15 → 13"
      }
    }
  ]
}
```

#### Warnings Log
```
2024-01-17 14:30:22 - FONT_FALLBACK: 楷体 not available, using 楷体_GB2312
2024-01-17 14:30:23 - NOOP: delete_section_by_heading "Appendix" - heading not found
2024-01-17 14:30:24 - INFO: TOC updated successfully, 3 entries modified
```

## Advanced Usage

### Custom Atomic Operations

While the system uses predefined atomic operations, you can influence their behavior through detailed user intents:

```python
# Precise heading matching
result = pipeline.process_document(
    "document.docx",
    "Delete the section with heading 'Abstract' (exact match, case sensitive)"
)

# Occurrence-specific operations
result = pipeline.process_document(
    "document.docx", 
    "Delete the second occurrence of 'Introduction' heading"
)

# Style-specific operations
result = pipeline.process_document(
    "document.docx",
    "Apply 楷体 font to Heading 1 style, 宋体 font to Normal style, both 12pt with 2.0 line spacing"
)
```

### Validation Customization

```python
from autoword.vnext import VNextPipeline, ValidationConfig

config = ValidationConfig(
    strict_chapter_validation=True,
    required_styles=["Heading 1", "Heading 2", "Normal"],
    forbidden_level1_headings=["摘要", "参考文献", "Abstract", "References"]
)

pipeline = VNextPipeline(validation_config=config)
```

### Error Handling

```python
from autoword.vnext import VNextPipeline, VNextError

pipeline = VNextPipeline()

try:
    result = pipeline.process_document("document.docx", user_intent)
    
    if result.status == "SUCCESS":
        print(f"Success: {result.output_path}")
    elif result.status == "FAILED_VALIDATION":
        print(f"Validation failed: {result.validation_errors}")
    elif result.status == "INVALID_PLAN":
        print(f"Invalid plan: {result.plan_errors}")
        
except VNextError as e:
    print(f"Pipeline error: {e}")
    # Check audit directory for details
```

## Best Practices

### 1. Clear User Intents

**Good**: "Delete the abstract section and update the table of contents"
**Bad**: "Clean up the document"

**Good**: "Set Heading 1 to use 楷体 font, 12pt, bold"
**Bad**: "Fix the headings"

### 2. Document Preparation

- Ensure documents are not password-protected
- Close documents in Word before processing
- Backup important documents before processing
- Use documents with standard heading styles

### 3. Batch Processing

- Test with a single document first
- Use consistent user intents across batch
- Monitor warnings.log for issues
- Process similar document types together

### 4. Error Recovery

- Always check result.status before using output
- Review warnings.log for potential issues
- Keep original documents as backup
- Use dry-run mode for testing complex operations

### 5. Performance Optimization

- Process documents sequentially (COM limitation)
- Close other Word instances during processing
- Use specific rather than broad user intents
- Monitor memory usage for large documents

## Troubleshooting

See the [Troubleshooting Guide](TROUBLESHOOTING.md) for detailed solutions to common issues.

## API Reference

See the [API Documentation](API_REFERENCE.md) for complete interface documentation.

## Examples

See the `examples/` directory for complete working examples of common use cases.