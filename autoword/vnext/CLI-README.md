# AutoWord vNext CLI

Command-line interface for the AutoWord vNext pipeline, providing comprehensive document processing capabilities with advanced monitoring and configuration options.

## Quick Start

```bash
# Check system status
python -m autoword.vnext.cli status

# Process a single document
python -m autoword.vnext.cli process document.docx "Remove abstract and references"

# Batch process documents
python -m autoword.vnext.cli batch ./documents "Apply standard formatting"

# Dry run to see generated plan
python -m autoword.vnext.cli dry-run document.docx "Update TOC and styles"
```

## Commands

### process
Process a single DOCX document through the complete vNext pipeline.

```bash
python -m autoword.vnext.cli process <input.docx> "<user_intent>"
```

**Arguments:**
- `input.docx`: Path to the input DOCX file
- `user_intent`: Description of what you want to do with the document

**Example:**
```bash
python -m autoword.vnext.cli process paper.docx "Remove abstract and references sections, update TOC"
```

### batch
Process multiple DOCX documents in a directory with the same user intent.

```bash
python -m autoword.vnext.cli batch <directory> "<user_intent>"
```

**Arguments:**
- `directory`: Directory containing DOCX files to process
- `user_intent`: Description applied to all documents

**Features:**
- Processes all `.docx` files in the specified directory
- Creates individual audit trails for each document
- Generates a batch summary report
- Continues processing even if some documents fail

**Example:**
```bash
python -m autoword.vnext.cli batch ./research-papers "Standardize formatting and remove abstracts"
```

### dry-run
Generate an execution plan without actually modifying the document.

```bash
python -m autoword.vnext.cli dry-run <input.docx> "<user_intent>"
```

**Use cases:**
- Preview what operations will be performed
- Validate LLM plan generation
- Test user intent phrasing
- Debug planning issues

**Example:**
```bash
python -m autoword.vnext.cli dry-run document.docx "Update styles and TOC"
```

### status
Check system requirements and configuration status.

```bash
python -m autoword.vnext.cli status
```

**Checks:**
- Python version compatibility
- Required module availability
- Microsoft Word COM interface
- Environment variables (API keys)
- System resources (disk space, memory)

### config
Configuration management commands.

#### Show current configuration
```bash
python -m autoword.vnext.cli config show
```

#### Create configuration template
```bash
python -m autoword.vnext.cli config create <output.json>
```

## Global Options

### Basic Options

| Option | Description | Default |
|--------|-------------|---------|
| `--verbose`, `-v` | Enable verbose output and progress reporting | `false` |
| `--audit-dir` | Base directory for audit trails | `./audit_trails` |
| `--visible` | Show Word application windows during processing | `false` |
| `--log-file` | Log file path | Console only |

### LLM Configuration

| Option | Description | Default |
|--------|-------------|---------|
| `--model` | LLM model (`gpt4`, `gpt35`, `claude37`, `claude3`) | Auto-detect |
| `--temperature` | LLM temperature (0.0-1.0) | `0.1` |

### Monitoring Options

| Option | Description | Default |
|--------|-------------|---------|
| `--monitoring-level` | Monitoring detail level | `detailed` |
| `--disable-memory-monitoring` | Disable memory usage monitoring | Enabled |
| `--memory-warning-threshold` | Memory warning threshold (MB) | `1024` |
| `--memory-critical-threshold` | Memory critical threshold (MB) | `2048` |

#### Monitoring Levels

- **basic**: Essential operations only
- **detailed**: All operations with timing (recommended)
- **debug**: Full debug with memory tracking
- **performance**: Performance optimization focus

### Configuration Files

| Option | Description |
|--------|-------------|
| `--config` | Load configuration from JSON file |
| `--save-config` | Save current configuration to file |

## Configuration Files

Configuration files use JSON format and can contain any of the CLI options:

```json
{
  "model": "gpt4",
  "temperature": 0.1,
  "audit_dir": "./my-audits",
  "visible": false,
  "verbose": true,
  "monitoring_level": "detailed",
  "enable_memory_monitoring": true,
  "memory_warning_threshold": 1024,
  "memory_critical_threshold": 2048,
  "log_file": "vnext-pipeline.log"
}
```

**Usage:**
```bash
# Use configuration file
python -m autoword.vnext.cli process document.docx "Process" --config my-config.json

# Create template
python -m autoword.vnext.cli config create my-config.json

# Save current settings
python -m autoword.vnext.cli process document.docx "Test" --model gpt4 --save-config my-config.json
```

**Priority:** CLI arguments override configuration file values.

## Environment Variables

Set these environment variables for LLM API access:

```bash
# For OpenAI models (gpt4, gpt35)
export OPENAI_API_KEY="your-openai-api-key"

# For Anthropic models (claude37, claude3)
export ANTHROPIC_API_KEY="your-anthropic-api-key"
```

## Exit Codes

| Code | Status | Description |
|------|--------|-------------|
| 0 | SUCCESS | Processing completed successfully |
| 1 | ERROR | General error or exception |
| 2 | FAILED_VALIDATION | Document validation failed, changes rolled back |
| 3 | ROLLBACK | Processing failed, rollback performed |
| 4 | INVALID_PLAN | LLM generated invalid plan |
| 5 | UNKNOWN | Unknown status returned |

## Output and Audit Trails

### Audit Directory Structure
```
audit_trails/
├── run_20240101_120000/
│   ├── before.docx
│   ├── after.docx
│   ├── structure.before.json
│   ├── structure.after.json
│   ├── plan.json
│   ├── diff.report.json
│   ├── warnings.log
│   └── result.status.txt
└── batch_summary_20240101_120000.json
```

### Verbose Output
When using `--verbose`, the CLI provides:
- Progress reporting for each pipeline stage
- Performance metrics and timing
- Memory usage alerts
- Detailed error information
- Operation-by-operation execution details

## Examples

### Basic Usage
```bash
# Simple document processing
python -m autoword.vnext.cli process thesis.docx "Remove abstract and references"

# Check what would happen first
python -m autoword.vnext.cli dry-run thesis.docx "Remove abstract and references"
```

### Advanced Usage
```bash
# High-performance batch processing with debug monitoring
python -m autoword.vnext.cli batch ./documents "Standardize formatting" \
  --monitoring-level performance \
  --memory-warning-threshold 512 \
  --verbose \
  --log-file batch-processing.log

# Process with custom configuration
python -m autoword.vnext.cli process document.docx "Format document" \
  --config production-config.json \
  --model gpt4 \
  --audit-dir ./production-audits
```

### Configuration Management
```bash
# Create and customize configuration
python -m autoword.vnext.cli config create my-config.json
# Edit my-config.json as needed
python -m autoword.vnext.cli process document.docx "Process" --config my-config.json

# Save successful configuration for reuse
python -m autoword.vnext.cli process document.docx "Test run" \
  --model gpt4 \
  --monitoring-level debug \
  --verbose \
  --save-config production-config.json
```

## Troubleshooting

### Common Issues

1. **"Microsoft Word COM not available"**
   - Ensure Microsoft Word is installed
   - Run as administrator if needed
   - Check Windows COM registration

2. **"No LLM API keys found"**
   - Set `OPENAI_API_KEY` or `ANTHROPIC_API_KEY` environment variables
   - Verify API key validity

3. **Memory alerts during processing**
   - Increase memory thresholds: `--memory-warning-threshold 2048`
   - Use `--monitoring-level basic` for large documents
   - Close other applications to free memory

4. **Unicode encoding errors**
   - Ensure console supports UTF-8
   - Use `--log-file` to redirect output to file

### Getting Help

```bash
# General help
python -m autoword.vnext.cli --help

# Command-specific help
python -m autoword.vnext.cli process --help
python -m autoword.vnext.cli batch --help

# Check system status
python -m autoword.vnext.cli status
```

## Integration

The CLI can be integrated into scripts and automation workflows:

```bash
#!/bin/bash
# Batch processing script
for dir in research-papers conference-papers journal-papers; do
  echo "Processing $dir..."
  python -m autoword.vnext.cli batch "./$dir" "Standardize formatting" \
    --config production-config.json \
    --audit-dir "./audits/$dir" || echo "Failed to process $dir"
done
```

For programmatic access, use the Python API directly:
```python
from autoword.vnext.pipeline import VNextPipeline
from autoword.vnext.monitoring import MonitoringLevel

pipeline = VNextPipeline(
    monitoring_level=MonitoringLevel.DETAILED,
    enable_memory_monitoring=True
)
result = pipeline.process_document("document.docx", "Remove abstract")
```