# AutoWord vNext User Guide

## Introduction

AutoWord vNext is a powerful document processing pipeline that transforms Word document modifications into a predictable, reproducible workflow. This guide will help you understand how to use the system effectively for your document processing needs.

## Getting Started

### System Requirements

- **Operating System**: Windows 10/11
- **Microsoft Word**: 2016 or later (required for COM automation)
- **Python**: 3.8 or later (recommended: 3.10+)
- **Memory**: 4GB RAM minimum (8GB+ recommended for large documents)
- **Storage**: 1GB free space for audit trails and temporary files

### Installation

1. **Install Python Dependencies**:
   ```bash
   pip install -r requirements-deployment.txt
   ```

2. **Verify Installation**:
   ```bash
   python -m autoword.vnext --version
   ```

3. **Check System Status**:
   ```bash
   python -m autoword.vnext status
   ```

### Quick Start

Process your first document:

```bash
python -m autoword.vnext process document.docx "Remove abstract and references sections"
```

## Understanding the Pipeline

### The Five-Stage Process

AutoWord vNext processes documents through five distinct stages:

1. **Extract**: Convert DOCX to structured JSON
2. **Plan**: Generate execution plan using AI
3. **Execute**: Apply changes through atomic operations
4. **Validate**: Verify changes meet requirements
5. **Audit**: Create complete record of all changes

### What Happens During Processing

```
Your Document → [Extract] → Document Structure (JSON)
                     ↓
Your Intent → [Plan] → Execution Plan (JSON)
                     ↓
Original Document → [Execute] → Modified Document
                     ↓
[Validate] → Check if changes are correct
                     ↓
[Audit] → Save complete record of all changes
```

## Basic Usage

### Processing a Single Document

The most common operation is processing a single document:

```bash
# Basic processing
python -m autoword.vnext process document.docx "Your processing intent"

# With verbose output to see what's happening
python -m autoword.vnext process document.docx "Your intent" --verbose

# Save audit trail to specific location
python -m autoword.vnext process document.docx "Your intent" --audit-dir ./my-audits
```

### Common Processing Intents

Here are examples of typical processing requests:

#### Remove Sections
```bash
python -m autoword.vnext process paper.docx "Remove abstract and references sections"
python -m autoword.vnext process document.docx "Delete the introduction and conclusion chapters"
```

#### Update Table of Contents
```bash
python -m autoword.vnext process document.docx "Update the table of contents"
python -m autoword.vnext process paper.docx "Refresh TOC and page numbers"
```

#### Apply Formatting
```bash
python -m autoword.vnext process document.docx "Apply standard heading styles"
python -m autoword.vnext process paper.docx "Format headings with proper fonts and spacing"
```

#### Combined Operations
```bash
python -m autoword.vnext process document.docx "Remove abstract section, update TOC, and apply standard formatting"
```

### Batch Processing

Process multiple documents with the same intent:

```bash
# Process all DOCX files in a directory
python -m autoword.vnext batch ./documents "Apply standard formatting"

# With verbose output
python -m autoword.vnext batch ./documents "Remove sections" --verbose
```

### Dry Run Mode

Test what changes would be made without actually modifying documents:

```bash
# See what would happen without making changes
python -m autoword.vnext dry-run document.docx "Remove abstract section"

# Dry run with verbose output
python -m autoword.vnext dry-run document.docx "Update formatting" --verbose
```

## Understanding Results

### Success Indicators

When processing succeeds, you'll see:
- **Status**: `SUCCESS`
- **Output**: Modified document saved
- **Audit Trail**: Complete record in audit directory

Example output:
```
Processing completed successfully!
Status: SUCCESS
Modified document: document_processed.docx
Audit trail: ./audit_trails/run_20240117_143022_12345/
```

### What's in the Audit Trail

Each processing run creates a timestamped directory with:

- `before.docx` - Your original document
- `after.docx` - The modified document
- `before_structure.v1.json` - Original document structure
- `after_structure.v1.json` - Modified document structure
- `plan.v1.json` - The execution plan that was used
- `diff.report.json` - Summary of changes made
- `warnings.log` - Any warnings or issues encountered
- `result.status.txt` - Final processing status

### Handling Failures

If processing fails, the system automatically:
1. **Rolls back** any partial changes
2. **Restores** your original document
3. **Creates audit trail** showing what went wrong
4. **Provides detailed error information**

Common failure scenarios:
- **Invalid Plan**: AI couldn't generate a valid execution plan
- **Validation Failed**: Changes didn't meet quality requirements
- **Execution Error**: Technical issue during processing

## Advanced Usage

### Configuration Files

Create a configuration file to customize behavior:

```bash
# Create configuration template
python -m autoword.vnext config create my-config.json
```

Example configuration:
```json
{
  "model": "gpt4",
  "temperature": 0.1,
  "audit_dir": "./my-audits",
  "verbose": true,
  "monitoring_level": "detailed"
}
```

Use your configuration:
```bash
python -m autoword.vnext process document.docx "Intent" --config my-config.json
```

### Choosing AI Models

Different AI models are available for plan generation:

```bash
# Use GPT-4 (most capable, slower)
python -m autoword.vnext process document.docx "Intent" --model gpt4

# Use GPT-3.5 (faster, less capable)
python -m autoword.vnext process document.docx "Intent" --model gpt35

# Use Claude (alternative provider)
python -m autoword.vnext process document.docx "Intent" --model claude37
```

### Monitoring and Debugging

Enable detailed monitoring:

```bash
# Basic monitoring
python -m autoword.vnext process document.docx "Intent" --monitoring-level basic

# Detailed monitoring (default)
python -m autoword.vnext process document.docx "Intent" --monitoring-level detailed

# Debug monitoring (very verbose)
python -m autoword.vnext process document.docx "Intent" --monitoring-level debug

# Performance monitoring
python -m autoword.vnext process document.docx "Intent" --monitoring-level performance
```

## Best Practices

### Writing Effective Intents

**Good intents are**:
- **Specific**: "Remove the abstract section" vs "Clean up document"
- **Clear**: "Update table of contents" vs "Fix TOC"
- **Actionable**: "Apply Heading 1 style to chapter titles" vs "Make it look better"

**Examples of effective intents**:
```
"Remove abstract and references sections, then update table of contents"
"Apply standard academic formatting with proper heading styles"
"Delete introduction chapter and update page numbering"
"Format all headings with correct fonts and spacing according to style guide"
```

### Document Preparation

**Before processing**:
- **Save your work**: Always have a backup
- **Close the document**: Don't have it open in Word
- **Check for track changes**: Accept or reject changes first
- **Verify document integrity**: Ensure document opens correctly

### Managing Large Documents

For large documents (>50 pages):
- **Use performance monitoring**: `--monitoring-level performance`
- **Increase memory thresholds**: `--memory-warning-threshold 2048`
- **Process in smaller batches**: Break into sections if possible
- **Monitor system resources**: Watch memory usage

### Troubleshooting Common Issues

#### "No valid plan generated"
- **Cause**: AI couldn't understand your intent
- **Solution**: Make your intent more specific and clear
- **Example**: Instead of "fix document", try "remove abstract section and update TOC"

#### "Validation failed"
- **Cause**: Changes didn't meet quality requirements
- **Solution**: Check audit trail for specific validation errors
- **Action**: Document is automatically restored to original state

#### "COM automation error"
- **Cause**: Issue with Microsoft Word integration
- **Solution**: Ensure Word is installed and not running
- **Check**: Run system status check

#### "Memory issues"
- **Cause**: Document too large for available memory
- **Solution**: Increase memory thresholds or process smaller documents
- **Command**: Use `--memory-warning-threshold` and `--memory-critical-threshold`

## Working with Different Document Types

### Academic Papers
```bash
# Common academic paper processing
python -m autoword.vnext process paper.docx "Remove abstract and references, update TOC, apply standard formatting"
```

### Business Documents
```bash
# Business document standardization
python -m autoword.vnext process report.docx "Apply corporate style guide formatting"
```

### Technical Documentation
```bash
# Technical document processing
python -m autoword.vnext process manual.docx "Update table of contents and apply consistent heading styles"
```

## Integration with Workflows

### Batch Processing Scripts

Create scripts for repeated operations:

```bash
#!/bin/bash
# Process all documents in projects directory
for dir in ./projects/*/documents; do
    echo "Processing $dir"
    python -m autoword.vnext batch "$dir" "Apply standard formatting" --config production.json
done
```

### Automated Quality Control

Use dry-run mode for quality checks:

```bash
# Check what would be changed
python -m autoword.vnext dry-run document.docx "Apply formatting" --verbose > changes.log
```

## Getting Help

### Built-in Help
```bash
# General help
python -m autoword.vnext --help

# Command-specific help
python -m autoword.vnext process --help
python -m autoword.vnext batch --help
```

### System Information
```bash
# Check system status and requirements
python -m autoword.vnext status
```

### Diagnostic Information

When reporting issues, include:
1. **System status output**
2. **Complete command used**
3. **Audit trail directory** (if processing started)
4. **Error messages** (exact text)
5. **Document characteristics** (size, complexity)

### Common Resources

- **KNOWN-ISSUES.md**: Known limitations and workarounds
- **Audit trails**: Complete processing records
- **Configuration examples**: Sample configuration files
- **Test scenarios**: Example documents and expected results

## Conclusion

AutoWord vNext provides a powerful, reliable way to process Word documents with complete auditability and automatic error recovery. By following the practices in this guide, you can effectively use the system for your document processing needs while maintaining confidence in the results.

Remember:
- **Start simple**: Begin with basic operations and build complexity
- **Use dry-run**: Test changes before applying them
- **Check audit trails**: Review what actually happened
- **Keep backups**: Always maintain original documents
- **Monitor performance**: Watch system resources for large operations

The system is designed to be safe and predictable - when in doubt, it will preserve your original document rather than risk making incorrect changes.