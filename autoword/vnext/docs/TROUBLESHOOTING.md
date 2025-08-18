# AutoWord vNext Troubleshooting Guide

## Overview

This guide helps you diagnose and resolve common issues with AutoWord vNext. The system is designed to be robust and self-recovering, but understanding how to troubleshoot problems will help you use it more effectively.

## Quick Diagnostic Steps

### 1. Check System Status
Always start with a system status check:
```bash
python -m autoword.vnext status
```

This will verify:
- Python installation and version
- Required dependencies
- Microsoft Word availability
- COM automation capability
- LLM API configuration

### 2. Review Audit Trail
If processing started but failed, check the audit trail:
```bash
# Look for the most recent run directory
ls -la audit_trails/
# Check the status and logs
cat audit_trails/run_YYYYMMDD_HHMMSS_XXXXX/result.status.txt
cat audit_trails/run_YYYYMMDD_HHMMSS_XXXXX/warnings.log
```

### 3. Enable Verbose Output
Run with verbose output to see detailed progress:
```bash
python -m autoword.vnext process document.docx "intent" --verbose
```

### 4. Use Debug Monitoring
For complex issues, enable debug monitoring:
```bash
python -m autoword.vnext process document.docx "intent" --monitoring-level debug --verbose
```

## Common Issues and Solutions

### Installation and Setup Issues

#### "Module not found" errors
**Symptoms**: `ModuleNotFoundError` when running commands
**Cause**: Missing Python dependencies
**Solution**:
```bash
# Reinstall dependencies
pip install -r requirements-deployment.txt

# Verify installation
python -c "import autoword.vnext; print('OK')"
```

#### "Microsoft Word not found"
**Symptoms**: COM automation errors, Word not detected in status check
**Cause**: Word not installed or not properly registered
**Solution**:
1. Ensure Microsoft Word 2016+ is installed
2. Try opening Word manually to verify it works
3. Re-register Word COM components:
   ```cmd
   # Run as administrator
   regsvr32 /i msword.olb
   ```

#### "No LLM API key configured"
**Symptoms**: Planning stage fails with authentication errors
**Cause**: Missing API keys for OpenAI or Anthropic
**Solution**:
```bash
# Set OpenAI API key
set OPENAI_API_KEY=your_key_here

# Or set Anthropic API key
set ANTHROPIC_API_KEY=your_key_here

# Verify configuration
python -m autoword.vnext status
```

### Processing Issues

#### "No valid plan generated" (INVALID_PLAN)
**Symptoms**: Processing fails at planning stage
**Cause**: LLM couldn't generate a valid execution plan
**Common Reasons**:
- Intent too vague or unclear
- Document structure too complex
- LLM model limitations

**Solutions**:
1. **Make intent more specific**:
   ```bash
   # Instead of: "fix document"
   python -m autoword.vnext process doc.docx "Remove abstract section and update table of contents"
   
   # Instead of: "format better"
   python -m autoword.vnext process doc.docx "Apply Heading 1 style to chapter titles and set font to 楷体"
   ```

2. **Try different LLM model**:
   ```bash
   # Try GPT-4 for complex documents
   python -m autoword.vnext process doc.docx "intent" --model gpt4
   
   # Try Claude for alternative approach
   python -m autoword.vnext process doc.docx "intent" --model claude37
   ```

3. **Use dry-run to test**:
   ```bash
   python -m autoword.vnext dry-run doc.docx "intent" --verbose
   ```

#### "Validation failed" (FAILED_VALIDATION)
**Symptoms**: Processing completes but validation fails, document is rolled back
**Cause**: Changes didn't meet quality requirements
**Common Reasons**:
- Required sections still present
- Style formatting incorrect
- TOC inconsistencies

**Solutions**:
1. **Check validation details**:
   ```bash
   cat audit_trails/run_*/warnings.log
   cat audit_trails/run_*/diff.report.json
   ```

2. **Review validation rules**:
   - Chapter assertions: No "摘要/参考文献" at level 1
   - Style assertions: Correct fonts, sizes, spacing
   - TOC assertions: Consistency with headings

3. **Adjust intent for validation requirements**:
   ```bash
   # Ensure complete section removal
   python -m autoword.vnext process doc.docx "Completely remove abstract and references sections, then update TOC"
   ```

#### "Execution error" (ROLLBACK)
**Symptoms**: Processing fails during execution, document is rolled back
**Cause**: Technical error during atomic operations
**Common Reasons**:
- Document corruption
- COM automation failure
- Resource exhaustion

**Solutions**:
1. **Check document integrity**:
   - Open document manually in Word
   - Save as new file if corrupted
   - Accept/reject all track changes

2. **Restart Word COM**:
   ```bash
   # Kill any Word processes
   taskkill /f /im winword.exe
   
   # Try processing again
   python -m autoword.vnext process doc.docx "intent"
   ```

3. **Use visible mode for debugging**:
   ```bash
   python -m autoword.vnext process doc.docx "intent" --visible --verbose
   ```

### Performance Issues

#### "Memory usage too high"
**Symptoms**: Warnings about memory usage, slow processing
**Cause**: Large documents or insufficient system memory
**Solutions**:
1. **Increase memory thresholds**:
   ```bash
   python -m autoword.vnext process doc.docx "intent" \
     --memory-warning-threshold 2048 \
     --memory-critical-threshold 4096
   ```

2. **Use performance monitoring**:
   ```bash
   python -m autoword.vnext process doc.docx "intent" --monitoring-level performance
   ```

3. **Process smaller documents**:
   - Split large documents into sections
   - Process sections individually
   - Combine results manually

#### "Processing timeout"
**Symptoms**: Processing hangs or times out
**Cause**: Very large documents or complex operations
**Solutions**:
1. **Check system resources**:
   - Available memory
   - CPU usage
   - Disk space

2. **Use timeout settings**:
   ```bash
   # Custom configuration with longer timeouts
   python -m autoword.vnext process doc.docx "intent" --config long-timeout-config.json
   ```

3. **Simplify operations**:
   - Break complex intents into simpler ones
   - Process in multiple steps

### Document-Specific Issues

#### "Headings in tables not found"
**Symptoms**: Section deletion fails for headings inside tables
**Cause**: Complex document structure with headings in table cells
**Solutions**:
1. **Use occurrence index**:
   ```bash
   python -m autoword.vnext dry-run doc.docx "Remove second occurrence of abstract section" --verbose
   ```

2. **Check document structure**:
   ```bash
   # Extract structure to examine
   python -c "
   from autoword.vnext.extractor import DocumentExtractor
   extractor = DocumentExtractor()
   structure = extractor.extract_structure('doc.docx')
   print([h.text for h in structure.headings])
   "
   ```

#### "Font fallback warnings"
**Symptoms**: Warnings about unavailable fonts
**Cause**: Requested fonts not installed on system
**Solutions**:
1. **Install required fonts**:
   - 楷体, 宋体, 黑体 for Chinese documents
   - Times New Roman, Arial for English documents

2. **Check font fallback chain**:
   ```bash
   cat autoword/vnext/config/localization.json
   ```

3. **Customize font fallbacks**:
   ```json
   {
     "font_fallbacks": {
       "楷体": ["楷体", "楷体_GB2312", "STKaiti", "KaiTi"]
     }
   }
   ```

#### "Track changes detected"
**Symptoms**: Processing fails or behaves unexpectedly
**Cause**: Document has unresolved track changes
**Solutions**:
1. **Accept all changes**:
   - Open document in Word
   - Review → Accept All Changes

2. **Reject all changes**:
   - Open document in Word
   - Review → Reject All Changes

3. **Configure revision handling**:
   ```json
   {
     "revision_strategy": "accept_all"
   }
   ```

### Configuration Issues

#### "Configuration file invalid"
**Symptoms**: Errors loading configuration files
**Cause**: Invalid JSON syntax or structure
**Solutions**:
1. **Validate JSON syntax**:
   ```bash
   python -c "import json; json.load(open('config.json'))"
   ```

2. **Use configuration template**:
   ```bash
   python -m autoword.vnext config create new-config.json
   ```

3. **Check configuration schema**:
   ```bash
   python -m autoword.vnext config show
   ```

#### "Style aliases not working"
**Symptoms**: Styles not found despite being present
**Cause**: Localization configuration issues
**Solutions**:
1. **Check style names in document**:
   ```python
   from autoword.vnext.extractor import DocumentExtractor
   extractor = DocumentExtractor()
   structure = extractor.extract_structure('doc.docx')
   print([s.name for s in structure.styles])
   ```

2. **Update localization config**:
   ```json
   {
     "style_aliases": {
       "Heading 1": "标题 1",
       "Normal": "正文"
     }
   }
   ```

## Advanced Troubleshooting

### Debug Mode Analysis

Enable comprehensive debugging:
```bash
python -m autoword.vnext process doc.docx "intent" \
  --monitoring-level debug \
  --verbose \
  --log-file debug.log
```

This provides:
- Detailed operation logging
- COM automation traces
- Memory usage patterns
- Performance metrics
- Error stack traces

### Audit Trail Analysis

Examine complete audit trail:
```bash
# Navigate to audit directory
cd audit_trails/run_YYYYMMDD_HHMMSS_XXXXX/

# Check final status
cat result.status.txt

# Review warnings and issues
cat warnings.log

# Compare before/after structures
python -c "
import json
before = json.load(open('before_structure.v1.json'))
after = json.load(open('after_structure.v1.json'))
print('Before headings:', len(before['headings']))
print('After headings:', len(after['headings']))
"

# Review execution plan
cat plan.v1.json | python -m json.tool
```

### System Environment Debugging

Check system environment:
```bash
# Python environment
python --version
pip list | grep -E "(pydantic|requests|win32com)"

# Word installation
reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office" /s | findstr Word

# System resources
wmic OS get TotalVisibleMemorySize,FreePhysicalMemory
wmic cpu get loadpercentage /value
```

### COM Automation Debugging

Test COM automation directly:
```python
import win32com.client

try:
    word = win32com.client.Dispatch("Word.Application")
    print("Word COM available")
    word.Visible = True
    doc = word.Documents.Open(r"C:\path\to\document.docx")
    print(f"Document opened: {doc.Name}")
    doc.Close()
    word.Quit()
    print("COM automation working")
except Exception as e:
    print(f"COM error: {e}")
```

## Error Code Reference

### Exit Codes
- `0`: Success
- `1`: General error
- `2`: Failed validation (FAILED_VALIDATION)
- `3`: Rollback performed (ROLLBACK)
- `4`: Invalid plan (INVALID_PLAN)
- `5`: Unknown status

### Status Codes
- `SUCCESS`: Processing completed successfully
- `FAILED_VALIDATION`: Changes didn't meet validation requirements
- `ROLLBACK`: Error occurred, changes rolled back
- `INVALID_PLAN`: LLM couldn't generate valid plan

### Common Error Messages

#### "Schema validation failed"
**Meaning**: JSON data doesn't match expected schema
**Action**: Check data structure and types

#### "Operation not in whitelist"
**Meaning**: Plan contains unauthorized operations
**Action**: Review plan generation and constraints

#### "COM automation failed"
**Meaning**: Word COM interface error
**Action**: Restart Word, check installation

#### "Document validation failed"
**Meaning**: Modified document doesn't meet requirements
**Action**: Review validation rules and intent

## Prevention Strategies

### Document Preparation
1. **Clean documents before processing**:
   - Accept/reject all track changes
   - Remove comments and annotations
   - Fix any corruption issues

2. **Backup important documents**:
   - Always keep original copies
   - Use version control for critical documents

3. **Test with sample documents**:
   - Use dry-run mode first
   - Test with copies, not originals

### System Maintenance
1. **Keep software updated**:
   - Update Python and dependencies
   - Keep Word installation current
   - Monitor system resources

2. **Regular system checks**:
   ```bash
   # Weekly system status check
   python -m autoword.vnext status
   ```

3. **Monitor audit trails**:
   - Review processing logs regularly
   - Clean up old audit directories
   - Monitor disk space usage

### Best Practices
1. **Write clear intents**:
   - Be specific about desired changes
   - Use standard terminology
   - Test with dry-run first

2. **Use appropriate models**:
   - GPT-4 for complex documents
   - GPT-3.5 for simple operations
   - Claude for alternative approaches

3. **Monitor performance**:
   - Use performance monitoring for large documents
   - Set appropriate memory thresholds
   - Process in batches when needed

## Getting Additional Help

### Information to Collect
When reporting issues, include:
1. **System status output**:
   ```bash
   python -m autoword.vnext status > system-status.txt
   ```

2. **Complete command and output**:
   ```bash
   python -m autoword.vnext process doc.docx "intent" --verbose > output.log 2>&1
   ```

3. **Audit trail directory** (if processing started)

4. **Document characteristics**:
   - File size
   - Number of pages
   - Complexity (tables, images, etc.)

5. **System information**:
   - Windows version
   - Word version
   - Python version
   - Available memory

### Resources
- **KNOWN-ISSUES.md**: Known limitations and workarounds
- **Configuration examples**: Sample configuration files
- **Test scenarios**: Example documents and expected results
- **API documentation**: Complete interface reference

This troubleshooting guide covers the most common issues and their solutions. The AutoWord vNext system is designed to be robust and self-recovering, but understanding these troubleshooting techniques will help you resolve issues quickly and effectively.