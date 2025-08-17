# AutoWord vNext - Troubleshooting Guide

## Common Issues and Solutions

### Installation and Setup Issues

#### Issue: "Module not found: autoword.vnext"

**Symptoms**:
```
ModuleNotFoundError: No module named 'autoword.vnext'
```

**Solutions**:
1. **Verify Installation**:
   ```bash
   pip list | grep autoword
   ```

2. **Reinstall Package**:
   ```bash
   pip uninstall autoword-vnext
   pip install autoword-vnext
   ```

3. **Check Python Environment**:
   ```bash
   python -c "import sys; print(sys.path)"
   ```

4. **Virtual Environment Issues**:
   ```bash
   # Activate correct environment
   source venv/bin/activate  # Linux/Mac
   venv\Scripts\activate     # Windows
   ```

#### Issue: "Word COM automation failed"

**Symptoms**:
```
COM Error: Microsoft Word is not installed or not accessible
```

**Solutions**:
1. **Verify Word Installation**:
   - Ensure Microsoft Word 2016+ is installed
   - Try opening Word manually
   - Check Word version: File → Account → About Word

2. **COM Registration**:
   ```cmd
   # Run as Administrator
   regsvr32 "C:\Program Files\Microsoft Office\root\Office16\MSWORD.OLB"
   ```

3. **Word Process Cleanup**:
   ```cmd
   taskkill /f /im winword.exe
   taskkill /f /im WINWORD.EXE
   ```

4. **Permissions**:
   - Run Python as Administrator
   - Check DCOM permissions for Word

#### Issue: "LLM API connection failed"

**Symptoms**:
```
OpenAI API Error: Authentication failed
Anthropic API Error: Invalid API key
```

**Solutions**:
1. **Verify API Key**:
   ```python
   import os
   print(os.getenv('OPENAI_API_KEY'))  # Should not be None
   ```

2. **Check Configuration**:
   ```json
   {
     "llm": {
       "provider": "openai",
       "api_key": "sk-...",
       "model": "gpt-4"
     }
   }
   ```

3. **Network Connectivity**:
   ```bash
   curl -I https://api.openai.com/v1/models
   ```

4. **Rate Limits**:
   - Check API usage dashboard
   - Implement retry logic
   - Use lower rate limits

### Document Processing Issues

#### Issue: "Document locked or in use"

**Symptoms**:
```
COM Error: Document is locked for editing
Permission denied: Cannot open document
```

**Solutions**:
1. **Close Word Instances**:
   ```cmd
   taskkill /f /im winword.exe
   ```

2. **Check File Permissions**:
   - Ensure file is not read-only
   - Check network drive permissions
   - Verify user has write access

3. **Document Recovery**:
   ```python
   # Use read-only mode for testing
   pipeline = VNextPipeline(read_only=True)
   ```

4. **Temporary File Issues**:
   - Clear temp directory: `%TEMP%`
   - Restart computer if persistent

#### Issue: "Structure extraction failed"

**Symptoms**:
```
ExtractionError: Unable to parse document structure
JSON serialization failed
```

**Solutions**:
1. **Document Validation**:
   - Open document in Word manually
   - Check for corruption: File → Info → Check for Issues
   - Save as new DOCX file

2. **Complex Elements**:
   ```python
   # Enable verbose logging
   import logging
   logging.basicConfig(level=logging.DEBUG)
   ```

3. **Memory Issues**:
   - Process smaller documents first
   - Increase available memory
   - Close other applications

4. **Encoding Issues**:
   - Ensure document uses UTF-8 encoding
   - Check for special characters
   - Validate file format

#### Issue: "Plan generation failed"

**Symptoms**:
```
PlanningError: LLM returned invalid JSON
Schema validation failed
INVALID_PLAN status returned
```

**Solutions**:
1. **User Intent Clarity**:
   ```python
   # Good intent
   "Delete the abstract section and update table of contents"
   
   # Bad intent  
   "Fix the document"
   ```

2. **LLM Model Issues**:
   ```json
   {
     "llm": {
       "model": "gpt-4",  # Use more capable model
       "temperature": 0.1,  # Lower temperature for consistency
       "max_tokens": 4000
     }
   }
   ```

3. **Schema Debugging**:
   ```python
   from autoword.vnext.planner import SchemaValidator
   
   validator = SchemaValidator()
   result = validator.validate_plan(plan_json)
   print(result.errors)  # Detailed error messages
   ```

4. **Retry Logic**:
   ```python
   # Implement retry with different prompts
   for attempt in range(3):
       try:
           plan = planner.generate_plan(structure, intent)
           break
       except PlanningError:
           intent = f"Attempt {attempt + 1}: {intent}"
   ```

#### Issue: "Execution failed with rollback"

**Symptoms**:
```
ExecutionError: Operation failed
Status: ROLLBACK
Document restored to original state
```

**Solutions**:
1. **Operation Debugging**:
   ```python
   # Enable detailed operation logging
   pipeline = VNextPipeline(debug_mode=True)
   ```

2. **Heading Not Found**:
   - Check exact heading text in document
   - Verify heading levels
   - Use case-insensitive matching

3. **Style Issues**:
   ```python
   # Check available styles
   extractor = DocumentExtractor()
   structure = extractor.extract_structure("document.docx")
   print([style.name for style in structure.styles])
   ```

4. **COM Automation Errors**:
   - Restart Word application
   - Check Word version compatibility
   - Verify COM registration

#### Issue: "Validation failed"

**Symptoms**:
```
ValidationError: Document validation failed
Status: FAILED_VALIDATION
Assertion failures: [...]
```

**Solutions**:
1. **Assertion Analysis**:
   ```python
   # Check specific assertion failures
   result = pipeline.process_document(doc_path, intent)
   if result.status == "FAILED_VALIDATION":
       print(result.validation_errors)
   ```

2. **Chapter Assertions**:
   - Verify no "摘要/参考文献" at level 1
   - Check heading structure after modifications
   - Ensure proper section deletion

3. **Style Assertions**:
   - Verify H1/H2/Normal style specifications
   - Check font availability and fallbacks
   - Validate line spacing and formatting

4. **TOC Assertions**:
   - Ensure TOC consistency with headings
   - Check page number updates
   - Verify TOC field integrity

### Performance Issues

#### Issue: "Processing very slow"

**Symptoms**:
- Long processing times (>5 minutes for normal documents)
- High memory usage
- System unresponsiveness

**Solutions**:
1. **Document Size**:
   - Split large documents into smaller sections
   - Remove unnecessary embedded objects
   - Compress images and media

2. **System Resources**:
   ```python
   # Monitor memory usage
   import psutil
   process = psutil.Process()
   print(f"Memory: {process.memory_info().rss / 1024 / 1024:.1f} MB")
   ```

3. **COM Optimization**:
   ```python
   # Batch operations
   pipeline = VNextPipeline(batch_operations=True)
   ```

4. **Parallel Processing**:
   - Process multiple documents sequentially (COM limitation)
   - Use separate Python processes for batch jobs

#### Issue: "Memory errors"

**Symptoms**:
```
MemoryError: Unable to allocate memory
OutOfMemoryError: COM object creation failed
```

**Solutions**:
1. **Memory Management**:
   ```python
   # Explicit cleanup
   import gc
   gc.collect()
   ```

2. **Document Chunking**:
   - Process documents in smaller sections
   - Use streaming JSON processing
   - Implement pagination for large structures

3. **System Configuration**:
   - Increase virtual memory
   - Close unnecessary applications
   - Use 64-bit Python

### Localization Issues

#### Issue: "Font fallback not working"

**Symptoms**:
```
Font '楷体' not found, no fallback applied
Style formatting incorrect
```

**Solutions**:
1. **Font Installation**:
   - Install required fonts system-wide
   - Check font names in Windows Fonts folder
   - Verify font licensing

2. **Fallback Configuration**:
   ```json
   {
     "localization": {
       "font_fallbacks": {
         "楷体": ["楷体", "楷体_GB2312", "STKaiti", "KaiTi"],
         "宋体": ["宋体", "SimSun", "NSimSun"]
       }
     }
   }
   ```

3. **Font Detection**:
   ```python
   from autoword.vnext.executor import LocalizationManager
   
   manager = LocalizationManager()
   available_fonts = manager.get_available_fonts()
   print(available_fonts)
   ```

#### Issue: "Style aliases not working"

**Symptoms**:
- English style names not mapping to Chinese
- Style operations failing
- Inconsistent style application

**Solutions**:
1. **Alias Configuration**:
   ```json
   {
     "localization": {
       "style_aliases": {
         "Heading 1": "标题 1",
         "Heading 2": "标题 2",
         "Normal": "正文"
       }
     }
   }
   ```

2. **Document Language**:
   - Check document language settings
   - Verify Word UI language
   - Test with both English and Chinese style names

3. **Dynamic Detection**:
   ```python
   # Let system detect available styles
   manager = LocalizationManager()
   style_mapping = manager.detect_style_mapping(document)
   ```

### Audit and Logging Issues

#### Issue: "Audit files not generated"

**Symptoms**:
- Missing audit directory
- Incomplete audit files
- No diff reports generated

**Solutions**:
1. **Permissions**:
   - Check write permissions to output directory
   - Ensure sufficient disk space
   - Verify directory creation rights

2. **Configuration**:
   ```json
   {
     "audit": {
       "save_snapshots": true,
       "generate_diff_reports": true,
       "output_directory": "./audit_output"
     }
   }
   ```

3. **Manual Audit**:
   ```python
   from autoword.vnext.auditor import DocumentAuditor
   
   auditor = DocumentAuditor()
   audit_dir = auditor.create_audit_directory()
   print(f"Audit directory: {audit_dir}")
   ```

#### Issue: "Logs not helpful"

**Symptoms**:
- Generic error messages
- Missing operation details
- No debugging information

**Solutions**:
1. **Logging Configuration**:
   ```python
   import logging
   
   logging.basicConfig(
       level=logging.DEBUG,
       format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
       handlers=[
           logging.FileHandler('vnext_debug.log'),
           logging.StreamHandler()
       ]
   )
   ```

2. **Verbose Mode**:
   ```python
   pipeline = VNextPipeline(verbose=True, debug_mode=True)
   ```

3. **Custom Logging**:
   ```python
   # Add custom log handlers
   from autoword.vnext.core import get_logger
   
   logger = get_logger(__name__)
   logger.info("Custom debug information")
   ```

## Diagnostic Tools

### System Information

```python
def diagnose_system():
    """Collect system diagnostic information"""
    import sys
    import platform
    import win32api
    
    print(f"Python: {sys.version}")
    print(f"Platform: {platform.platform()}")
    print(f"Word Version: {win32api.GetFileVersionInfo('winword.exe')}")
    print(f"Available Memory: {psutil.virtual_memory().available / 1024**3:.1f} GB")
```

### Document Analysis

```python
def analyze_document(docx_path):
    """Analyze document for potential issues"""
    from autoword.vnext.extractor import DocumentExtractor
    
    extractor = DocumentExtractor()
    try:
        structure = extractor.extract_structure(docx_path)
        print(f"Paragraphs: {len(structure.paragraphs)}")
        print(f"Headings: {len(structure.headings)}")
        print(f"Styles: {len(structure.styles)}")
        print(f"Tables: {len(structure.tables)}")
        print(f"Fields: {len(structure.fields)}")
    except Exception as e:
        print(f"Analysis failed: {e}")
```

### Configuration Validation

```python
def validate_configuration(config_path):
    """Validate configuration file"""
    import json
    from jsonschema import validate
    
    with open(config_path) as f:
        config = json.load(f)
    
    # Validate against schema
    try:
        validate(config, CONFIG_SCHEMA)
        print("Configuration valid")
    except Exception as e:
        print(f"Configuration error: {e}")
```

## Getting Help

### Debug Information Collection

When reporting issues, please include:

1. **System Information**:
   ```bash
   python --version
   pip list | grep autoword
   ```

2. **Error Messages**:
   - Complete error traceback
   - Log file contents
   - Configuration used

3. **Document Information**:
   - Document size and complexity
   - Word version used to create
   - Any special elements (tables, TOC, etc.)

4. **Reproduction Steps**:
   - Exact command or code used
   - User intent provided
   - Expected vs actual behavior

### Support Channels

1. **Documentation**: Check all documentation files first
2. **Examples**: Review example code in `examples/` directory
3. **Tests**: Look at test cases for usage patterns
4. **Issues**: Create detailed issue reports with diagnostic information

### Emergency Recovery

If documents are corrupted or lost:

1. **Check Audit Directory**: Look for `before.docx` backup
2. **Word AutoRecover**: Check Word's AutoRecover location
3. **File History**: Use Windows File History if enabled
4. **Version Control**: Restore from git/backup if available

### Performance Monitoring

```python
def monitor_performance():
    """Monitor pipeline performance"""
    import time
    import psutil
    
    start_time = time.time()
    start_memory = psutil.Process().memory_info().rss
    
    # Run pipeline
    result = pipeline.process_document(doc_path, intent)
    
    end_time = time.time()
    end_memory = psutil.Process().memory_info().rss
    
    print(f"Processing time: {end_time - start_time:.2f}s")
    print(f"Memory used: {(end_memory - start_memory) / 1024**2:.1f} MB")
    print(f"Status: {result.status}")
```

This troubleshooting guide covers the most common issues encountered when using AutoWord vNext. For issues not covered here, please collect diagnostic information and consult the support channels.