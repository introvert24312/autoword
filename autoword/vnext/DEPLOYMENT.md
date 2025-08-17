# AutoWord vNext Deployment Guide

## Overview

This guide covers the complete deployment process for AutoWord vNext pipeline, including system requirements, installation steps, configuration, and verification procedures.

## System Requirements

### Operating System
- **Windows 10/11** (required for Word COM automation)
- **Windows Server 2016+** (for server deployments)

### Software Dependencies
- **Microsoft Word 2016+** (required for COM automation)
- **Python 3.8+** (recommended: Python 3.10+)
- **pip** (Python package manager)

### Hardware Requirements
- **RAM**: Minimum 4GB, recommended 8GB+
- **Storage**: 500MB for installation, additional space for audit trails
- **CPU**: Modern multi-core processor recommended for large documents

## Installation Methods

### Method 1: Package Installation (Recommended)

1. **Download Package**
   ```bash
   # Extract the autoword-vnext package
   unzip autoword-vnext-1.0.0.zip
   cd autoword-vnext
   ```

2. **Install Dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Verify Installation**
   ```bash
   python run_vnext.py --version
   ```

### Method 2: Source Installation

1. **Clone Repository**
   ```bash
   git clone <repository-url>
   cd autoword/vnext
   ```

2. **Create Package**
   ```bash
   python setup.py --target dist/vnext-deployment
   ```

3. **Install Package**
   ```bash
   cd dist/vnext-deployment
   pip install -r requirements.txt
   ```

## Configuration

### 1. Localization Settings

Edit `autoword/vnext/config/localization.json`:

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

### 2. Pipeline Settings

Edit `autoword/vnext/config/pipeline.json`:

```json
{
  "pipeline_settings": {
    "max_execution_time_seconds": 300,
    "enable_rollback": true
  }
}
```

### 3. Validation Rules

Edit `autoword/vnext/config/validation_rules.json`:

```json
{
  "chapter_assertions": {
    "forbidden_level_1_headings": ["摘要", "参考文献"]
  }
}
```

## Deployment Scenarios

### Scenario 1: Single User Desktop

**Use Case**: Individual user processing documents locally

**Setup**:
1. Install on user's Windows machine
2. Configure for local Word installation
3. Set up personal audit directory

**Command**:
```bash
python run_vnext.py process document.docx "删除摘要章节"
```

### Scenario 2: Server Deployment

**Use Case**: Centralized document processing service

**Setup**:
1. Install on Windows Server with Word
2. Configure service account with Word permissions
3. Set up network audit storage
4. Implement queue-based processing

**Considerations**:
- Word COM requires interactive desktop session
- Service account needs Word licensing
- Network storage for audit trails

### Scenario 3: Batch Processing

**Use Case**: Processing multiple documents automatically

**Setup**:
1. Create batch processing scripts
2. Configure input/output directories
3. Set up error handling and logging

**Example Script**:
```python
from autoword.vnext.pipeline import VNextPipeline

pipeline = VNextPipeline()
for docx_file in input_directory.glob("*.docx"):
    result = pipeline.process_document(docx_file, user_intent)
    if result.status != "SUCCESS":
        log_error(docx_file, result.error)
```

## Verification

### 1. System Verification

```bash
# Check Python version
python --version

# Check Word COM availability
python -c "import win32com.client; word = win32com.client.Dispatch('Word.Application'); print('Word COM available')"

# Check package integrity
python run_vnext.py --verify
```

### 2. Functional Testing

```bash
# Run test scenarios
python run_vnext.py test --scenario all

# Process sample document
python run_vnext.py process test_data/scenario_1_normal_paper/input.docx "删除摘要章节"
```

### 3. Performance Testing

```bash
# Test with large document
python run_vnext.py process large_document.docx "标准化格式" --benchmark

# Memory usage monitoring
python -m memory_profiler run_vnext.py process document.docx "用户意图"
```

## Security Considerations

### 1. File System Permissions

- **Input Directory**: Read access for service account
- **Output Directory**: Write access for processed documents
- **Audit Directory**: Write access for audit trails
- **Temp Directory**: Full access for temporary files

### 2. Word COM Security

- **Macro Security**: Disable macros in Word settings
- **Protected View**: Configure for automated processing
- **File Blocking**: Allow DOCX files, block potentially dangerous formats

### 3. Network Security

- **Firewall**: Allow Word COM communication
- **Antivirus**: Exclude processing directories from real-time scanning
- **Access Control**: Restrict access to processing service

## Monitoring and Maintenance

### 1. Log Monitoring

Monitor these log files:
- `audit_runs/*/warnings.log` - Processing warnings
- `audit_runs/*/result.status.txt` - Processing results
- System event logs for COM errors

### 2. Performance Monitoring

Track these metrics:
- Processing time per document
- Memory usage during processing
- Success/failure rates
- Audit trail storage usage

### 3. Maintenance Tasks

**Daily**:
- Check processing logs for errors
- Verify audit trail storage space
- Monitor system resource usage

**Weekly**:
- Clean up old audit trails (if configured)
- Update font availability checks
- Review processing statistics

**Monthly**:
- Update configuration files if needed
- Review and update validation rules
- Performance optimization review

## Troubleshooting

### Common Issues

**Issue**: "Word COM not available"
**Solution**: 
- Install Microsoft Word
- Register Word COM components: `regsvr32 "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE" /regserver`

**Issue**: "Font not found" warnings
**Solution**:
- Install required fonts
- Update font fallback configuration
- Check font licensing

**Issue**: "Processing timeout"
**Solution**:
- Increase timeout in pipeline.json
- Check document complexity
- Verify system resources

**Issue**: "Validation failed"
**Solution**:
- Review validation assertions
- Check document structure
- Verify style specifications

### Debug Mode

Enable debug logging:
```bash
python run_vnext.py process document.docx "用户意图" --debug --verbose
```

### Support Information

For technical support:
1. Collect system information
2. Gather processing logs
3. Provide sample document (if possible)
4. Include error messages and stack traces

## Upgrade Procedures

### Version Updates

1. **Backup Current Installation**
   ```bash
   cp -r autoword-vnext autoword-vnext-backup
   ```

2. **Install New Version**
   ```bash
   # Extract new package
   # Copy configuration files from backup
   # Test with sample documents
   ```

3. **Verify Upgrade**
   ```bash
   python run_vnext.py --version
   python run_vnext.py test --scenario basic
   ```

### Configuration Migration

When upgrading, preserve these files:
- `config/localization.json`
- `config/pipeline.json` 
- `config/validation_rules.json`
- Custom test scenarios

## Performance Optimization

### System Optimization

1. **Disable unnecessary Word add-ins**
2. **Increase virtual memory if processing large documents**
3. **Use SSD storage for temp and audit directories**
4. **Close other applications during batch processing**

### Configuration Optimization

1. **Adjust timeout values based on document complexity**
2. **Configure appropriate audit retention policies**
3. **Optimize validation rules for specific document types**
4. **Use efficient font fallback chains**

This deployment guide ensures successful installation and operation of AutoWord vNext in various environments.