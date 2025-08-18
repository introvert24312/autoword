# AutoWord vNext Deployment Guide

This document provides comprehensive guidance for deploying the AutoWord vNext pipeline in production environments.

## Overview

AutoWord vNext uses a structured deployment system that packages all components into a complete, self-contained distribution. The deployment system includes:

- **Core Pipeline**: All five modules (Extractor, Planner, Executor, Validator, Auditor)
- **Configuration System**: Comprehensive configuration files for all aspects
- **Schema Validation**: JSON schemas with documentation
- **Test Data**: All 7 DoD validation scenarios
- **Documentation**: Complete user and technical documentation
- **Deployment Tools**: Scripts for packaging, verification, and installation

## Deployment Architecture

```
autoword-vnext/
├── autoword/vnext/              # Core pipeline modules
│   ├── extractor/               # DOCX to JSON extraction
│   ├── planner/                 # LLM-based plan generation
│   ├── executor/                # Atomic operation execution
│   ├── validator/               # Document validation
│   ├── auditor/                 # Audit trail generation
│   ├── config/                  # Configuration files
│   └── schemas/                 # JSON schema definitions
├── test_data/                   # DoD validation scenarios
├── docs/                        # Documentation
├── run_vnext.py                 # Cross-platform entry point
├── run_vnext.bat               # Windows launcher
├── run_vnext.sh                # Unix launcher
├── INSTALLATION.md             # Installation guide
├── DEPLOYMENT_MANIFEST.json    # Package manifest
└── requirements-deployment.txt  # Production dependencies
```

## Deployment Process

### 1. Create Deployment Package

Use the enhanced deployment script:

```bash
# Create complete deployment package
python deploy.py --target dist/autoword-vnext

# Create package without test data (smaller)
python deploy.py --target dist/autoword-vnext --no-test-data

# Verify existing package
python deploy.py --verify-only --target dist/autoword-vnext
```

### 2. Verify Package Integrity

Run comprehensive verification:

```bash
# Verify deployment package
python verify_deployment.py --package-dir dist/autoword-vnext

# Quick verification (summary only)
python verify_deployment.py --package-dir dist/autoword-vnext --quiet
```

### 3. Distribution

The deployment process creates:
- **Package Directory**: Complete deployment ready for installation
- **ZIP Archive**: Compressed distribution file with timestamp
- **Manifest**: Detailed package contents and verification info

## Configuration Management

### Configuration Files

| File | Purpose | Customization |
|------|---------|---------------|
| `localization.json` | Style aliases and font fallbacks | Language-specific mappings |
| `pipeline.json` | Pipeline execution settings | Timeouts, directories, features |
| `validation_rules.json` | Document validation rules | Style specs, assertions |
| `deployment.json` | Deployment configuration | System requirements, installation |
| `example-config.json` | Runtime configuration template | User-specific settings |

### Configuration Hierarchy

1. **System Defaults**: Built-in configuration in code
2. **Package Defaults**: Configuration files in package
3. **User Configuration**: Custom configuration files
4. **Runtime Parameters**: Command-line arguments

### Customization Examples

#### Custom Font Fallbacks
```json
{
  "font_fallbacks": {
    "CustomFont": ["CustomFont", "Fallback1", "Fallback2"],
    "楷体": ["楷体", "楷体_GB2312", "STKaiti", "KaiTi"]
  }
}
```

#### Pipeline Timeouts
```json
{
  "pipeline_settings": {
    "max_execution_time_seconds": 600,
    "com_timeout_seconds": 60
  }
}
```

#### Validation Rules
```json
{
  "style_assertions": {
    "Heading 1": {
      "font_east_asian": "CustomFont",
      "font_size_pt": 14,
      "font_bold": true
    }
  }
}
```

## Environment-Specific Deployment

### Development Environment

```bash
# Install with development dependencies
pip install -r requirements-deployment.txt[dev]

# Run with verbose logging
python run_vnext.py --verbose --config dev-config.json document.docx "intent"
```

### Production Environment

```bash
# Install production dependencies only
pip install -r requirements-deployment.txt

# Run with production configuration
python run_vnext.py --config prod-config.json document.docx "intent"
```

### Enterprise Environment

```bash
# Custom configuration for enterprise
python run_vnext.py --config enterprise-config.json \
                    --audit-dir /shared/audit \
                    --log-file /logs/vnext.log \
                    document.docx "intent"
```

## Monitoring and Logging

### Log Files

- **vnext-pipeline.log**: Main pipeline execution log
- **warnings.log**: Warnings and NOOP operations
- **audit_runs/*/result.status.txt**: Per-execution status

### Monitoring Levels

- **Basic**: Essential operations and errors
- **Detailed**: Full operation tracking and performance metrics
- **Debug**: Comprehensive debugging information

### Performance Metrics

- Processing time per document
- Memory usage patterns
- Operation success/failure rates
- Font fallback frequency

## Security Considerations

### Access Control

- **File System**: Secure audit directory permissions
- **COM Automation**: Proper Word application isolation
- **Configuration**: Protect sensitive configuration files

### Data Protection

- **Audit Trails**: Complete before/after snapshots
- **Rollback**: Automatic rollback on failures
- **Validation**: Strict schema validation prevents malicious plans

### Network Security

- **LLM Integration**: Secure API communication
- **No Data Leakage**: Only structure data sent to LLM, never content

## Troubleshooting

### Common Deployment Issues

#### Package Verification Fails
```bash
# Check specific issues
python verify_deployment.py --package-dir dist/autoword-vnext

# Common fixes:
# - Ensure all required files are present
# - Validate JSON configuration files
# - Check Python dependencies
```

#### COM Automation Issues
```bash
# Verify Word installation
# Check COM registration
# Ensure proper permissions
```

#### Font Fallback Problems
```bash
# Install required fonts
# Update fallback chains in localization.json
# Check font availability on target system
```

### Diagnostic Commands

```bash
# Test basic functionality
python run_vnext.py --help

# Test with sample scenario
python run_vnext.py test_data/scenario_1_normal_paper/input.docx \
                    "$(cat test_data/scenario_1_normal_paper/user_intent.txt)"

# Verify configuration
python -c "import json; print(json.load(open('autoword/vnext/config/pipeline.json')))"
```

## Maintenance

### Regular Tasks

1. **Update Dependencies**: Keep Python packages current
2. **Validate Configurations**: Ensure JSON files remain valid
3. **Test Scenarios**: Run DoD validation scenarios regularly
4. **Monitor Logs**: Review audit trails and error patterns
5. **Performance Review**: Analyze processing metrics

### Version Updates

1. **Backup Current**: Save current configuration and customizations
2. **Deploy New Version**: Use deployment scripts
3. **Migrate Configuration**: Update configuration files as needed
4. **Test Thoroughly**: Run all validation scenarios
5. **Monitor Rollout**: Watch for issues in production

## Support and Documentation

### Documentation Files

- **INSTALLATION.md**: Step-by-step installation guide
- **KNOWN-ISSUES.md**: Known limitations and workarounds
- **schemas/*.md**: JSON schema documentation
- **test_data/README.md**: Test scenario documentation

### Getting Help

1. Check KNOWN-ISSUES.md for documented problems
2. Review audit trails and log files
3. Test with provided scenarios
4. Verify configuration files
5. Check system requirements

## Deployment Checklist

### Pre-Deployment

- [ ] System requirements verified
- [ ] Python 3.8+ installed
- [ ] Microsoft Word installed and tested
- [ ] Dependencies installed
- [ ] Configuration customized

### Deployment

- [ ] Package created with deploy.py
- [ ] Package verified with verify_deployment.py
- [ ] Distribution archive created
- [ ] Installation guide reviewed
- [ ] Test scenarios validated

### Post-Deployment

- [ ] Basic functionality tested
- [ ] Sample documents processed
- [ ] Monitoring configured
- [ ] Audit trails verified
- [ ] Performance baseline established

### Production Readiness

- [ ] All DoD scenarios pass
- [ ] Performance meets requirements
- [ ] Error handling tested
- [ ] Rollback procedures verified
- [ ] Documentation complete

This deployment guide ensures a robust, maintainable AutoWord vNext installation that meets enterprise requirements while providing comprehensive monitoring and troubleshooting capabilities.