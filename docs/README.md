# AutoWord vNext Documentation

Welcome to the comprehensive documentation for AutoWord vNext, an advanced document processing system that transforms Word document modifications into a predictable, reproducible pipeline.

## Documentation Overview

This documentation suite provides everything you need to understand, use, and maintain AutoWord vNext:

### 📋 [User Guide](USER_GUIDE.md)
**Start here if you're new to AutoWord vNext**
- Quick start and installation
- Basic usage examples
- Configuration options
- Common use cases
- Best practices

### 🏗️ [Technical Architecture](TECHNICAL_ARCHITECTURE.md)
**For developers and system architects**
- System architecture overview
- Module specifications
- Data flow and integration points
- Security model
- Performance considerations

### ⚙️ [Atomic Operations Reference](ATOMIC_OPERATIONS.md)
**Complete reference for all document operations**
- Detailed operation specifications
- Parameter documentation
- Usage examples
- Error handling
- Performance tips

### 🔧 [Troubleshooting Guide](TROUBLESHOOTING.md)
**Solutions for common issues**
- Installation problems
- Document processing errors
- Performance issues
- Localization problems
- Diagnostic tools

### 📚 [API Reference](API_REFERENCE.md)
**Complete API documentation**
- Core pipeline API
- Module interfaces
- Data models and schemas
- Configuration classes
- Exception handling

## Quick Navigation

### Getting Started
1. **Installation**: See [User Guide - Installation](USER_GUIDE.md#installation)
2. **First Steps**: Follow [User Guide - Quick Start](USER_GUIDE.md#quick-start)
3. **Configuration**: Review [User Guide - Configuration](USER_GUIDE.md#configuration)

### Common Tasks
- **Process a Document**: [User Guide - Basic Usage](USER_GUIDE.md#basic-usage)
- **Batch Processing**: [User Guide - Batch Processing](USER_GUIDE.md#batch-processing)
- **Custom Operations**: [Atomic Operations Reference](ATOMIC_OPERATIONS.md)
- **Error Resolution**: [Troubleshooting Guide](TROUBLESHOOTING.md)

### Advanced Topics
- **System Architecture**: [Technical Architecture](TECHNICAL_ARCHITECTURE.md)
- **API Integration**: [API Reference](API_REFERENCE.md)
- **Performance Optimization**: [Technical Architecture - Performance](TECHNICAL_ARCHITECTURE.md#performance-considerations)
- **Security Model**: [Technical Architecture - Security](TECHNICAL_ARCHITECTURE.md#security-model)

## Key Concepts

### The Five-Stage Pipeline

AutoWord vNext processes documents through five distinct stages:

1. **Extract** 📤: Convert DOCX to structured JSON
2. **Plan** 🤖: Generate AI-powered execution plan
3. **Execute** ⚡: Apply atomic operations via Word COM
4. **Validate** ✅: Verify document integrity
5. **Audit** 📊: Create complete audit trail

### Core Principles

- **Zero Information Loss**: Complete document preservation through inventory system
- **Atomic Operations**: Safe, predictable modifications only
- **Complete Auditability**: Full before/after snapshots and diff reports
- **Strict Validation**: JSON schema enforcement and rollback capability

### Supported Operations

- **Section Management**: Delete sections by heading
- **Table of Contents**: Update and manage TOCs
- **Style Management**: Modify and apply document styles
- **Formatting**: Clear direct formatting with authorization
- **Localization**: Chinese/English style and font support

## Document Structure

```
docs/
├── README.md                    # This overview document
├── USER_GUIDE.md               # Complete user guide
├── TECHNICAL_ARCHITECTURE.md   # System architecture
├── ATOMIC_OPERATIONS.md        # Operations reference
├── TROUBLESHOOTING.md          # Problem resolution
├── API_REFERENCE.md            # Complete API docs
└── examples/                   # Usage examples
    ├── basic_usage.py
    ├── batch_processing.py
    ├── custom_configuration.py
    └── error_handling.py
```

## Examples Directory

The `examples/` directory contains practical code examples:

- **basic_usage.py**: Simple document processing
- **batch_processing.py**: Multiple document handling
- **custom_configuration.py**: Advanced configuration
- **error_handling.py**: Robust error management

## Version Information

This documentation covers AutoWord vNext with the following schema versions:
- Structure Schema: `v1`
- Plan Schema: `v1`
- Inventory Schema: `v1`

## Support and Contribution

### Getting Help

1. **Check Documentation**: Start with relevant documentation section
2. **Review Examples**: Look at practical examples in `examples/`
3. **Troubleshooting**: Use the [Troubleshooting Guide](TROUBLESHOOTING.md)
4. **Diagnostic Tools**: Run system diagnostics as described

### Reporting Issues

When reporting issues, please include:
- System information (Python version, Word version, OS)
- Complete error messages and stack traces
- Minimal reproduction case
- Configuration used
- Expected vs actual behavior

### Documentation Updates

This documentation is maintained alongside the codebase. When contributing:
- Update relevant documentation sections
- Add examples for new features
- Update API reference for interface changes
- Test all examples and code snippets

## Changelog

### Version 1.0.0
- Initial release with complete five-stage pipeline
- Full Chinese localization support
- Comprehensive atomic operations
- Complete audit trail system
- JSON schema validation

## License

AutoWord vNext is licensed under [LICENSE]. See the main project repository for complete license information.

---

**Need help?** Start with the [User Guide](USER_GUIDE.md) or check the [Troubleshooting Guide](TROUBLESHOOTING.md) for common issues.