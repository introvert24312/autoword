# AutoWord vNext Installation Guide

## Overview

This guide provides step-by-step instructions for installing AutoWord vNext on Windows systems. The installation process includes system requirements verification, dependency installation, and configuration setup.

## System Requirements

### Operating System
- **Windows 10** (version 1903 or later) or **Windows 11**
- **64-bit architecture** required

### Microsoft Word
- **Microsoft Word 2016** or later
- **Office 365** or standalone installation
- **COM automation** must be enabled (default)

### Python Environment
- **Python 3.8** or later (recommended: **Python 3.10+**)
- **pip** package manager
- **Virtual environment** support (recommended)

### System Resources
- **Memory**: 4GB RAM minimum (8GB+ recommended for large documents)
- **Storage**: 2GB free space (1GB for installation, 1GB for audit trails)
- **Network**: Internet connection for LLM API access

## Pre-Installation Verification

### 1. Check Windows Version
```cmd
winver
```
Ensure you have Windows 10 (1903+) or Windows 11.

### 2. Verify Microsoft Word Installation
```cmd
# Check if Word is installed
reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office" /s | findstr Word
```

Open Microsoft Word manually to ensure it works correctly.

### 3. Verify Python Installation
```cmd
python --version
pip --version
```

If Python is not installed, download from [python.org](https://www.python.org/downloads/).

### 4. Test Word COM Automation
```python
# Save as test_com.py and run
import win32com.client

try:
    word = win32com.client.Dispatch("Word.Application")
    print("Word COM automation available")
    word.Quit()
except Exception as e:
    print(f"COM automation error: {e}")
```

## Installation Methods

### Method 1: Standard Installation (Recommended)

#### Step 1: Create Virtual Environment
```cmd
# Create virtual environment
python -m venv autoword-vnext-env

# Activate virtual environment
autoword-vnext-env\Scripts\activate

# Verify activation
where python
```

#### Step 2: Install Dependencies
```cmd
# Upgrade pip
python -m pip install --upgrade pip

# Install AutoWord vNext dependencies
pip install -r requirements-deployment.txt
```

#### Step 3: Verify Installation
```cmd
# Check installation
python -m autoword.vnext --version

# Run system status check
python -m autoword.vnext status
```

### Method 2: Development Installation

For developers who want to modify or extend AutoWord vNext:

#### Step 1: Clone Repository
```cmd
git clone [repository-url]
cd autoword-vnext
```

#### Step 2: Create Development Environment
```cmd
# Create virtual environment
python -m venv venv

# Activate environment
venv\Scripts\activate

# Install development dependencies
pip install -r requirements.txt
pip install -r requirements-dev.txt
```

#### Step 3: Install in Development Mode
```cmd
# Install in editable mode
pip install -e .

# Verify installation
python -m autoword.vnext status
```

### Method 3: Deployment Package Installation

For production deployments using pre-built packages:

#### Step 1: Extract Deployment Package
```cmd
# Extract the deployment package
unzip autoword-vnext-deployment-YYYYMMDD.zip
cd autoword-vnext
```

#### Step 2: Run Installation Script
```cmd
# Run the installation script
python install.py

# Or manual installation
pip install -r requirements-deployment.txt
```

#### Step 3: Verify Deployment
```cmd
# Verify deployment integrity
python verify_deployment.py

# Test basic functionality
python run_vnext.py --help
```

## Configuration Setup

### 1. LLM API Configuration

#### OpenAI Configuration
```cmd
# Set OpenAI API key
set OPENAI_API_KEY=your_openai_api_key_here

# Make permanent (add to system environment variables)
setx OPENAI_API_KEY "your_openai_api_key_here"
```

#### Anthropic Configuration
```cmd
# Set Anthropic API key
set ANTHROPIC_API_KEY=your_anthropic_api_key_here

# Make permanent
setx ANTHROPIC_API_KEY "your_anthropic_api_key_here"
```

### 2. Create Configuration File
```cmd
# Generate configuration template
python -m autoword.vnext config create my-config.json

# Edit configuration file
notepad my-config.json
```

Example configuration:
```json
{
  "model": "gpt4",
  "temperature": 0.1,
  "audit_dir": "./audit_trails",
  "verbose": false,
  "monitoring_level": "detailed"
}
```

### 3. Test Configuration
```cmd
# Test with configuration file
python -m autoword.vnext status --config my-config.json
```

## Post-Installation Verification

### 1. System Status Check
```cmd
python -m autoword.vnext status
```

Expected output:
```
AutoWord vNext System Status
============================
✓ Python version: 3.10.x
✓ Dependencies: All required packages installed
✓ Microsoft Word: Available (COM automation working)
✓ LLM API: OpenAI API key configured
✓ Configuration: Default configuration loaded
✓ Audit directory: ./audit_trails (writable)

System ready for document processing.
```

### 2. Test Document Processing
```cmd
# Create test document (or use existing)
echo "Test document content" > test.docx

# Test dry run
python -m autoword.vnext dry-run test.docx "Update formatting" --verbose
```

### 3. Verify Audit Trail Creation
```cmd
# Process test document
python -m autoword.vnext process test.docx "Test processing" --verbose

# Check audit trail
dir audit_trails
```

## Troubleshooting Installation Issues

### Common Issues

#### "Module not found" errors
**Solution**:
```cmd
# Reinstall dependencies
pip uninstall -r requirements-deployment.txt -y
pip install -r requirements-deployment.txt
```

#### "Word COM not available"
**Solutions**:
1. Ensure Word is properly installed
2. Try running Word as administrator once
3. Re-register COM components:
   ```cmd
   # Run as administrator
   regsvr32 /i msword.olb
   ```

#### "Permission denied" errors
**Solutions**:
1. Run command prompt as administrator
2. Check file/directory permissions
3. Ensure antivirus is not blocking files

#### "API key not configured"
**Solution**:
```cmd
# Verify environment variables
echo %OPENAI_API_KEY%
echo %ANTHROPIC_API_KEY%

# Reset if needed
setx OPENAI_API_KEY "your_key_here"
```

### Advanced Troubleshooting

#### Clean Installation
```cmd
# Remove virtual environment
rmdir /s autoword-vnext-env

# Clear pip cache
pip cache purge

# Start fresh installation
python -m venv autoword-vnext-env
autoword-vnext-env\Scripts\activate
pip install -r requirements-deployment.txt
```

#### Dependency Conflicts
```cmd
# Check for conflicts
pip check

# Create requirements freeze
pip freeze > current-requirements.txt

# Compare with expected requirements
fc current-requirements.txt requirements-deployment.txt
```

#### COM Registration Issues
```cmd
# Re-register Word COM components
regsvr32 /u msword.olb
regsvr32 /i msword.olb

# Check COM registration
reg query "HKEY_CLASSES_ROOT\Word.Application"
```

## Environment-Specific Installation

### Corporate/Enterprise Environment

#### Network Restrictions
```cmd
# Install with corporate proxy
pip install -r requirements-deployment.txt --proxy http://proxy.company.com:8080

# Use corporate certificate bundle
pip install -r requirements-deployment.txt --cert /path/to/corporate-cert.pem
```

#### Offline Installation
```cmd
# Download packages on connected machine
pip download -r requirements-deployment.txt -d packages/

# Install on offline machine
pip install --no-index --find-links packages/ -r requirements-deployment.txt
```

### Multi-User Environment

#### System-Wide Installation
```cmd
# Install for all users (run as administrator)
pip install -r requirements-deployment.txt --system

# Create shared configuration
mkdir C:\ProgramData\AutoWordVNext
copy config\*.json C:\ProgramData\AutoWordVNext\
```

#### User-Specific Installation
```cmd
# Install for current user only
pip install -r requirements-deployment.txt --user

# Create user configuration
mkdir %APPDATA%\AutoWordVNext
copy config\*.json %APPDATA%\AutoWordVNext\
```

## Maintenance and Updates

### Regular Maintenance
```cmd
# Update dependencies
pip install -r requirements-deployment.txt --upgrade

# Clean temporary files
python -c "import tempfile, shutil; shutil.rmtree(tempfile.gettempdir(), ignore_errors=True)"

# Verify system status
python -m autoword.vnext status
```

### Version Updates
```cmd
# Check current version
python -m autoword.vnext --version

# Backup configuration
copy my-config.json my-config-backup.json

# Update installation
pip install -r requirements-deployment.txt --upgrade

# Verify update
python -m autoword.vnext status
```

## Uninstallation

### Complete Removal
```cmd
# Deactivate virtual environment
deactivate

# Remove virtual environment
rmdir /s autoword-vnext-env

# Remove configuration files
rmdir /s %APPDATA%\AutoWordVNext

# Remove audit trails (if desired)
rmdir /s audit_trails

# Remove environment variables
reg delete "HKEY_CURRENT_USER\Environment" /v OPENAI_API_KEY /f
reg delete "HKEY_CURRENT_USER\Environment" /v ANTHROPIC_API_KEY /f
```

## Next Steps

After successful installation:

1. **Read the [User Guide](USER_GUIDE.md)** for usage instructions
2. **Review [CLI Usage Guide](../CLI_USAGE_GUIDE.md)** for command-line operations
3. **Check [Configuration Examples](../config/)** for advanced setup
4. **Test with sample documents** from the test_data directory
5. **Set up monitoring and logging** for production use

## Support

If you encounter issues during installation:

1. **Check system status**: `python -m autoword.vnext status`
2. **Review error messages** carefully
3. **Consult [Troubleshooting Guide](TROUBLESHOOTING.md)**
4. **Verify system requirements** are met
5. **Try clean installation** if problems persist

For additional support, provide:
- System status output
- Complete error messages
- Installation method used
- System configuration details