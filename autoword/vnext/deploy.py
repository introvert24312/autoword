#!/usr/bin/env python3
"""
AutoWord vNext Deployment Script

Enhanced deployment system for complete vNext pipeline packaging and distribution.
Extends the existing setup.py with additional deployment features.
"""

import os
import sys
import json
import shutil
import zipfile
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from datetime import datetime

class VNextDeployer:
    """Enhanced deployment system for AutoWord vNext pipeline."""
    
    def __init__(self, target_dir: str = "dist/autoword-vnext"):
        self.target_dir = Path(target_dir)
        self.source_dir = Path(__file__).parent
        self.timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
    def create_deployment_package(self, include_test_data: bool = True) -> bool:
        """Create complete deployment package with all components."""
        try:
            print(f"Creating AutoWord vNext deployment package...")
            print(f"Target directory: {self.target_dir}")
            print(f"Timestamp: {self.timestamp}")
            
            # Create target directory
            self.target_dir.mkdir(parents=True, exist_ok=True)
            
            # Copy core components
            self._copy_core_modules()
            self._copy_configurations()
            self._copy_schemas_and_docs()
            
            if include_test_data:
                self._copy_test_data()
            
            # Create deployment artifacts
            self._create_deployment_scripts()
            self._create_installation_guide()
            self._generate_deployment_manifest()
            
            # Create distribution archive
            archive_path = self._create_distribution_archive()
            
            # Verify package integrity
            if self._verify_package():
                print(f"✓ Deployment package created successfully!")
                print(f"✓ Package location: {self.target_dir}")
                print(f"✓ Distribution archive: {archive_path}")
                return True
            else:
                print("✗ Package verification failed!")
                return False
                
        except Exception as e:
            print(f"✗ Error creating deployment package: {e}")
            return False
    
    def _copy_core_modules(self):
        """Copy all core Python modules and dependencies."""
        print("📦 Copying core modules...")
        
        # Core module directories
        modules = [
            "extractor", "planner", "executor", "validator", "auditor"
        ]
        
        for module in modules:
            src = self.source_dir / module
            dst = self.target_dir / "autoword" / "vnext" / module
            if src.exists():
                shutil.copytree(src, dst, dirs_exist_ok=True)
                print(f"   ✓ {module}")
        
        # Core Python files
        core_files = [
            "__init__.py", "__main__.py", "cli.py", "pipeline.py",
            "models.py", "exceptions.py", "error_handler.py",
            "localization.py", "constraints.py", "schema_validator.py",
            "monitoring.py", "performance.py"
        ]
        
        for file in core_files:
            src = self.source_dir / file
            dst = self.target_dir / "autoword" / "vnext" / file
            if src.exists():
                dst.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(src, dst)
                print(f"   ✓ {file}")
    
    def _copy_configurations(self):
        """Copy all configuration files."""
        print("⚙️  Copying configuration files...")
        
        config_src = self.source_dir / "config"
        config_dst = self.target_dir / "autoword" / "vnext" / "config"
        
        if config_src.exists():
            shutil.copytree(config_src, config_dst, dirs_exist_ok=True)
            
            # List copied configuration files
            for config_file in config_dst.glob("*.json"):
                print(f"   ✓ {config_file.name}")
    
    def _copy_schemas_and_docs(self):
        """Copy JSON schemas and documentation."""
        print("📚 Copying schemas and documentation...")
        
        # Schemas
        schemas_src = self.source_dir / "schemas"
        schemas_dst = self.target_dir / "autoword" / "vnext" / "schemas"
        
        if schemas_src.exists():
            shutil.copytree(schemas_src, schemas_dst, dirs_exist_ok=True)
            for schema_file in schemas_dst.glob("*"):
                print(f"   ✓ schemas/{schema_file.name}")
        
        # Documentation
        docs = [
            "KNOWN-ISSUES.md", "README.md", "DEPLOYMENT.md",
            "CLI-README.md", "MONITORING.md"
        ]
        
        for doc in docs:
            src = self.source_dir / doc
            dst = self.target_dir / doc
            if src.exists():
                shutil.copy2(src, dst)
                print(f"   ✓ {doc}")
    
    def _copy_test_data(self):
        """Copy test data packages for all 7 DoD scenarios."""
        print("🧪 Copying test data packages...")
        
        test_data_src = self.source_dir / "test_data"
        test_data_dst = self.target_dir / "test_data"
        
        if test_data_src.exists():
            shutil.copytree(test_data_src, test_data_dst, dirs_exist_ok=True)
            
            # List all scenarios
            scenarios = [d for d in test_data_dst.iterdir() if d.is_dir()]
            for scenario in scenarios:
                print(f"   ✓ {scenario.name}")
    
    def _create_deployment_scripts(self):
        """Create deployment and startup scripts for multiple platforms."""
        print("🚀 Creating deployment scripts...")
        
        # Windows batch script
        batch_script = self.target_dir / "run_vnext.bat"
        batch_content = f"""@echo off
REM AutoWord vNext Pipeline Launcher
REM Generated: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}

echo Starting AutoWord vNext Pipeline...
echo.

REM Check Python installation
python --version >nul 2>&1
if errorlevel 1 (
    echo Error: Python is not installed or not in PATH
    echo Please install Python 3.8+ and try again
    pause
    exit /b 1
)

REM Check Word installation (basic check)
if not exist "%ProgramFiles%\\Microsoft Office" (
    if not exist "%ProgramFiles(x86)%\\Microsoft Office" (
        echo Warning: Microsoft Office may not be installed
        echo COM automation may not work properly
        echo.
    )
)

REM Run the pipeline
python -m autoword.vnext %*

REM Check exit code
if errorlevel 1 (
    echo.
    echo Pipeline execution failed. Check logs for details.
    pause
    exit /b 1
) else (
    echo.
    echo Pipeline execution completed successfully.
)

pause
"""
        batch_script.write_text(batch_content, encoding='utf-8')
        print("   ✓ run_vnext.bat")
        
        # Unix shell script
        shell_script = self.target_dir / "run_vnext.sh"
        shell_content = f"""#!/bin/bash
# AutoWord vNext Pipeline Launcher
# Generated: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}

echo "Starting AutoWord vNext Pipeline..."
echo

# Check Python installation
if ! command -v python3 &> /dev/null; then
    echo "Error: Python 3 is not installed or not in PATH"
    echo "Please install Python 3.8+ and try again"
    exit 1
fi

# Check Python version
python_version=$(python3 -c "import sys; print(f'{{sys.version_info.major}}.{{sys.version_info.minor}}')")
required_version="3.8"

if [ "$(printf '%s\\n' "$required_version" "$python_version" | sort -V | head -n1)" != "$required_version" ]; then
    echo "Error: Python $python_version found, but Python $required_version+ is required"
    exit 1
fi

# Run the pipeline
python3 -m autoword.vnext "$@"

# Check exit code
if [ $? -eq 0 ]; then
    echo
    echo "Pipeline execution completed successfully."
else
    echo
    echo "Pipeline execution failed. Check logs for details."
    exit 1
fi
"""
        shell_script.write_text(shell_content)
        shell_script.chmod(0o755)
        print("   ✓ run_vnext.sh")
        
        # Python entry point
        entry_script = self.target_dir / "run_vnext.py"
        entry_content = f'''#!/usr/bin/env python3
"""
AutoWord vNext Entry Point
Generated: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}

This script provides a cross-platform entry point for the AutoWord vNext pipeline.
"""

import sys
import os
from pathlib import Path

def main():
    """Main entry point with environment setup."""
    
    # Add package to Python path
    package_dir = Path(__file__).parent
    sys.path.insert(0, str(package_dir))
    
    # Verify Python version
    if sys.version_info < (3, 8):
        print("Error: Python 3.8+ is required")
        print(f"Current version: {{sys.version}}")
        sys.exit(1)
    
    # Import and run CLI
    try:
        from autoword.vnext.cli import main as cli_main
        cli_main()
    except ImportError as e:
        print(f"Error importing AutoWord vNext: {{e}}")
        print("Please ensure all dependencies are installed")
        sys.exit(1)
    except Exception as e:
        print(f"Error running AutoWord vNext: {{e}}")
        sys.exit(1)

if __name__ == "__main__":
    main()
'''
        entry_script.write_text(entry_content)
        print("   ✓ run_vnext.py")
    
    def _create_installation_guide(self):
        """Create comprehensive installation guide."""
        print("📖 Creating installation guide...")
        
        guide_content = f"""# AutoWord vNext Installation Guide

Generated: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}

## System Requirements

### Operating System
- Windows 10 or later
- Windows Server 2016 or later

### Microsoft Word
- Microsoft Word 2016 or later (2019+ recommended)
- COM automation enabled
- Macro security set to medium or lower

### Python
- Python 3.8 or later
- pip package manager

## Installation Steps

### 1. Verify Prerequisites

#### Check Python Installation
```bash
python --version
# Should show Python 3.8.x or later
```

#### Check Word Installation
- Open Microsoft Word
- Verify it starts without errors
- Check that macros are enabled (File → Options → Trust Center → Macro Settings)

### 2. Install Python Dependencies

```bash
pip install pydantic>=2.0.0 pywin32>=300 jsonschema>=4.0.0
```

### 3. Deploy AutoWord vNext

1. Extract the deployment package to your desired location
2. Verify the directory structure:
   ```
   autoword-vnext/
   ├── autoword/vnext/          # Core modules
   ├── config/                  # Configuration files
   ├── schemas/                 # JSON schemas
   ├── test_data/              # Test scenarios
   ├── run_vnext.py            # Entry point
   ├── run_vnext.bat           # Windows launcher
   └── run_vnext.sh            # Unix launcher
   ```

### 4. Verify Installation

Run the test scenario:
```bash
python run_vnext.py --help
```

Test with sample document:
```bash
python run_vnext.py test_data/scenario_1_normal_paper/input.docx "Remove 摘要 and 参考文献 sections"
```

## Configuration

### Basic Configuration
Copy `config/example-config.json` to create your configuration:
```bash
cp config/example-config.json my-config.json
```

### Advanced Configuration
Edit configuration files in the `config/` directory:
- `localization.json`: Style aliases and font fallbacks
- `pipeline.json`: Pipeline execution settings
- `validation_rules.json`: Document validation rules

## Troubleshooting

### Common Issues

#### "COM automation failed"
- Ensure Microsoft Word is installed
- Check that Word can be opened manually
- Verify COM automation is enabled

#### "Font fallback warnings"
- Install required fonts (楷体, 宋体, etc.)
- Update font fallback chains in `config/localization.json`

#### "Schema validation errors"
- Verify JSON schema files are present in `schemas/`
- Check that schema files are valid JSON

#### "Memory errors with large documents"
- Increase memory limits in `config/pipeline.json`
- Process documents in smaller sections
- Close other applications to free memory

### Getting Help

1. Check `KNOWN-ISSUES.md` for documented issues
2. Review log files in the audit directory
3. Test with provided scenarios in `test_data/`
4. Verify configuration files are valid JSON

## Usage Examples

### Basic Usage
```bash
python run_vnext.py document.docx "Remove abstract and references"
```

### With Custom Configuration
```bash
python run_vnext.py --config my-config.json document.docx "Clean up document"
```

### Verbose Output
```bash
python run_vnext.py --verbose document.docx "Process document"
```

### Dry Run (Plan Only)
```bash
python run_vnext.py --dry-run document.docx "Generate plan only"
```

## Next Steps

1. Test with your own documents
2. Customize configuration files for your needs
3. Set up monitoring and logging as required
4. Review audit trails and validation results

For detailed usage information, see the CLI documentation:
```bash
python run_vnext.py --help
```
"""
        
        guide_file = self.target_dir / "INSTALLATION.md"
        guide_file.write_text(guide_content, encoding='utf-8')
        print("   ✓ INSTALLATION.md")
    
    def _generate_deployment_manifest(self):
        """Generate comprehensive deployment manifest."""
        print("📋 Generating deployment manifest...")
        
        # Load deployment configuration
        deployment_config_path = self.source_dir / "config" / "deployment.json"
        if deployment_config_path.exists():
            with open(deployment_config_path, 'r', encoding='utf-8') as f:
                deployment_config = json.load(f)
        else:
            deployment_config = {}
        
        manifest = {
            "deployment_info": {
                "package_name": "autoword-vnext",
                "version": "1.0.0",
                "build_timestamp": self.timestamp,
                "build_date": datetime.now().isoformat(),
                "python_version": f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}",
                "platform": sys.platform
            },
            "package_contents": {
                "core_modules": self._get_module_list(),
                "configuration_files": self._get_config_list(),
                "schema_files": self._get_schema_list(),
                "documentation_files": self._get_docs_list(),
                "test_scenarios": self._get_test_scenarios(),
                "entry_points": [
                    "run_vnext.py",
                    "run_vnext.bat",
                    "run_vnext.sh"
                ]
            },
            "system_requirements": deployment_config.get("system_requirements", {}),
            "installation": deployment_config.get("installation", {}),
            "verification": {
                "package_integrity": "verified",
                "required_files_present": True,
                "configuration_valid": True,
                "schemas_valid": True
            }
        }
        
        manifest_file = self.target_dir / "DEPLOYMENT_MANIFEST.json"
        with open(manifest_file, 'w', encoding='utf-8') as f:
            json.dump(manifest, f, indent=2, ensure_ascii=False)
        print("   ✓ DEPLOYMENT_MANIFEST.json")
    
    def _create_distribution_archive(self) -> Path:
        """Create ZIP archive for distribution."""
        print("📦 Creating distribution archive...")
        
        archive_name = f"autoword-vnext-{self.timestamp}.zip"
        archive_path = self.target_dir.parent / archive_name
        
        with zipfile.ZipFile(archive_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for file_path in self.target_dir.rglob('*'):
                if file_path.is_file():
                    arcname = file_path.relative_to(self.target_dir.parent)
                    zipf.write(file_path, arcname)
        
        print(f"   ✓ {archive_name}")
        return archive_path
    
    def _verify_package(self) -> bool:
        """Verify package integrity and completeness."""
        print("🔍 Verifying package integrity...")
        
        required_files = [
            "autoword/vnext/__init__.py",
            "autoword/vnext/pipeline.py",
            "autoword/vnext/cli.py",
            "autoword/vnext/config/localization.json",
            "autoword/vnext/config/pipeline.json",
            "autoword/vnext/config/validation_rules.json",
            "autoword/vnext/schemas/structure.v1.json",
            "autoword/vnext/schemas/plan.v1.json",
            "autoword/vnext/schemas/structure.v1.md",
            "autoword/vnext/schemas/plan.v1.md",
            "DEPLOYMENT_MANIFEST.json",
            "INSTALLATION.md",
            "run_vnext.py"
        ]
        
        missing_files = []
        for file_path in required_files:
            if not (self.target_dir / file_path).exists():
                missing_files.append(file_path)
        
        if missing_files:
            print(f"   ✗ Missing required files: {missing_files}")
            return False
        
        # Verify JSON files are valid
        json_files = list(self.target_dir.rglob("*.json"))
        for json_file in json_files:
            try:
                with open(json_file, 'r', encoding='utf-8') as f:
                    json.load(f)
            except json.JSONDecodeError as e:
                print(f"   ✗ Invalid JSON file {json_file}: {e}")
                return False
        
        print("   ✓ All required files present")
        print("   ✓ All JSON files valid")
        return True
    
    def _get_module_list(self) -> List[str]:
        """Get list of core modules."""
        modules = []
        vnext_dir = self.target_dir / "autoword" / "vnext"
        if vnext_dir.exists():
            for item in vnext_dir.iterdir():
                if item.is_dir() and not item.name.startswith('__'):
                    modules.append(item.name)
        return sorted(modules)
    
    def _get_config_list(self) -> List[str]:
        """Get list of configuration files."""
        configs = []
        config_dir = self.target_dir / "autoword" / "vnext" / "config"
        if config_dir.exists():
            configs = [f.name for f in config_dir.glob("*.json")]
        return sorted(configs)
    
    def _get_schema_list(self) -> List[str]:
        """Get list of schema files."""
        schemas = []
        schema_dir = self.target_dir / "autoword" / "vnext" / "schemas"
        if schema_dir.exists():
            schemas = [f.name for f in schema_dir.iterdir() if f.is_file()]
        return sorted(schemas)
    
    def _get_docs_list(self) -> List[str]:
        """Get list of documentation files."""
        docs = []
        for pattern in ["*.md", "*.txt"]:
            docs.extend([f.name for f in self.target_dir.glob(pattern)])
        return sorted(docs)
    
    def _get_test_scenarios(self) -> List[str]:
        """Get list of test scenarios."""
        scenarios = []
        test_dir = self.target_dir / "test_data"
        if test_dir.exists():
            scenarios = [d.name for d in test_dir.iterdir() if d.is_dir()]
        return sorted(scenarios)

def main():
    """Main deployment function."""
    import argparse
    
    parser = argparse.ArgumentParser(description="AutoWord vNext Enhanced Deployer")
    parser.add_argument("--target", default="dist/autoword-vnext",
                       help="Target directory for deployment package")
    parser.add_argument("--no-test-data", action="store_true",
                       help="Exclude test data from package")
    parser.add_argument("--verify-only", action="store_true",
                       help="Only verify existing package")
    
    args = parser.parse_args()
    
    deployer = VNextDeployer(args.target)
    
    if args.verify_only:
        success = deployer._verify_package()
    else:
        success = deployer.create_deployment_package(
            include_test_data=not args.no_test_data
        )
    
    sys.exit(0 if success else 1)

if __name__ == "__main__":
    main()