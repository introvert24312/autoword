#!/usr/bin/env python3
"""
AutoWord vNext Setup Script

This script sets up the complete vNext pipeline for deployment including:
- Configuration files
- Schema validation
- Test data packages
- Documentation
- Dependencies
"""

import os
import sys
import json
import shutil
from pathlib import Path
from typing import Dict, List, Optional

class VNextPackager:
    """Handles packaging and deployment of AutoWord vNext pipeline."""
    
    def __init__(self, target_dir: str = "dist/autoword-vnext"):
        self.target_dir = Path(target_dir)
        self.source_dir = Path(__file__).parent
        
    def create_package(self) -> bool:
        """Create complete deployment package."""
        try:
            print("Creating AutoWord vNext deployment package...")
            
            # Create target directory
            self.target_dir.mkdir(parents=True, exist_ok=True)
            
            # Copy core modules
            self._copy_core_modules()
            
            # Copy configuration files
            self._copy_configurations()
            
            # Copy schemas and documentation
            self._copy_schemas_and_docs()
            
            # Copy test data (optional)
            self._copy_test_data()
            
            # Create deployment scripts
            self._create_deployment_scripts()
            
            # Generate package manifest
            self._generate_manifest()
            
            print(f"Package created successfully at: {self.target_dir}")
            return True
            
        except Exception as e:
            print(f"Error creating package: {e}")
            return False
    
    def _copy_core_modules(self):
        """Copy all core Python modules."""
        print("Copying core modules...")
        
        # Core module directories
        modules = [
            "extractor",
            "planner", 
            "executor",
            "validator",
            "auditor"
        ]
        
        for module in modules:
            src = self.source_dir / module
            dst = self.target_dir / "autoword" / "vnext" / module
            if src.exists():
                shutil.copytree(src, dst, dirs_exist_ok=True)
        
        # Core Python files
        core_files = [
            "__init__.py",
            "__main__.py",
            "cli.py",
            "pipeline.py",
            "models.py",
            "exceptions.py",
            "error_handler.py",
            "localization.py",
            "constraints.py",
            "schema_validator.py"
        ]
        
        for file in core_files:
            src = self.source_dir / file
            dst = self.target_dir / "autoword" / "vnext" / file
            if src.exists():
                dst.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(src, dst)
    
    def _copy_configurations(self):
        """Copy configuration files."""
        print("Copying configuration files...")
        
        config_src = self.source_dir / "config"
        config_dst = self.target_dir / "autoword" / "vnext" / "config"
        
        if config_src.exists():
            shutil.copytree(config_src, config_dst, dirs_exist_ok=True)
    
    def _copy_schemas_and_docs(self):
        """Copy JSON schemas and documentation."""
        print("Copying schemas and documentation...")
        
        # Schemas
        schemas_src = self.source_dir / "schemas"
        schemas_dst = self.target_dir / "autoword" / "vnext" / "schemas"
        
        if schemas_src.exists():
            shutil.copytree(schemas_src, schemas_dst, dirs_exist_ok=True)
        
        # Documentation
        docs = [
            "KNOWN-ISSUES.md",
            "README.md"
        ]
        
        for doc in docs:
            src = self.source_dir / doc
            dst = self.target_dir / doc
            if src.exists():
                shutil.copy2(src, dst)
    
    def _copy_test_data(self):
        """Copy test data packages (optional)."""
        print("Copying test data packages...")
        
        test_data_src = self.source_dir / "test_data"
        test_data_dst = self.target_dir / "test_data"
        
        if test_data_src.exists():
            shutil.copytree(test_data_src, test_data_dst, dirs_exist_ok=True)
    
    def _create_deployment_scripts(self):
        """Create deployment and startup scripts."""
        print("Creating deployment scripts...")
        
        # Windows batch script
        batch_script = self.target_dir / "run_vnext.bat"
        batch_content = """@echo off
echo Starting AutoWord vNext Pipeline...
python -m autoword.vnext %*
pause
"""
        batch_script.write_text(batch_content)
        
        # Unix shell script
        shell_script = self.target_dir / "run_vnext.sh"
        shell_content = """#!/bin/bash
echo "Starting AutoWord vNext Pipeline..."
python -m autoword.vnext "$@"
"""
        shell_script.write_text(shell_content)
        shell_script.chmod(0o755)
        
        # Python entry point
        entry_script = self.target_dir / "run_vnext.py"
        entry_content = """#!/usr/bin/env python3
\"\"\"
AutoWord vNext Entry Point
\"\"\"
import sys
from pathlib import Path

# Add package to path
package_dir = Path(__file__).parent
sys.path.insert(0, str(package_dir))

# Import and run
from autoword.vnext.cli import main

if __name__ == "__main__":
    main()
"""
        entry_script.write_text(entry_content)
    
    def _generate_manifest(self):
        """Generate package manifest with version and contents."""
        print("Generating package manifest...")
        
        manifest = {
            "package": "autoword-vnext",
            "version": "1.0.0",
            "description": "AutoWord vNext Pipeline - Structured Document Processing",
            "author": "AutoWord Development Team",
            "created": "2024-01-01",
            "python_version": f"{sys.version_info.major}.{sys.version_info.minor}",
            "components": {
                "core_modules": [
                    "extractor", "planner", "executor", "validator", "auditor"
                ],
                "configurations": [
                    "localization.json", "pipeline.json", "validation_rules.json"
                ],
                "schemas": [
                    "structure.v1.json", "plan.v1.json", "inventory.full.v1.json"
                ],
                "documentation": [
                    "structure.v1.md", "plan.v1.md", "KNOWN-ISSUES.md"
                ],
                "test_data": [
                    "scenario_1_normal_paper",
                    "scenario_2_no_toc_document", 
                    "scenario_3_duplicate_headings",
                    "scenario_4_headings_in_tables",
                    "scenario_5_missing_fonts",
                    "scenario_6_complex_objects",
                    "scenario_7_revision_tracking"
                ]
            },
            "requirements": [
                "pydantic>=2.0.0",
                "pywin32>=300",
                "jsonschema>=4.0.0"
            ],
            "entry_points": [
                "run_vnext.py",
                "run_vnext.bat", 
                "run_vnext.sh"
            ]
        }
        
        manifest_file = self.target_dir / "manifest.json"
        with open(manifest_file, 'w', encoding='utf-8') as f:
            json.dump(manifest, f, indent=2, ensure_ascii=False)
    
    def verify_package(self) -> bool:
        """Verify package integrity."""
        print("Verifying package integrity...")
        
        required_files = [
            "autoword/vnext/__init__.py",
            "autoword/vnext/pipeline.py",
            "autoword/vnext/config/localization.json",
            "autoword/vnext/schemas/structure.v1.json",
            "autoword/vnext/schemas/plan.v1.json",
            "manifest.json"
        ]
        
        missing_files = []
        for file_path in required_files:
            if not (self.target_dir / file_path).exists():
                missing_files.append(file_path)
        
        if missing_files:
            print(f"Missing required files: {missing_files}")
            return False
        
        print("Package verification successful!")
        return True

def main():
    """Main packaging function."""
    import argparse
    
    parser = argparse.ArgumentParser(description="AutoWord vNext Packager")
    parser.add_argument("--target", default="dist/autoword-vnext", 
                       help="Target directory for package")
    parser.add_argument("--verify-only", action="store_true",
                       help="Only verify existing package")
    
    args = parser.parse_args()
    
    packager = VNextPackager(args.target)
    
    if args.verify_only:
        success = packager.verify_package()
    else:
        success = packager.create_package()
        if success:
            success = packager.verify_package()
    
    sys.exit(0 if success else 1)

if __name__ == "__main__":
    main()