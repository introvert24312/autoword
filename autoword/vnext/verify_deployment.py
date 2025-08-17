#!/usr/bin/env python3
"""
AutoWord vNext Deployment Verification Script

This script verifies that the vNext deployment package is complete and functional.
"""

import os
import sys
import json
import importlib
from pathlib import Path
from typing import List, Dict, Tuple

class DeploymentVerifier:
    """Verifies AutoWord vNext deployment package integrity."""
    
    def __init__(self, package_root: str = None):
        if package_root:
            self.package_root = Path(package_root)
        else:
            self.package_root = Path(__file__).parent
        
        self.errors = []
        self.warnings = []
    
    def verify_all(self) -> bool:
        """Run all verification checks."""
        print("AutoWord vNext Deployment Verification")
        print("=" * 50)
        
        checks = [
            ("File Structure", self.verify_file_structure),
            ("Python Modules", self.verify_python_modules),
            ("Configuration Files", self.verify_configuration_files),
            ("JSON Schemas", self.verify_json_schemas),
            ("Documentation", self.verify_documentation),
            ("Test Data", self.verify_test_data),
            ("Dependencies", self.verify_dependencies)
        ]
        
        all_passed = True
        
        for check_name, check_func in checks:
            print(f"\n{check_name}:")
            print("-" * len(check_name))
            
            try:
                passed = check_func()
                if passed:
                    print("✓ PASSED")
                else:
                    print("✗ FAILED")
                    all_passed = False
            except Exception as e:
                print(f"✗ ERROR: {e}")
                self.errors.append(f"{check_name}: {e}")
                all_passed = False
        
        self._print_summary()
        return all_passed
    
    def verify_file_structure(self) -> bool:
        """Verify required files and directories exist."""
        required_structure = {
            "files": [
                "__init__.py",
                "pipeline.py",
                "cli.py",
                "models.py",
                "exceptions.py",
                "error_handler.py",
                "localization.py",
                "constraints.py",
                "schema_validator.py",
                "setup.py",
                "requirements.txt",
                "README.md",
                "DEPLOYMENT.md",
                "KNOWN-ISSUES.md"
            ],
            "directories": [
                "config",
                "schemas", 
                "extractor",
                "planner",
                "executor",
                "validator",
                "auditor",
                "test_data"
            ]
        }
        
        missing_files = []
        missing_dirs = []
        
        for file_path in required_structure["files"]:
            if not (self.package_root / file_path).exists():
                missing_files.append(file_path)
        
        for dir_path in required_structure["directories"]:
            if not (self.package_root / dir_path).is_dir():
                missing_dirs.append(dir_path)
        
        if missing_files:
            self.errors.append(f"Missing files: {missing_files}")
        
        if missing_dirs:
            self.errors.append(f"Missing directories: {missing_dirs}")
        
        return len(missing_files) == 0 and len(missing_dirs) == 0
    
    def verify_python_modules(self) -> bool:
        """Verify Python modules can be imported."""
        modules_to_test = [
            "models",
            "exceptions", 
            "error_handler",
            "localization",
            "constraints",
            "schema_validator"
        ]
        
        # Add package root to Python path
        sys.path.insert(0, str(self.package_root))
        
        import_errors = []
        
        for module_name in modules_to_test:
            try:
                module_path = f"autoword.vnext.{module_name}"
                importlib.import_module(module_path)
                print(f"  ✓ {module_name}")
            except ImportError as e:
                import_errors.append(f"{module_name}: {e}")
                print(f"  ✗ {module_name}: {e}")
        
        if import_errors:
            self.errors.extend(import_errors)
        
        return len(import_errors) == 0
    
    def verify_configuration_files(self) -> bool:
        """Verify configuration files are valid JSON."""
        config_files = [
            "config/localization.json",
            "config/pipeline.json", 
            "config/validation_rules.json"
        ]
        
        config_errors = []
        
        for config_file in config_files:
            config_path = self.package_root / config_file
            
            if not config_path.exists():
                config_errors.append(f"Missing: {config_file}")
                continue
            
            try:
                with open(config_path, 'r', encoding='utf-8') as f:
                    json.load(f)
                print(f"  ✓ {config_file}")
            except json.JSONDecodeError as e:
                config_errors.append(f"Invalid JSON in {config_file}: {e}")
                print(f"  ✗ {config_file}: Invalid JSON")
            except Exception as e:
                config_errors.append(f"Error reading {config_file}: {e}")
                print(f"  ✗ {config_file}: {e}")
        
        if config_errors:
            self.errors.extend(config_errors)
        
        return len(config_errors) == 0
    
    def verify_json_schemas(self) -> bool:
        """Verify JSON schema files are valid."""
        schema_files = [
            "schemas/structure.v1.json",
            "schemas/plan.v1.json",
            "schemas/inventory.full.v1.json"
        ]
        
        schema_docs = [
            "schemas/structure.v1.md",
            "schemas/plan.v1.md"
        ]
        
        schema_errors = []
        
        # Check schema JSON files
        for schema_file in schema_files:
            schema_path = self.package_root / schema_file
            
            if not schema_path.exists():
                schema_errors.append(f"Missing schema: {schema_file}")
                continue
            
            try:
                with open(schema_path, 'r', encoding='utf-8') as f:
                    schema = json.load(f)
                
                # Basic schema validation
                if "$schema" not in schema:
                    self.warnings.append(f"Schema {schema_file} missing $schema field")
                
                print(f"  ✓ {schema_file}")
            except json.JSONDecodeError as e:
                schema_errors.append(f"Invalid JSON in {schema_file}: {e}")
                print(f"  ✗ {schema_file}: Invalid JSON")
        
        # Check schema documentation
        for doc_file in schema_docs:
            doc_path = self.package_root / doc_file
            
            if not doc_path.exists():
                schema_errors.append(f"Missing documentation: {doc_file}")
            else:
                print(f"  ✓ {doc_file}")
        
        if schema_errors:
            self.errors.extend(schema_errors)
        
        return len(schema_errors) == 0
    
    def verify_documentation(self) -> bool:
        """Verify documentation files exist and are readable."""
        doc_files = [
            "README.md",
            "DEPLOYMENT.md", 
            "KNOWN-ISSUES.md",
            "test_data/README.md"
        ]
        
        doc_errors = []
        
        for doc_file in doc_files:
            doc_path = self.package_root / doc_file
            
            if not doc_path.exists():
                doc_errors.append(f"Missing documentation: {doc_file}")
                continue
            
            try:
                with open(doc_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                if len(content.strip()) == 0:
                    doc_errors.append(f"Empty documentation: {doc_file}")
                else:
                    print(f"  ✓ {doc_file} ({len(content)} chars)")
            except Exception as e:
                doc_errors.append(f"Error reading {doc_file}: {e}")
        
        if doc_errors:
            self.errors.extend(doc_errors)
        
        return len(doc_errors) == 0
    
    def verify_test_data(self) -> bool:
        """Verify test data scenarios are present."""
        expected_scenarios = [
            "scenario_1_normal_paper",
            "scenario_2_no_toc_document",
            "scenario_3_duplicate_headings", 
            "scenario_4_headings_in_tables",
            "scenario_5_missing_fonts",
            "scenario_6_complex_objects",
            "scenario_7_revision_tracking"
        ]
        
        test_data_dir = self.package_root / "test_data"
        
        if not test_data_dir.exists():
            self.errors.append("Missing test_data directory")
            return False
        
        missing_scenarios = []
        
        for scenario in expected_scenarios:
            scenario_dir = test_data_dir / scenario
            if not scenario_dir.exists():
                missing_scenarios.append(scenario)
            else:
                # Check for key files
                key_files = ["user_intent.txt", "expected_plan.v1.json"]
                for key_file in key_files:
                    if (scenario_dir / key_file).exists():
                        print(f"  ✓ {scenario}/{key_file}")
                    else:
                        self.warnings.append(f"Missing {scenario}/{key_file}")
        
        if missing_scenarios:
            self.errors.append(f"Missing test scenarios: {missing_scenarios}")
        
        return len(missing_scenarios) == 0
    
    def verify_dependencies(self) -> bool:
        """Verify Python dependencies can be imported."""
        requirements_file = self.package_root / "requirements.txt"
        
        if not requirements_file.exists():
            self.errors.append("Missing requirements.txt")
            return False
        
        try:
            with open(requirements_file, 'r') as f:
                requirements = f.read()
            
            # Check for key dependencies
            key_deps = ["pydantic", "pywin32", "jsonschema"]
            missing_deps = []
            
            for dep in key_deps:
                if dep not in requirements:
                    missing_deps.append(dep)
                else:
                    print(f"  ✓ {dep} listed in requirements")
            
            if missing_deps:
                self.errors.append(f"Missing key dependencies: {missing_deps}")
            
            return len(missing_deps) == 0
            
        except Exception as e:
            self.errors.append(f"Error reading requirements.txt: {e}")
            return False
    
    def _print_summary(self):
        """Print verification summary."""
        print("\n" + "=" * 50)
        print("VERIFICATION SUMMARY")
        print("=" * 50)
        
        if not self.errors and not self.warnings:
            print("✓ All checks passed! Deployment package is ready.")
        else:
            if self.errors:
                print(f"✗ {len(self.errors)} error(s) found:")
                for error in self.errors:
                    print(f"  - {error}")
            
            if self.warnings:
                print(f"⚠ {len(self.warnings)} warning(s):")
                for warning in self.warnings:
                    print(f"  - {warning}")
        
        print()

def main():
    """Main verification function."""
    import argparse
    
    parser = argparse.ArgumentParser(description="AutoWord vNext Deployment Verifier")
    parser.add_argument("--package-root", help="Root directory of package to verify")
    
    args = parser.parse_args()
    
    verifier = DeploymentVerifier(args.package_root)
    success = verifier.verify_all()
    
    sys.exit(0 if success else 1)

if __name__ == "__main__":
    main()