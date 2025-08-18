#!/usr/bin/env python3
"""
AutoWord vNext Deployment Verification Script

Comprehensive verification of deployment package integrity and functionality.
"""

import os
import sys
import json
import subprocess
from pathlib import Path
from typing import Dict, List, Tuple, Optional

class DeploymentVerifier:
    """Verifies AutoWord vNext deployment package."""
    
    def __init__(self, package_dir: str = "."):
        self.package_dir = Path(package_dir).resolve()
        self.errors = []
        self.warnings = []
        
    def verify_deployment(self) -> bool:
        """Run complete deployment verification."""
        print("🔍 AutoWord vNext Deployment Verification")
        print("=" * 50)
        
        # Run all verification checks
        checks = [
            ("System Requirements", self._check_system_requirements),
            ("Package Structure", self._check_package_structure),
            ("Configuration Files", self._check_configuration_files),
            ("Schema Files", self._check_schema_files),
            ("Python Dependencies", self._check_python_dependencies),
            ("Entry Points", self._check_entry_points),
            ("Test Data", self._check_test_data),
            ("Documentation", self._check_documentation)
        ]
        
        all_passed = True
        for check_name, check_func in checks:
            print(f"\n📋 {check_name}")
            print("-" * 30)
            
            try:
                passed = check_func()
                if passed:
                    print(f"✅ {check_name}: PASSED")
                else:
                    print(f"❌ {check_name}: FAILED")
                    all_passed = False
            except Exception as e:
                print(f"💥 {check_name}: ERROR - {e}")
                self.errors.append(f"{check_name}: {e}")
                all_passed = False
        
        # Print summary
        self._print_summary(all_passed)
        return all_passed
    
    def _check_system_requirements(self) -> bool:
        """Check system requirements."""
        passed = True
        
        # Check Python version
        if sys.version_info < (3, 8):
            self.errors.append(f"Python 3.8+ required, found {sys.version}")
            passed = False
        else:
            print(f"✓ Python {sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}")
        
        # Check platform
        if sys.platform != "win32":
            self.warnings.append("AutoWord vNext is designed for Windows with Microsoft Word")
            print(f"⚠️  Platform: {sys.platform} (Windows recommended)")
        else:
            print(f"✓ Platform: {sys.platform}")
        
        # Check Word installation (basic check)
        word_paths = [
            Path(os.environ.get("ProgramFiles", "")) / "Microsoft Office",
            Path(os.environ.get("ProgramFiles(x86)", "")) / "Microsoft Office"
        ]
        
        word_found = any(path.exists() for path in word_paths)
        if not word_found:
            self.warnings.append("Microsoft Office installation not detected")
            print("⚠️  Microsoft Office not detected (may affect COM automation)")
        else:
            print("✓ Microsoft Office installation detected")
        
        return passed
    
    def _check_package_structure(self) -> bool:
        """Check package directory structure."""
        required_structure = {
            "autoword/vnext/__init__.py": "Core package init",
            "autoword/vnext/pipeline.py": "Main pipeline module",
            "autoword/vnext/cli.py": "Command-line interface",
            "autoword/vnext/extractor/": "Extractor module",
            "autoword/vnext/planner/": "Planner module", 
            "autoword/vnext/executor/": "Executor module",
            "autoword/vnext/validator/": "Validator module",
            "autoword/vnext/auditor/": "Auditor module",
            "autoword/vnext/config/": "Configuration directory",
            "autoword/vnext/schemas/": "Schema directory"
        }
        
        passed = True
        for path, description in required_structure.items():
            full_path = self.package_dir / path
            if full_path.exists():
                print(f"✓ {path}")
            else:
                print(f"✗ {path} - {description}")
                self.errors.append(f"Missing: {path}")
                passed = False
        
        return passed
    
    def _check_configuration_files(self) -> bool:
        """Check configuration files."""
        config_dir = self.package_dir / "autoword" / "vnext" / "config"
        
        required_configs = [
            "localization.json",
            "pipeline.json", 
            "validation_rules.json",
            "deployment.json"
        ]
        
        passed = True
        for config_file in required_configs:
            config_path = config_dir / config_file
            if not config_path.exists():
                print(f"✗ {config_file} - Missing")
                self.errors.append(f"Missing config: {config_file}")
                passed = False
                continue
            
            # Validate JSON
            try:
                with open(config_path, 'r', encoding='utf-8') as f:
                    json.load(f)
                print(f"✓ {config_file} - Valid JSON")
            except json.JSONDecodeError as e:
                print(f"✗ {config_file} - Invalid JSON: {e}")
                self.errors.append(f"Invalid JSON in {config_file}: {e}")
                passed = False
        
        return passed
    
    def _check_schema_files(self) -> bool:
        """Check JSON schema files."""
        schema_dir = self.package_dir / "autoword" / "vnext" / "schemas"
        
        required_schemas = [
            ("structure.v1.json", "structure.v1.md"),
            ("plan.v1.json", "plan.v1.md"),
            ("inventory.full.v1.json", None)
        ]
        
        passed = True
        for schema_file, doc_file in required_schemas:
            schema_path = schema_dir / schema_file
            
            # Check schema file
            if not schema_path.exists():
                print(f"✗ {schema_file} - Missing")
                self.errors.append(f"Missing schema: {schema_file}")
                passed = False
                continue
            
            # Validate JSON schema
            try:
                with open(schema_path, 'r', encoding='utf-8') as f:
                    schema_data = json.load(f)
                print(f"✓ {schema_file} - Valid JSON schema")
                
                # Check for required schema properties
                if "$schema" not in schema_data:
                    self.warnings.append(f"{schema_file} missing $schema property")
                    
            except json.JSONDecodeError as e:
                print(f"✗ {schema_file} - Invalid JSON: {e}")
                self.errors.append(f"Invalid JSON in {schema_file}: {e}")
                passed = False
            
            # Check documentation file
            if doc_file:
                doc_path = schema_dir / doc_file
                if doc_path.exists():
                    print(f"✓ {doc_file} - Documentation present")
                else:
                    self.warnings.append(f"Missing documentation: {doc_file}")
                    print(f"⚠️  {doc_file} - Documentation missing")
        
        return passed
    
    def _check_python_dependencies(self) -> bool:
        """Check Python dependencies."""
        required_packages = [
            ("pydantic", "2.0.0"),
            ("pywin32", "300"),
            ("jsonschema", "4.0.0")
        ]
        
        passed = True
        for package, min_version in required_packages:
            try:
                __import__(package)
                print(f"✓ {package} - Available")
            except ImportError:
                print(f"✗ {package} - Not installed")
                self.errors.append(f"Missing dependency: {package}>={min_version}")
                passed = False
        
        return passed
    
    def _check_entry_points(self) -> bool:
        """Check entry point scripts."""
        entry_points = [
            "run_vnext.py",
            "run_vnext.bat",
            "run_vnext.sh"
        ]
        
        passed = True
        for entry_point in entry_points:
            entry_path = self.package_dir / entry_point
            if entry_path.exists():
                print(f"✓ {entry_point}")
                
                # Check if executable (Unix)
                if entry_point.endswith('.sh') and not os.access(entry_path, os.X_OK):
                    self.warnings.append(f"{entry_point} not executable")
                    print(f"⚠️  {entry_point} - Not executable")
            else:
                print(f"✗ {entry_point} - Missing")
                self.errors.append(f"Missing entry point: {entry_point}")
                passed = False
        
        return passed
    
    def _check_test_data(self) -> bool:
        """Check test data scenarios."""
        test_dir = self.package_dir / "test_data"
        
        if not test_dir.exists():
            self.warnings.append("Test data directory not found")
            print("⚠️  test_data/ - Directory missing (optional)")
            return True
        
        expected_scenarios = [
            "scenario_1_normal_paper",
            "scenario_2_no_toc_document", 
            "scenario_3_duplicate_headings",
            "scenario_4_headings_in_tables",
            "scenario_5_missing_fonts",
            "scenario_6_complex_objects",
            "scenario_7_revision_tracking"
        ]
        
        passed = True
        for scenario in expected_scenarios:
            scenario_dir = test_dir / scenario
            if scenario_dir.exists():
                # Check required files
                required_files = ["user_intent.txt", "expected_plan.v1.json"]
                scenario_complete = True
                
                for req_file in required_files:
                    if not (scenario_dir / req_file).exists():
                        scenario_complete = False
                        break
                
                if scenario_complete:
                    print(f"✓ {scenario}")
                else:
                    print(f"⚠️  {scenario} - Incomplete")
                    self.warnings.append(f"Incomplete test scenario: {scenario}")
            else:
                print(f"✗ {scenario} - Missing")
                self.errors.append(f"Missing test scenario: {scenario}")
                passed = False
        
        return passed
    
    def _check_documentation(self) -> bool:
        """Check documentation files."""
        required_docs = [
            "KNOWN-ISSUES.md",
            "README.md",
            "INSTALLATION.md"
        ]
        
        passed = True
        for doc_file in required_docs:
            doc_path = self.package_dir / doc_file
            if doc_path.exists():
                print(f"✓ {doc_file}")
            else:
                print(f"⚠️  {doc_file} - Missing")
                self.warnings.append(f"Missing documentation: {doc_file}")
        
        return passed
    
    def _print_summary(self, all_passed: bool):
        """Print verification summary."""
        print("\n" + "=" * 50)
        print("📊 VERIFICATION SUMMARY")
        print("=" * 50)
        
        if all_passed:
            print("🎉 DEPLOYMENT VERIFICATION PASSED")
            print("   All critical checks completed successfully!")
        else:
            print("❌ DEPLOYMENT VERIFICATION FAILED")
            print("   Critical issues found that must be resolved.")
        
        if self.errors:
            print(f"\n🚨 ERRORS ({len(self.errors)}):")
            for i, error in enumerate(self.errors, 1):
                print(f"   {i}. {error}")
        
        if self.warnings:
            print(f"\n⚠️  WARNINGS ({len(self.warnings)}):")
            for i, warning in enumerate(self.warnings, 1):
                print(f"   {i}. {warning}")
        
        if not self.errors and not self.warnings:
            print("\n✨ Perfect deployment! No issues found.")
        
        print("\n" + "=" * 50)

def main():
    """Main verification function."""
    import argparse
    
    parser = argparse.ArgumentParser(description="AutoWord vNext Deployment Verifier")
    parser.add_argument("--package-dir", default=".",
                       help="Path to deployment package directory")
    parser.add_argument("--quiet", action="store_true",
                       help="Suppress detailed output")
    
    args = parser.parse_args()
    
    verifier = DeploymentVerifier(args.package_dir)
    
    if args.quiet:
        # Redirect stdout to suppress detailed output
        import io
        old_stdout = sys.stdout
        sys.stdout = io.StringIO()
        
        try:
            success = verifier.verify_deployment()
        finally:
            sys.stdout = old_stdout
        
        # Print only summary
        verifier._print_summary(success)
    else:
        success = verifier.verify_deployment()
    
    sys.exit(0 if success else 1)

if __name__ == "__main__":
    main()