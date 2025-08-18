#!/usr/bin/env python3
"""
Test runner for VNextPipeline integration tests.

This script runs the complete integration test suite for the pipeline orchestrator
and provides detailed reporting on test results.
"""

import sys
import os
import subprocess
import tempfile
import json
from pathlib import Path

# Add the autoword package to Python path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..'))

def run_integration_tests():
    """Run the complete integration test suite."""
    print("=" * 60)
    print("AutoWord vNext Pipeline Integration Tests")
    print("=" * 60)
    
    # Change to the vnext directory
    vnext_dir = Path(__file__).parent
    os.chdir(vnext_dir)
    
    try:
        # Run pytest with verbose output
        cmd = [
            sys.executable, "-m", "pytest", 
            "tests/test_pipeline_integration.py",
            "-v",
            "--tb=short",
            "--no-header"
        ]
        
        print(f"Running command: {' '.join(cmd)}")
        print("-" * 60)
        
        result = subprocess.run(cmd, capture_output=True, text=True)
        
        print("STDOUT:")
        print(result.stdout)
        
        if result.stderr:
            print("STDERR:")
            print(result.stderr)
        
        print("-" * 60)
        print(f"Test execution completed with return code: {result.returncode}")
        
        if result.returncode == 0:
            print("✅ All integration tests passed!")
        else:
            print("❌ Some integration tests failed!")
            
        return result.returncode == 0
        
    except Exception as e:
        print(f"❌ Error running tests: {e}")
        return False

def create_sample_test_data():
    """Create minimal sample test data for testing."""
    test_data_dir = Path("test_data")
    
    # Create scenario 1 if it doesn't exist
    scenario_1_dir = test_data_dir / "scenario_1_normal_paper"
    scenario_1_dir.mkdir(parents=True, exist_ok=True)
    
    # Create user intent file
    user_intent_file = scenario_1_dir / "user_intent.txt"
    if not user_intent_file.exists():
        with open(user_intent_file, 'w', encoding='utf-8') as f:
            f.write("Remove abstract and references sections, update table of contents")
    
    # Create expected plan file
    expected_plan_file = scenario_1_dir / "expected_plan.v1.json"
    if not expected_plan_file.exists():
        expected_plan = {
            "schema_version": "plan.v1",
            "ops": [
                {
                    "operation_type": "delete_section_by_heading",
                    "heading_text": "摘要",
                    "level": 1,
                    "match": "EXACT",
                    "case_sensitive": False
                },
                {
                    "operation_type": "delete_section_by_heading",
                    "heading_text": "参考文献",
                    "level": 1,
                    "match": "EXACT",
                    "case_sensitive": False
                },
                {
                    "operation_type": "update_toc"
                }
            ]
        }
        
        with open(expected_plan_file, 'w', encoding='utf-8') as f:
            json.dump(expected_plan, f, indent=2, ensure_ascii=False)
    
    print(f"✅ Sample test data created in {test_data_dir}")

def run_basic_pipeline_test():
    """Run a basic pipeline functionality test."""
    print("\n" + "=" * 60)
    print("Basic Pipeline Functionality Test")
    print("=" * 60)
    
    try:
        from autoword.vnext.pipeline import VNextPipeline, ProgressReporter
        from autoword.vnext.models import StructureV1, DocumentMetadata
        
        # Test progress reporter
        print("Testing ProgressReporter...")
        callback_calls = []
        
        def test_callback(stage, progress):
            callback_calls.append((stage, progress))
            print(f"  Progress: {stage} - {progress}%")
        
        reporter = ProgressReporter(test_callback)
        
        stages = ["Extract", "Plan", "Execute", "Validate", "Audit"]
        for stage in stages:
            reporter.start_stage(stage)
            reporter.report_substep(f"Processing {stage.lower()}")
            reporter.complete_stage()
        
        print(f"✅ ProgressReporter test completed. Recorded {len(callback_calls)} callbacks.")
        
        # Test pipeline initialization
        print("\nTesting VNextPipeline initialization...")
        pipeline = VNextPipeline(
            base_audit_dir="./test_audit",
            visible=False,
            progress_callback=test_callback
        )
        
        print("✅ VNextPipeline initialization test completed.")
        
        # Test data model creation
        print("\nTesting data model creation...")
        structure = StructureV1(
            metadata=DocumentMetadata(
                title="Test Document",
                author="Test Author",
                paragraph_count=5
            )
        )
        
        print(f"✅ Created StructureV1 with title: {structure.metadata.title}")
        
        return True
        
    except Exception as e:
        print(f"❌ Basic pipeline test failed: {e}")
        import traceback
        traceback.print_exc()
        return False

def main():
    """Main test runner function."""
    print("AutoWord vNext Pipeline Test Runner")
    print("This script tests the main pipeline orchestrator implementation.")
    print()
    
    # Create sample test data
    create_sample_test_data()
    
    # Run basic functionality test
    basic_test_passed = run_basic_pipeline_test()
    
    # Run integration tests
    integration_tests_passed = run_integration_tests()
    
    print("\n" + "=" * 60)
    print("Test Summary")
    print("=" * 60)
    print(f"Basic functionality test: {'✅ PASSED' if basic_test_passed else '❌ FAILED'}")
    print(f"Integration tests: {'✅ PASSED' if integration_tests_passed else '❌ FAILED'}")
    
    if basic_test_passed and integration_tests_passed:
        print("\n🎉 All tests passed! Pipeline orchestrator is working correctly.")
        return 0
    else:
        print("\n⚠️  Some tests failed. Please check the output above for details.")
        return 1

if __name__ == "__main__":
    sys.exit(main())