#!/usr/bin/env python3
"""
Comprehensive Test Runner for AutoWord vNext.

This script runs all comprehensive test automation including:
- Atomic operations testing with real Word documents
- DoD regression testing
- Performance benchmarking
- Stress testing
- CI integration testing

Usage:
    python run_comprehensive_tests.py [options]

Options:
    --suite <suite_name>    Run specific test suite (comprehensive, dod, performance, ci)
    --output-dir <dir>      Output directory for test results
    --verbose               Enable verbose output
    --junit                 Generate JUnit XML reports
    --html                  Generate HTML reports
    --ci                    Run in CI mode (faster, essential tests only)
    --help                  Show this help message
"""

import sys
import os
import argparse
import json
import time
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Any, Optional

# Add the project root to Python path
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

# Import test suites
try:
    from tests.test_comprehensive_automation import TestAutomationRunner
    from tests.test_dod_validation_suite import DoDValidationSuite
    from tests.test_performance_benchmarks import PerformanceTestRunner
    from tests.test_ci_integration import CITestRunner
except ImportError as e:
    print(f"Error importing test modules: {e}")
    print("Please ensure all test dependencies are installed.")
    sys.exit(1)


class ComprehensiveTestRunner:
    """Main comprehensive test runner orchestrator."""
    
    def __init__(self, output_dir: str = "test_results", verbose: bool = False):
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(exist_ok=True)
        self.verbose = verbose
        self.start_time = None
        self.results = {}
        
    def run_all_tests(self, ci_mode: bool = False) -> Dict[str, Any]:
        """Run all comprehensive test suites."""
        self.start_time = time.time()
        
        print("🚀 Starting AutoWord vNext Comprehensive Test Automation")
        print(f"📁 Output directory: {self.output_dir}")
        print(f"⏰ Started at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
        if ci_mode:
            print("🔄 Running in CI mode (essential tests only)")
        
        try:
            # Run test suites
            if not ci_mode:
                self.results['comprehensive'] = self._run_comprehensive_tests()
                self.results['performance'] = self._run_performance_tests()
                
            self.results['dod'] = self._run_dod_tests()
            self.results['ci'] = self._run_ci_tests()
            
            # Generate summary
            self.results['summary'] = self._generate_overall_summary()
            
            # Save results
            self._save_results()
            
            # Print final summary
            self._print_final_summary()
            
        except KeyboardInterrupt:
            print("\n⚠️  Test execution interrupted by user")
            self.results['summary'] = {'status': 'INTERRUPTED', 'message': 'Test execution was interrupted'}
        except Exception as e:
            print(f"\n💥 Test execution failed: {e}")
            self.results['summary'] = {'status': 'ERROR', 'message': str(e)}
            
        return self.results
        
    def run_specific_suite(self, suite_name: str) -> Dict[str, Any]:
        """Run a specific test suite."""
        self.start_time = time.time()
        
        print(f"🎯 Running {suite_name} test suite")
        
        suite_runners = {
            'comprehensive': self._run_comprehensive_tests,
            'dod': self._run_dod_tests,
            'performance': self._run_performance_tests,
            'ci': self._run_ci_tests
        }
        
        if suite_name not in suite_runners:
            raise ValueError(f"Unknown test suite: {suite_name}")
            
        try:
            self.results[suite_name] = suite_runners[suite_name]()
            self.results['summary'] = self._generate_suite_summary(suite_name)
            self._save_results()
            self._print_final_summary()
            
        except Exception as e:
            print(f"💥 {suite_name} test suite failed: {e}")
            self.results['summary'] = {'status': 'ERROR', 'message': str(e)}
            
        return self.results
        
    def _run_comprehensive_tests(self) -> Dict[str, Any]:
        """Run comprehensive automation tests."""
        print("\n📋 Running Comprehensive Automation Tests...")
        
        try:
            runner = TestAutomationRunner(str(self.output_dir / "comprehensive"))
            results = runner.run_comprehensive_tests()
            
            if self.verbose:
                self._print_suite_details("Comprehensive", results)
                
            return {
                'status': 'COMPLETED',
                'results': results,
                'summary': results.get('summary', {})
            }
            
        except Exception as e:
            print(f"❌ Comprehensive tests failed: {e}")
            return {
                'status': 'FAILED',
                'error': str(e)
            }
            
    def _run_dod_tests(self) -> Dict[str, Any]:
        """Run DoD validation tests."""
        print("\n✅ Running DoD Validation Tests...")
        
        try:
            dod_suite = DoDValidationSuite()
            results = dod_suite.run_all_dod_validations()
            
            if self.verbose:
                self._print_suite_details("DoD Validation", results)
                
            return {
                'status': 'COMPLETED',
                'results': results,
                'summary': results.get('overall_assessment', {})
            }
            
        except Exception as e:
            print(f"❌ DoD validation tests failed: {e}")
            return {
                'status': 'FAILED',
                'error': str(e)
            }
            
    def _run_performance_tests(self) -> Dict[str, Any]:
        """Run performance benchmark tests."""
        print("\n⚡ Running Performance Benchmark Tests...")
        
        try:
            runner = PerformanceTestRunner()
            results = runner.run_performance_tests()
            
            if self.verbose:
                self._print_suite_details("Performance", results)
                
            return {
                'status': 'COMPLETED',
                'results': results,
                'summary': results.get('overall_assessment', {})
            }
            
        except Exception as e:
            print(f"❌ Performance tests failed: {e}")
            return {
                'status': 'FAILED',
                'error': str(e)
            }
            
    def _run_ci_tests(self) -> Dict[str, Any]:
        """Run CI integration tests."""
        print("\n🔄 Running CI Integration Tests...")
        
        try:
            runner = CITestRunner()
            # Capture the results instead of just the exit code
            ci_suite = runner.ci_suite
            results = ci_suite.run_ci_tests()
            
            if self.verbose:
                self._print_suite_details("CI Integration", results)
                
            return {
                'status': 'COMPLETED',
                'results': results,
                'summary': results.get('ci_summary', {})
            }
            
        except Exception as e:
            print(f"❌ CI integration tests failed: {e}")
            return {
                'status': 'FAILED',
                'error': str(e)
            }
            
    def _print_suite_details(self, suite_name: str, results: Dict[str, Any]):
        """Print detailed results for a test suite."""
        print(f"\n📊 {suite_name} Test Suite Details:")
        
        if 'summary' in results:
            summary = results['summary']
            print(f"  Status: {summary.get('status', 'Unknown')}")
            
            if 'total_tests_run' in summary:
                print(f"  Total Tests: {summary['total_tests_run']}")
                print(f"  Passed: {summary.get('total_tests_passed', 0)}")
                print(f"  Failed: {summary.get('total_tests_failed', 0)}")
                print(f"  Success Rate: {summary.get('overall_success_rate', 0):.2%}")
                
        # Print any critical issues
        if 'critical_failures' in results.get('summary', {}):
            failures = results['summary']['critical_failures']
            if failures:
                print(f"  Critical Failures: {len(failures)}")
                for failure in failures[:3]:  # Show first 3
                    print(f"    - {failure}")
                if len(failures) > 3:
                    print(f"    ... and {len(failures) - 3} more")
                    
    def _generate_overall_summary(self) -> Dict[str, Any]:
        """Generate overall test summary."""
        summary = {
            'execution_time_seconds': time.time() - self.start_time if self.start_time else 0,
            'timestamp': datetime.now().isoformat(),
            'total_suites_run': 0,
            'successful_suites': 0,
            'failed_suites': 0,
            'overall_status': 'PASSED',
            'critical_issues': [],
            'recommendations': []
        }
        
        # Analyze suite results
        for suite_name, suite_result in self.results.items():
            if suite_name == 'summary':
                continue
                
            summary['total_suites_run'] += 1
            
            if suite_result.get('status') == 'COMPLETED':
                summary['successful_suites'] += 1
            else:
                summary['failed_suites'] += 1
                summary['overall_status'] = 'FAILED'
                summary['critical_issues'].append(f"{suite_name} suite failed")
                
        # Generate recommendations
        if summary['overall_status'] == 'FAILED':
            summary['recommendations'].append("Address all failing test suites before release")
            
        if summary['failed_suites'] > 0:
            summary['recommendations'].append(f"Investigate and fix {summary['failed_suites']} failing test suites")
            
        # Check for specific issues
        dod_result = self.results.get('dod', {})
        if dod_result.get('status') == 'FAILED':
            summary['critical_issues'].append("DoD compliance failures detected")
            summary['recommendations'].append("DoD compliance is mandatory - resolve all DoD issues")
            
        return summary
        
    def _generate_suite_summary(self, suite_name: str) -> Dict[str, Any]:
        """Generate summary for a specific suite."""
        suite_result = self.results.get(suite_name, {})
        
        return {
            'execution_time_seconds': time.time() - self.start_time if self.start_time else 0,
            'timestamp': datetime.now().isoformat(),
            'suite_name': suite_name,
            'suite_status': suite_result.get('status', 'UNKNOWN'),
            'suite_summary': suite_result.get('summary', {}),
            'overall_status': 'PASSED' if suite_result.get('status') == 'COMPLETED' else 'FAILED'
        }
        
    def _save_results(self):
        """Save test results to files."""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # Save JSON results
        json_path = self.output_dir / f"comprehensive_results_{timestamp}.json"
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(self.results, f, indent=2, default=str)
            
        # Save summary report
        summary_path = self.output_dir / f"test_summary_{timestamp}.txt"
        with open(summary_path, 'w', encoding='utf-8') as f:
            f.write(self._generate_text_summary())
            
        print(f"📄 Results saved to: {json_path}")
        print(f"📄 Summary saved to: {summary_path}")
        
    def _generate_text_summary(self) -> str:
        """Generate text summary report."""
        summary = self.results.get('summary', {})
        
        report = f"""
AutoWord vNext Comprehensive Test Results
=========================================

Execution Time: {summary.get('execution_time_seconds', 0):.2f} seconds
Timestamp: {summary.get('timestamp', 'Unknown')}
Overall Status: {summary.get('overall_status', 'Unknown')}

Test Suites Summary:
- Total Suites Run: {summary.get('total_suites_run', 0)}
- Successful Suites: {summary.get('successful_suites', 0)}
- Failed Suites: {summary.get('failed_suites', 0)}

"""
        
        if summary.get('critical_issues'):
            report += "Critical Issues:\n"
            for issue in summary['critical_issues']:
                report += f"- {issue}\n"
            report += "\n"
            
        if summary.get('recommendations'):
            report += "Recommendations:\n"
            for rec in summary['recommendations']:
                report += f"- {rec}\n"
            report += "\n"
            
        # Add suite details
        for suite_name, suite_result in self.results.items():
            if suite_name == 'summary':
                continue
                
            report += f"{suite_name.title()} Suite:\n"
            report += f"- Status: {suite_result.get('status', 'Unknown')}\n"
            
            if 'summary' in suite_result:
                suite_summary = suite_result['summary']
                if 'total_tests_run' in suite_summary:
                    report += f"- Tests Run: {suite_summary['total_tests_run']}\n"
                    report += f"- Success Rate: {suite_summary.get('overall_success_rate', 0):.2%}\n"
                    
            report += "\n"
            
        return report
        
    def _print_final_summary(self):
        """Print final test summary."""
        summary = self.results.get('summary', {})
        
        print("\n" + "="*60)
        print("🏁 COMPREHENSIVE TEST EXECUTION COMPLETE")
        print("="*60)
        
        status_emoji = "✅" if summary.get('overall_status') == 'PASSED' else "❌"
        print(f"{status_emoji} Overall Status: {summary.get('overall_status', 'Unknown')}")
        print(f"⏱️  Execution Time: {summary.get('execution_time_seconds', 0):.2f} seconds")
        print(f"📊 Test Suites: {summary.get('successful_suites', 0)}/{summary.get('total_suites_run', 0)} passed")
        
        if summary.get('critical_issues'):
            print(f"\n🚨 Critical Issues ({len(summary['critical_issues'])}):")
            for issue in summary['critical_issues']:
                print(f"   • {issue}")
                
        if summary.get('recommendations'):
            print(f"\n💡 Recommendations:")
            for rec in summary['recommendations']:
                print(f"   • {rec}")
                
        print(f"\n📁 Detailed results available in: {self.output_dir}")


def main():
    """Main entry point."""
    parser = argparse.ArgumentParser(
        description="AutoWord vNext Comprehensive Test Runner",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__
    )
    
    parser.add_argument(
        '--suite',
        choices=['comprehensive', 'dod', 'performance', 'ci'],
        help='Run specific test suite'
    )
    
    parser.add_argument(
        '--output-dir',
        default='test_results',
        help='Output directory for test results (default: test_results)'
    )
    
    parser.add_argument(
        '--verbose', '-v',
        action='store_true',
        help='Enable verbose output'
    )
    
    parser.add_argument(
        '--ci',
        action='store_true',
        help='Run in CI mode (faster, essential tests only)'
    )
    
    parser.add_argument(
        '--junit',
        action='store_true',
        help='Generate JUnit XML reports'
    )
    
    parser.add_argument(
        '--html',
        action='store_true',
        help='Generate HTML reports'
    )
    
    args = parser.parse_args()
    
    # Create test runner
    runner = ComprehensiveTestRunner(
        output_dir=args.output_dir,
        verbose=args.verbose
    )
    
    try:
        # Run tests
        if args.suite:
            results = runner.run_specific_suite(args.suite)
        else:
            results = runner.run_all_tests(ci_mode=args.ci)
            
        # Determine exit code
        summary = results.get('summary', {})
        if summary.get('overall_status') == 'PASSED':
            print("\n🎉 All tests completed successfully!")
            return 0
        else:
            print("\n💥 Some tests failed - check the results for details")
            return 1
            
    except KeyboardInterrupt:
        print("\n⚠️  Test execution interrupted")
        return 130
    except Exception as e:
        print(f"\n💥 Test runner error: {e}")
        return 2


if __name__ == "__main__":
    sys.exit(main())