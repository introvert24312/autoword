"""
Demonstration of the comprehensive logging and monitoring system.

This script shows how to use the VNext monitoring system for detailed
operation logging, performance tracking, memory monitoring, and debug logging.
"""

import os
import sys
import time
import tempfile
from pathlib import Path

# Add the parent directory to the path to import vnext modules
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..', '..'))

from autoword.vnext.monitoring import (
    VNextLogger, MonitoringLevel, create_vnext_logger,
    log_large_document_warning, log_complex_document_scenario
)


def simulate_document_extraction(logger: VNextLogger):
    """Simulate document extraction with monitoring."""
    print("=== Simulating Document Extraction ===")
    
    with logger.track_operation("document_loading", file_size_mb=25.5, paragraph_count=1500):
        time.sleep(0.1)  # Simulate loading time
        logger.log_debug("Document loaded successfully", format="DOCX", version="2019")
    
    with logger.track_operation("structure_parsing"):
        time.sleep(0.2)  # Simulate parsing time
        logger.log_debug("Structure parsing completed", 
                        headings=45, paragraphs=1500, tables=8, images=12)
    
    with logger.track_operation("inventory_extraction"):
        time.sleep(0.15)  # Simulate inventory extraction
        logger.log_debug("Inventory extraction completed",
                        formulas=3, charts=2, content_controls=5)
    
    # Log complex document scenario
    log_complex_document_scenario(logger, "contains_formulas", {
        "formula_count": 3,
        "complexity": "medium",
        "recommendation": "Use careful processing for formula preservation"
    })


def simulate_plan_generation(logger: VNextLogger):
    """Simulate plan generation with monitoring."""
    print("=== Simulating Plan Generation ===")
    
    with logger.track_operation("llm_request_preparation", model="gpt-4", tokens_estimated=2500):
        time.sleep(0.05)
        logger.log_debug("LLM request prepared", prompt_length=2500, model="gpt-4")
    
    with logger.track_operation("llm_api_call"):
        time.sleep(0.8)  # Simulate LLM API call
        logger.log_debug("LLM API call completed", 
                        response_tokens=450, processing_time_ms=800)
    
    with logger.track_operation("plan_validation", operations_count=6):
        time.sleep(0.1)
        logger.log_debug("Plan validation completed", 
                        operations=6, validation_errors=0, warnings=1)
        logger.log_warning("Plan contains clear_direct_formatting operation", 
                          {"operation": "clear_direct_formatting", "requires_authorization": True})


def simulate_plan_execution(logger: VNextLogger):
    """Simulate plan execution with monitoring."""
    print("=== Simulating Plan Execution ===")
    
    operations = [
        ("delete_section_by_heading", 0.3, {"heading": "摘要", "level": 1}),
        ("set_style_rule", 0.2, {"style": "Heading 1", "font": "楷体"}),
        ("reassign_paragraphs_to_style", 0.4, {"target_style": "Normal", "count": 25}),
        ("update_toc", 0.6, {"toc_entries": 45, "pages_updated": 12}),
        ("clear_direct_formatting", 0.3, {"scope": "document", "authorization": True}),
    ]
    
    for i, (op_name, duration, metadata) in enumerate(operations):
        with logger.track_operation(f"atomic_operation_{i+1}", 
                                   operation_type=op_name, **metadata):
            time.sleep(duration)
            logger.log_debug(f"Atomic operation completed: {op_name}", **metadata)


def simulate_validation(logger: VNextLogger):
    """Simulate document validation with monitoring."""
    print("=== Simulating Document Validation ===")
    
    with logger.track_operation("structure_comparison"):
        time.sleep(0.2)
        logger.log_debug("Structure comparison completed", 
                        changes_detected=5, paragraphs_modified=3)
    
    with logger.track_operation("assertion_checking"):
        time.sleep(0.15)
        
        # Simulate assertion checks
        assertions = [
            ("chapter_assertions", True, "No 摘要/参考文献 at level 1"),
            ("style_assertions", True, "H1/H2/Normal styles match specifications"),
            ("toc_assertions", True, "TOC items match heading tree"),
            ("pagination_assertions", True, "Fields updated, metadata changed")
        ]
        
        for assertion_name, passed, message in assertions:
            if passed:
                logger.log_debug(f"Assertion passed: {assertion_name}", assertion_message=message)
            else:
                logger.log_warning(f"Assertion failed: {assertion_name}", assertion_message=message)


def simulate_audit_trail(logger: VNextLogger):
    """Simulate audit trail creation with monitoring."""
    print("=== Simulating Audit Trail Creation ===")
    
    with logger.track_operation("snapshot_creation"):
        time.sleep(0.1)
        logger.log_debug("Snapshots created", 
                        before_docx="original.docx", after_docx="modified.docx")
    
    with logger.track_operation("diff_report_generation"):
        time.sleep(0.2)
        logger.log_debug("Diff report generated",
                        added_paragraphs=0, removed_paragraphs=2, modified_paragraphs=3)
    
    with logger.track_operation("status_writing"):
        time.sleep(0.05)
        logger.log_debug("Status written", status="SUCCESS", completion_message="Pipeline completed successfully")


def demonstrate_memory_monitoring():
    """Demonstrate memory monitoring capabilities."""
    print("\n=== Memory Monitoring Demonstration ===")
    
    with tempfile.TemporaryDirectory() as temp_dir:
        # Create logger with aggressive memory monitoring for demo
        logger = create_vnext_logger(
            audit_directory=temp_dir,
            monitoring_level=MonitoringLevel.DEBUG,
            enable_memory_monitoring=True,
            memory_warning_threshold_mb=100,  # Low threshold for demo
            memory_critical_threshold_mb=200
        )
        
        print(f"Monitoring logs will be saved to: {temp_dir}")
        
        # Simulate memory-intensive operation
        with logger.track_operation("memory_intensive_operation"):
            # Allocate some memory to potentially trigger alerts
            data = []
            for i in range(1000):
                data.append("x" * 1000)  # Allocate some memory
                if i % 100 == 0:
                    time.sleep(0.01)  # Give monitoring time to check
            
            logger.log_debug("Memory intensive operation completed", 
                           data_size=len(data), memory_allocated_kb=len(data))
        
        # Check for memory alerts
        if logger.memory_monitor:
            alerts = logger.memory_monitor.get_alerts()
            if alerts:
                print(f"Memory alerts generated: {len(alerts)}")
                for alert in alerts[-3:]:  # Show last 3 alerts
                    print(f"  - {alert.alert_level}: {alert.message}")
            else:
                print("No memory alerts generated (memory usage was below thresholds)")
            
            peak_memory = logger.memory_monitor.get_peak_memory_mb()
            print(f"Peak memory usage: {peak_memory:.1f}MB")
        
        logger.cleanup()


def demonstrate_performance_reporting():
    """Demonstrate performance reporting capabilities."""
    print("\n=== Performance Reporting Demonstration ===")
    
    with tempfile.TemporaryDirectory() as temp_dir:
        logger = create_vnext_logger(
            audit_directory=temp_dir,
            monitoring_level=MonitoringLevel.PERFORMANCE,
            enable_memory_monitoring=True
        )
        
        print(f"Performance logs will be saved to: {temp_dir}")
        
        # Simulate a complete pipeline with performance tracking
        with logger.track_stage("Complete_Pipeline"):
            with logger.track_stage("Extract"):
                simulate_document_extraction(logger)
            
            with logger.track_stage("Plan"):
                simulate_plan_generation(logger)
            
            with logger.track_stage("Execute"):
                simulate_plan_execution(logger)
            
            with logger.track_stage("Validate"):
                simulate_validation(logger)
            
            with logger.track_stage("Audit"):
                simulate_audit_trail(logger)
        
        # Generate and display performance report
        report = logger.generate_performance_report()
        
        print("\n--- Performance Summary ---")
        print(f"Monitoring Level: {report['monitoring_level']}")
        print(f"Total Operations: {report['operation_stats']['total_operations']}")
        print(f"Total Stages: {report['stage_stats']['total_stages']}")
        print(f"Success Rate: {report['operation_stats']['success_rate']:.1%}")
        
        if 'total_pipeline_duration_ms' in report['stage_stats']:
            total_duration = report['stage_stats']['total_pipeline_duration_ms']
            print(f"Total Pipeline Duration: {total_duration:.1f}ms")
        
        print("\n--- Stage Breakdown ---")
        for stage_name, stage_data in report['stage_stats']['stage_breakdown'].items():
            duration = stage_data.get('duration_ms', 0)
            ops_count = stage_data.get('operations_count', 0)
            print(f"  {stage_name}: {duration:.1f}ms ({ops_count} operations)")
        
        if 'memory_stats' in report:
            memory_stats = report['memory_stats']
            print(f"\n--- Memory Usage ---")
            print(f"Peak Memory: {memory_stats['peak_memory_mb']:.1f}MB")
            print(f"Current Memory: {memory_stats['current_memory_mb']:.1f}MB")
            if memory_stats['alerts']:
                print(f"Memory Alerts: {len(memory_stats['alerts'])}")
        
        # Save performance report
        logger.save_performance_report()
        
        report_file = Path(temp_dir) / "performance_report.json"
        print(f"\nDetailed performance report saved to: {report_file}")
        
        # Show log files created
        print(f"\nLog files created:")
        for log_file in Path(temp_dir).glob("*.log"):
            size_kb = log_file.stat().st_size / 1024
            print(f"  - {log_file.name}: {size_kb:.1f}KB")
        
        # Explicit cleanup before temp directory cleanup
        logger.cleanup()
        time.sleep(0.1)  # Give Windows time to release file handles


def demonstrate_different_monitoring_levels():
    """Demonstrate different monitoring levels."""
    print("\n=== Different Monitoring Levels Demonstration ===")
    
    levels = [
        (MonitoringLevel.BASIC, "Basic monitoring - essential operations only"),
        (MonitoringLevel.DETAILED, "Detailed monitoring - all operations with timing"),
        (MonitoringLevel.DEBUG, "Debug monitoring - full debug with memory tracking"),
        (MonitoringLevel.PERFORMANCE, "Performance monitoring - optimization focus")
    ]
    
    for level, description in levels:
        print(f"\n--- {description} ---")
        
        with tempfile.TemporaryDirectory() as temp_dir:
            logger = create_vnext_logger(
                audit_directory=temp_dir,
                monitoring_level=level,
                enable_memory_monitoring=(level in [MonitoringLevel.DEBUG, MonitoringLevel.PERFORMANCE])
            )
            
            # Perform some operations
            with logger.track_stage("Demo_Stage"):
                with logger.track_operation("demo_operation", param="value"):
                    time.sleep(0.05)
                    logger.log_debug("Debug message", detail="This is debug info")
                    logger.log_warning("Warning message", context={"level": level.value})
            
            # Check what was logged
            log_files = list(Path(temp_dir).glob("*.log"))
            print(f"Log files created: {[f.name for f in log_files]}")
            
            # Show pipeline log content (first few lines)
            pipeline_log = Path(temp_dir) / "pipeline.log"
            if pipeline_log.exists():
                content = pipeline_log.read_text(encoding='utf-8')
                lines = content.split('\n')[:5]  # First 5 lines
                print("Sample log content:")
                for line in lines:
                    if line.strip():
                        print(f"  {line}")
            
            logger.cleanup()
            time.sleep(0.1)  # Give Windows time to release file handles


def main():
    """Main demonstration function."""
    print("AutoWord vNext - Comprehensive Logging and Monitoring Demonstration")
    print("=" * 70)
    
    # Demonstrate memory monitoring
    demonstrate_memory_monitoring()
    
    # Demonstrate performance reporting
    demonstrate_performance_reporting()
    
    # Demonstrate different monitoring levels
    demonstrate_different_monitoring_levels()
    
    print("\n" + "=" * 70)
    print("Demonstration completed!")
    print("\nKey features demonstrated:")
    print("✓ Detailed operation logging throughout all pipeline stages")
    print("✓ Performance monitoring for large document processing")
    print("✓ Debug logging for troubleshooting complex document scenarios")
    print("✓ Memory usage monitoring and optimization alerts")
    print("✓ Execution time tracking for each pipeline stage and atomic operation")
    print("\nThe monitoring system provides comprehensive observability")
    print("into the vNext pipeline execution with configurable detail levels.")


if __name__ == "__main__":
    main()