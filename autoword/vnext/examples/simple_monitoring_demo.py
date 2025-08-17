"""
Simple demonstration of the monitoring system without file cleanup issues.

This script shows the core monitoring functionality in a simplified way.
"""

import os
import sys
import time
from pathlib import Path

# Add the parent directory to the path to import vnext modules
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..', '..'))

from autoword.vnext.monitoring import (
    VNextLogger, MonitoringLevel, PerformanceTracker, MemoryMonitor
)


def demonstrate_performance_tracking():
    """Demonstrate performance tracking without file I/O."""
    print("=== Performance Tracking Demonstration ===")
    
    tracker = PerformanceTracker()
    
    # Simulate some operations
    print("Tracking operations...")
    
    with tracker.track_operation("document_loading", {"file_size_mb": 25.5}):
        time.sleep(0.1)
        print("  - Document loaded")
    
    with tracker.track_operation("structure_parsing"):
        time.sleep(0.2)
        print("  - Structure parsed")
    
    # Simulate a failed operation
    try:
        with tracker.track_operation("failing_operation"):
            time.sleep(0.05)
            raise ValueError("Simulated error")
    except ValueError:
        print("  - Operation failed (expected)")
    
    # Track a pipeline stage
    with tracker.track_stage("Complete_Processing"):
        with tracker.track_operation("validation"):
            time.sleep(0.1)
            print("  - Validation completed")
        
        with tracker.track_operation("audit"):
            time.sleep(0.05)
            print("  - Audit completed")
    
    # Get statistics
    op_stats = tracker.get_operation_stats()
    stage_stats = tracker.get_stage_stats()
    
    print(f"\n--- Operation Statistics ---")
    print(f"Total operations: {op_stats['total_operations']}")
    print(f"Successful operations: {op_stats['successful_operations']}")
    print(f"Failed operations: {op_stats['failed_operations']}")
    print(f"Success rate: {op_stats['success_rate']:.1%}")
    
    if 'total_duration_ms' in op_stats:
        print(f"Total duration: {op_stats['total_duration_ms']:.1f}ms")
        print(f"Average duration: {op_stats['average_duration_ms']:.1f}ms")
    
    print(f"\n--- Stage Statistics ---")
    print(f"Total stages: {stage_stats['total_stages']}")
    print(f"Successful stages: {stage_stats['successful_stages']}")
    
    if 'stage_breakdown' in stage_stats:
        print("Stage breakdown:")
        for stage_name, stage_data in stage_stats['stage_breakdown'].items():
            duration = stage_data.get('duration_ms', 0)
            ops_count = stage_data.get('operations_count', 0)
            print(f"  - {stage_name}: {duration:.1f}ms ({ops_count} operations)")


def demonstrate_memory_monitoring():
    """Demonstrate memory monitoring."""
    print("\n=== Memory Monitoring Demonstration ===")
    
    # Create memory monitor with low thresholds for demo
    monitor = MemoryMonitor(
        warning_threshold_mb=50,  # Low threshold for demo
        critical_threshold_mb=100,
        check_interval_seconds=0.1
    )
    
    print(f"Current memory usage: {monitor.get_current_memory_mb():.1f}MB")
    print(f"Warning threshold: {monitor.warning_threshold_mb}MB")
    print(f"Critical threshold: {monitor.critical_threshold_mb}MB")
    
    # Start monitoring
    monitor.start_monitoring()
    print("Memory monitoring started...")
    
    # Simulate some memory-intensive work
    monitor.set_current_operation("memory_test")
    
    # Allocate some memory
    data = []
    for i in range(500):
        data.append("x" * 1000)
        if i % 100 == 0:
            time.sleep(0.05)  # Give monitoring time to check
    
    time.sleep(0.2)  # Let monitoring run
    
    # Stop monitoring
    monitor.stop_monitoring()
    
    # Check results
    alerts = monitor.get_alerts()
    peak_memory = monitor.get_peak_memory_mb()
    
    print(f"Peak memory usage: {peak_memory:.1f}MB")
    print(f"Memory alerts generated: {len(alerts)}")
    
    for alert in alerts:
        print(f"  - {alert.alert_level}: {alert.message}")
    
    # Clean up
    del data


def demonstrate_monitoring_levels():
    """Demonstrate different monitoring levels."""
    print("\n=== Monitoring Levels Demonstration ===")
    
    levels = [
        MonitoringLevel.BASIC,
        MonitoringLevel.DETAILED,
        MonitoringLevel.DEBUG,
        MonitoringLevel.PERFORMANCE
    ]
    
    for level in levels:
        print(f"\n--- {level.value} Level ---")
        
        # Create a temporary directory that we'll clean up manually
        temp_dir = f"./temp_monitoring_{level.value.lower()}"
        os.makedirs(temp_dir, exist_ok=True)
        
        try:
            logger = VNextLogger(
                audit_directory=temp_dir,
                monitoring_level=level,
                enable_memory_monitoring=False  # Disable to avoid file handle issues
            )
            
            # Perform some operations
            with logger.track_operation("demo_operation"):
                time.sleep(0.05)
                logger.log_debug("Debug message for level test")
                logger.log_warning("Warning message for level test")
            
            # Check what files were created
            log_files = list(Path(temp_dir).glob("*.log"))
            print(f"Log files created: {[f.name for f in log_files]}")
            
            # Cleanup logger
            logger.cleanup()
            time.sleep(0.1)
            
        except Exception as e:
            print(f"Error with {level.value}: {e}")
        
        finally:
            # Manual cleanup
            try:
                import shutil
                if os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir, ignore_errors=True)
            except Exception:
                pass


def main():
    """Main demonstration function."""
    print("AutoWord vNext - Simple Monitoring System Demonstration")
    print("=" * 60)
    
    # Demonstrate performance tracking
    demonstrate_performance_tracking()
    
    # Demonstrate memory monitoring
    demonstrate_memory_monitoring()
    
    # Demonstrate monitoring levels
    demonstrate_monitoring_levels()
    
    print("\n" + "=" * 60)
    print("Simple demonstration completed!")
    print("\nKey features demonstrated:")
    print("✓ Performance tracking with operation and stage metrics")
    print("✓ Memory monitoring with configurable thresholds and alerts")
    print("✓ Different monitoring levels (BASIC, DETAILED, DEBUG, PERFORMANCE)")
    print("✓ Comprehensive statistics and reporting")
    print("\nThe monitoring system provides detailed observability")
    print("into pipeline execution with minimal overhead.")


if __name__ == "__main__":
    main()