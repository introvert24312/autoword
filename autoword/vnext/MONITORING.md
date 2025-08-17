# AutoWord vNext - Comprehensive Logging and Monitoring

This document describes the comprehensive logging and monitoring system implemented for AutoWord vNext pipeline. The system provides detailed operation logging, performance monitoring, debug logging, memory usage monitoring, and execution time tracking.

## Overview

The monitoring system consists of several key components:

- **VNextLogger**: Main logging coordinator with configurable detail levels
- **PerformanceTracker**: Tracks operation and pipeline stage performance metrics
- **MemoryMonitor**: Monitors memory usage with configurable alerts
- **MonitoringLevel**: Configurable monitoring detail levels

## Key Features

### ✅ Detailed Operation Logging
- Logs all pipeline operations with timestamps and context
- Configurable log levels (DEBUG, INFO, WARNING, ERROR, CRITICAL)
- Separate log files for different purposes (pipeline.log, debug.log, performance.log)

### ✅ Performance Monitoring
- Tracks execution time for each operation and pipeline stage
- Provides comprehensive statistics (total, average, min, max durations)
- Memory usage tracking with before/after measurements
- CPU usage monitoring during operations

### ✅ Debug Logging
- Detailed debug information for troubleshooting complex scenarios
- Full stack traces for errors
- Context-aware logging with operation parameters
- Complex document scenario detection and logging

### ✅ Memory Usage Monitoring
- Real-time memory usage tracking with configurable thresholds
- Warning and critical alerts for high memory usage
- Peak memory usage tracking
- Background monitoring with minimal overhead

### ✅ Execution Time Tracking
- Precise timing for each atomic operation
- Pipeline stage duration tracking
- Performance bottleneck identification
- Comprehensive performance reports

## Usage

### Basic Usage

```python
from autoword.vnext.monitoring import create_vnext_logger, MonitoringLevel

# Create logger with default settings
logger = create_vnext_logger(
    audit_directory="./logs",
    monitoring_level=MonitoringLevel.DETAILED
)

# Track an operation
with logger.track_operation("document_processing") as metrics:
    # Your operation code here
    process_document()

# Track a pipeline stage
with logger.track_stage("Extract") as stage_metrics:
    # Multiple operations in this stage
    with logger.track_operation("structure_extraction"):
        extract_structure()
    
    with logger.track_operation("inventory_extraction"):
        extract_inventory()

# Cleanup when done
logger.cleanup()
```

### Advanced Configuration

```python
from autoword.vnext.monitoring import VNextLogger, MonitoringLevel, LogLevel

logger = VNextLogger(
    audit_directory="./detailed_logs",
    monitoring_level=MonitoringLevel.DEBUG,
    console_level=LogLevel.INFO,
    file_level=LogLevel.DEBUG,
    enable_memory_monitoring=True,
    memory_warning_threshold_mb=1024,
    memory_critical_threshold_mb=2048
)
```

### Integration with Pipeline

The monitoring system is fully integrated into the VNextPipeline:

```python
from autoword.vnext.pipeline import VNextPipeline
from autoword.vnext.monitoring import MonitoringLevel

pipeline = VNextPipeline(
    monitoring_level=MonitoringLevel.PERFORMANCE,
    enable_memory_monitoring=True,
    memory_warning_threshold_mb=512,
    memory_critical_threshold_mb=1024
)

result = pipeline.process_document("document.docx", "user intent")
```

## Monitoring Levels

### BASIC
- Essential operations only
- Minimal logging overhead
- Basic error reporting

### DETAILED
- All operations with timing
- Comprehensive operation logging
- Performance metrics collection

### DEBUG
- Full debug information
- Memory usage tracking
- Detailed error context
- Complex scenario detection

### PERFORMANCE
- Optimization-focused monitoring
- Detailed performance metrics
- Memory and CPU tracking
- Performance bottleneck identification

## Log Files

The monitoring system creates several log files:

### pipeline.log
- Main pipeline operations
- Stage transitions
- Operation start/complete messages
- Warnings and errors

### debug.log
- Detailed debug information
- Full stack traces
- Operation parameters and context
- Complex document scenarios

### performance.log
- Performance-focused logging
- Operation timing details
- Memory usage alerts
- Performance optimization data

### performance_report.json
- Comprehensive performance report
- Operation and stage statistics
- Memory usage summary
- Detailed metrics in JSON format

## Memory Monitoring

The memory monitoring system provides:

### Real-time Monitoring
```python
# Memory monitor runs in background thread
monitor = MemoryMonitor(
    warning_threshold_mb=1024,
    critical_threshold_mb=2048,
    check_interval_seconds=1.0
)

monitor.start_monitoring()
# ... your operations ...
monitor.stop_monitoring()

# Check for alerts
alerts = monitor.get_alerts()
peak_memory = monitor.get_peak_memory_mb()
```

### Memory Alerts
- **WARNING**: Memory usage exceeds warning threshold
- **CRITICAL**: Memory usage exceeds critical threshold
- Alerts include timestamp, operation context, and recommendations

## Performance Tracking

### Operation Metrics
- Operation name and type
- Start/end timestamps
- Duration in milliseconds
- Memory usage before/after
- Success/failure status
- Error messages for failed operations

### Stage Metrics
- Pipeline stage name
- Total stage duration
- Operations within the stage
- Peak and average memory usage
- Success/failure status

### Statistics
```python
# Get operation statistics
op_stats = tracker.get_operation_stats()
print(f"Success rate: {op_stats['success_rate']:.1%}")
print(f"Average duration: {op_stats['average_duration_ms']:.1f}ms")

# Get stage statistics
stage_stats = tracker.get_stage_stats()
for stage_name, stage_data in stage_stats['stage_breakdown'].items():
    print(f"{stage_name}: {stage_data['duration_ms']:.1f}ms")
```

## Error Handling and Logging

### Comprehensive Error Context
```python
try:
    # Operation that might fail
    risky_operation()
except Exception as e:
    logger.log_error(e, {
        "operation": "risky_operation",
        "parameters": {"param1": "value1"},
        "context": "additional context"
    })
```

### Automatic Error Tracking
- Failed operations are automatically tracked
- Full stack traces saved to debug log
- Error context preserved for troubleshooting
- Performance impact of failures measured

## Complex Document Scenarios

The system automatically detects and logs complex document scenarios:

```python
# Automatically detected scenarios
log_complex_document_scenario(logger, "many_tables", {
    "table_count": 25,
    "complexity": "high",
    "recommendation": "Use careful processing"
})

# Large document warnings
log_large_document_warning(logger, 75.5, 5000, threshold_mb=50)
```

## Best Practices

### 1. Choose Appropriate Monitoring Level
- Use **BASIC** for production with minimal overhead
- Use **DETAILED** for development and testing
- Use **DEBUG** for troubleshooting complex issues
- Use **PERFORMANCE** for optimization work

### 2. Configure Memory Thresholds
- Set warning threshold to 70-80% of available memory
- Set critical threshold to 90% of available memory
- Monitor alerts and adjust thresholds based on usage patterns

### 3. Regular Cleanup
- Always call `logger.cleanup()` when done
- Use context managers for automatic cleanup
- Monitor log file sizes in production

### 4. Performance Considerations
- Higher monitoring levels have more overhead
- Memory monitoring uses background thread
- File I/O for logging can impact performance
- Consider log rotation for long-running processes

## Troubleshooting

### Common Issues

#### File Handle Leaks
```python
# Always cleanup properly
try:
    logger = create_vnext_logger(...)
    # ... use logger ...
finally:
    logger.cleanup()
```

#### Memory Alerts
- Check for memory leaks in operations
- Increase thresholds if alerts are false positives
- Monitor peak memory usage patterns
- Consider processing documents in batches

#### Performance Issues
- Use appropriate monitoring level
- Disable memory monitoring if not needed
- Check for I/O bottlenecks in logging
- Monitor log file sizes

### Debug Information

The debug log contains detailed information for troubleshooting:
- Full operation parameters
- Memory usage at each step
- Complex document characteristics
- Error stack traces with context

## Examples

See the following example files for detailed usage:
- `examples/simple_monitoring_demo.py` - Basic usage examples
- `examples/monitoring_demo.py` - Advanced usage scenarios
- `tests/test_monitoring.py` - Comprehensive test cases

## Requirements

The monitoring system requires:
- `psutil>=5.8.0` for memory and CPU monitoring
- `pydantic>=2.0.0` for data validation
- Standard library modules: `logging`, `threading`, `time`, `json`

## Integration with Requirements

This monitoring system addresses the following requirements:

- **8.1**: Comprehensive error handling with detailed logging
- **8.2**: NOOP operation logging and error recovery
- **8.5**: Detailed error context for troubleshooting

The system provides complete observability into the vNext pipeline execution, enabling effective monitoring, debugging, and performance optimization.