"""
Comprehensive logging and monitoring system for AutoWord vNext pipeline.

This module provides detailed operation logging, performance monitoring, debug logging,
memory usage monitoring, and execution time tracking for all pipeline stages and
atomic operations.
"""

import os
import sys
import time
import psutil
import logging
import threading
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Any, Callable, Union
from pathlib import Path
from contextlib import contextmanager
from dataclasses import dataclass, field
from enum import Enum
import json
import traceback


class LogLevel(str, Enum):
    """Log level enumeration."""
    DEBUG = "DEBUG"
    INFO = "INFO"
    WARNING = "WARNING"
    ERROR = "ERROR"
    CRITICAL = "CRITICAL"


class MonitoringLevel(str, Enum):
    """Monitoring level enumeration."""
    BASIC = "BASIC"      # Essential operations only
    DETAILED = "DETAILED"  # All operations with timing
    DEBUG = "DEBUG"      # Full debug with memory tracking
    PERFORMANCE = "PERFORMANCE"  # Performance optimization focus


@dataclass
class PerformanceMetrics:
    """Performance metrics for operations."""
    operation_name: str
    start_time: datetime
    end_time: Optional[datetime] = None
    duration_ms: Optional[float] = None
    memory_before_mb: Optional[float] = None
    memory_after_mb: Optional[float] = None
    memory_delta_mb: Optional[float] = None
    cpu_percent: Optional[float] = None
    success: bool = True
    error_message: Optional[str] = None
    metadata: Dict[str, Any] = field(default_factory=dict)


@dataclass
class PipelineStageMetrics:
    """Metrics for entire pipeline stage."""
    stage_name: str
    start_time: datetime
    end_time: Optional[datetime] = None
    duration_ms: Optional[float] = None
    operations: List[PerformanceMetrics] = field(default_factory=list)
    memory_peak_mb: Optional[float] = None
    memory_average_mb: Optional[float] = None
    success: bool = True
    error_message: Optional[str] = None


@dataclass
class MemoryAlert:
    """Memory usage alert."""
    timestamp: datetime
    current_memory_mb: float
    threshold_mb: float
    operation_name: str
    alert_level: str  # WARNING, CRITICAL
    message: str


class MemoryMonitor:
    """Memory usage monitoring with alerts."""
    
    def __init__(self, warning_threshold_mb: float = 1024, 
                 critical_threshold_mb: float = 2048,
                 check_interval_seconds: float = 1.0):
        """
        Initialize memory monitor.
        
        Args:
            warning_threshold_mb: Memory threshold for warnings (MB)
            critical_threshold_mb: Memory threshold for critical alerts (MB)
            check_interval_seconds: Monitoring check interval
        """
        self.warning_threshold_mb = warning_threshold_mb
        self.critical_threshold_mb = critical_threshold_mb
        self.check_interval_seconds = check_interval_seconds
        
        self.process = psutil.Process()
        self.alerts: List[MemoryAlert] = []
        self.peak_memory_mb = 0.0
        self.monitoring_active = False
        self.monitor_thread: Optional[threading.Thread] = None
        self.current_operation = "Unknown"
        
        self._lock = threading.Lock()
    
    def start_monitoring(self):
        """Start background memory monitoring."""
        if self.monitoring_active:
            return
            
        self.monitoring_active = True
        self.monitor_thread = threading.Thread(target=self._monitor_loop, daemon=True)
        self.monitor_thread.start()
    
    def stop_monitoring(self):
        """Stop background memory monitoring."""
        self.monitoring_active = False
        if self.monitor_thread:
            self.monitor_thread.join(timeout=2.0)
    
    def set_current_operation(self, operation_name: str):
        """Set current operation name for context."""
        with self._lock:
            self.current_operation = operation_name
    
    def get_current_memory_mb(self) -> float:
        """Get current memory usage in MB."""
        try:
            memory_info = self.process.memory_info()
            return memory_info.rss / (1024 * 1024)  # Convert to MB
        except Exception:
            return 0.0
    
    def get_peak_memory_mb(self) -> float:
        """Get peak memory usage since monitoring started."""
        return self.peak_memory_mb
    
    def get_alerts(self) -> List[MemoryAlert]:
        """Get all memory alerts."""
        with self._lock:
            return self.alerts.copy()
    
    def clear_alerts(self):
        """Clear all memory alerts."""
        with self._lock:
            self.alerts.clear()
    
    def _monitor_loop(self):
        """Background monitoring loop."""
        while self.monitoring_active:
            try:
                current_memory = self.get_current_memory_mb()
                
                with self._lock:
                    # Update peak memory
                    if current_memory > self.peak_memory_mb:
                        self.peak_memory_mb = current_memory
                    
                    # Check thresholds
                    if current_memory > self.critical_threshold_mb:
                        alert = MemoryAlert(
                            timestamp=datetime.now(),
                            current_memory_mb=current_memory,
                            threshold_mb=self.critical_threshold_mb,
                            operation_name=self.current_operation,
                            alert_level="CRITICAL",
                            message=f"Critical memory usage: {current_memory:.1f}MB > {self.critical_threshold_mb}MB"
                        )
                        self.alerts.append(alert)
                        
                    elif current_memory > self.warning_threshold_mb:
                        # Only add warning if we don't already have recent warnings
                        recent_warnings = [a for a in self.alerts 
                                         if a.alert_level == "WARNING" and 
                                         (datetime.now() - a.timestamp).seconds < 30]
                        if not recent_warnings:
                            alert = MemoryAlert(
                                timestamp=datetime.now(),
                                current_memory_mb=current_memory,
                                threshold_mb=self.warning_threshold_mb,
                                operation_name=self.current_operation,
                                alert_level="WARNING",
                                message=f"High memory usage: {current_memory:.1f}MB > {self.warning_threshold_mb}MB"
                            )
                            self.alerts.append(alert)
                
                time.sleep(self.check_interval_seconds)
                
            except Exception:
                # Silently continue monitoring even if there are errors
                time.sleep(self.check_interval_seconds)


class PerformanceTracker:
    """Performance tracking for operations and pipeline stages."""
    
    def __init__(self, memory_monitor: Optional[MemoryMonitor] = None):
        """
        Initialize performance tracker.
        
        Args:
            memory_monitor: Optional memory monitor for memory tracking
        """
        self.memory_monitor = memory_monitor
        self.current_metrics: Dict[str, PerformanceMetrics] = {}
        self.completed_metrics: List[PerformanceMetrics] = []
        self.stage_metrics: Dict[str, PipelineStageMetrics] = {}
        self.completed_stages: List[PipelineStageMetrics] = []
        
        self._lock = threading.Lock()
    
    @contextmanager
    def track_operation(self, operation_name: str, metadata: Optional[Dict[str, Any]] = None):
        """
        Context manager for tracking operation performance.
        
        Args:
            operation_name: Name of the operation being tracked
            metadata: Optional metadata to include with metrics
        """
        metrics = PerformanceMetrics(
            operation_name=operation_name,
            start_time=datetime.now(),
            metadata=metadata or {}
        )
        
        # Set memory monitor context
        if self.memory_monitor:
            self.memory_monitor.set_current_operation(operation_name)
            metrics.memory_before_mb = self.memory_monitor.get_current_memory_mb()
        
        # Get CPU usage before
        try:
            cpu_percent = psutil.cpu_percent(interval=None)
        except Exception:
            cpu_percent = None
        
        with self._lock:
            self.current_metrics[operation_name] = metrics
        
        try:
            yield metrics
            metrics.success = True
            
        except Exception as e:
            metrics.success = False
            metrics.error_message = str(e)
            raise
            
        finally:
            # Complete metrics
            metrics.end_time = datetime.now()
            metrics.duration_ms = (metrics.end_time - metrics.start_time).total_seconds() * 1000
            
            if self.memory_monitor:
                metrics.memory_after_mb = self.memory_monitor.get_current_memory_mb()
                if metrics.memory_before_mb is not None:
                    metrics.memory_delta_mb = metrics.memory_after_mb - metrics.memory_before_mb
            
            try:
                metrics.cpu_percent = psutil.cpu_percent(interval=None)
            except Exception:
                pass
            
            with self._lock:
                if operation_name in self.current_metrics:
                    del self.current_metrics[operation_name]
                self.completed_metrics.append(metrics)
    
    @contextmanager
    def track_stage(self, stage_name: str):
        """
        Context manager for tracking pipeline stage performance.
        
        Args:
            stage_name: Name of the pipeline stage
        """
        stage_metrics = PipelineStageMetrics(
            stage_name=stage_name,
            start_time=datetime.now()
        )
        
        with self._lock:
            self.stage_metrics[stage_name] = stage_metrics
        
        # Track memory during stage
        memory_samples = []
        
        try:
            yield stage_metrics
            stage_metrics.success = True
            
        except Exception as e:
            stage_metrics.success = False
            stage_metrics.error_message = str(e)
            raise
            
        finally:
            # Complete stage metrics
            stage_metrics.end_time = datetime.now()
            stage_metrics.duration_ms = (stage_metrics.end_time - stage_metrics.start_time).total_seconds() * 1000
            
            # Collect operations that occurred during this stage
            stage_start = stage_metrics.start_time
            stage_end = stage_metrics.end_time or datetime.now()
            
            with self._lock:
                stage_operations = [
                    m for m in self.completed_metrics
                    if stage_start <= m.start_time <= stage_end
                ]
                stage_metrics.operations = stage_operations
                
                if stage_name in self.stage_metrics:
                    del self.stage_metrics[stage_name]
                self.completed_stages.append(stage_metrics)
            
            # Calculate memory statistics
            if self.memory_monitor:
                stage_metrics.memory_peak_mb = self.memory_monitor.get_peak_memory_mb()
                
                memory_values = [op.memory_after_mb for op in stage_operations 
                               if op.memory_after_mb is not None]
                if memory_values:
                    stage_metrics.memory_average_mb = sum(memory_values) / len(memory_values)
    
    def get_operation_stats(self) -> Dict[str, Any]:
        """Get operation performance statistics."""
        with self._lock:
            operations = self.completed_metrics.copy()
        
        if not operations:
            return {"total_operations": 0}
        
        successful_ops = [op for op in operations if op.success]
        failed_ops = [op for op in operations if not op.success]
        
        durations = [op.duration_ms for op in operations if op.duration_ms is not None]
        memory_deltas = [op.memory_delta_mb for op in operations 
                        if op.memory_delta_mb is not None]
        
        stats = {
            "total_operations": len(operations),
            "successful_operations": len(successful_ops),
            "failed_operations": len(failed_ops),
            "success_rate": len(successful_ops) / len(operations) if operations else 0,
        }
        
        if durations:
            stats.update({
                "total_duration_ms": sum(durations),
                "average_duration_ms": sum(durations) / len(durations),
                "min_duration_ms": min(durations),
                "max_duration_ms": max(durations),
            })
        
        if memory_deltas:
            stats.update({
                "total_memory_delta_mb": sum(memory_deltas),
                "average_memory_delta_mb": sum(memory_deltas) / len(memory_deltas),
                "max_memory_delta_mb": max(memory_deltas),
            })
        
        return stats
    
    def get_stage_stats(self) -> Dict[str, Any]:
        """Get pipeline stage performance statistics."""
        with self._lock:
            stages = self.completed_stages.copy()
        
        if not stages:
            return {"total_stages": 0}
        
        successful_stages = [stage for stage in stages if stage.success]
        failed_stages = [stage for stage in stages if not stage.success]
        
        durations = [stage.duration_ms for stage in stages if stage.duration_ms is not None]
        
        stats = {
            "total_stages": len(stages),
            "successful_stages": len(successful_stages),
            "failed_stages": len(failed_stages),
            "success_rate": len(successful_stages) / len(stages) if stages else 0,
        }
        
        if durations:
            stats.update({
                "total_pipeline_duration_ms": sum(durations),
                "average_stage_duration_ms": sum(durations) / len(durations),
                "min_stage_duration_ms": min(durations),
                "max_stage_duration_ms": max(durations),
            })
        
        # Stage breakdown
        stage_breakdown = {}
        for stage in stages:
            stage_breakdown[stage.stage_name] = {
                "duration_ms": stage.duration_ms,
                "operations_count": len(stage.operations),
                "success": stage.success,
                "memory_peak_mb": stage.memory_peak_mb,
                "memory_average_mb": stage.memory_average_mb,
            }
        
        stats["stage_breakdown"] = stage_breakdown
        
        return stats


class VNextLogger:
    """Comprehensive logging system for vNext pipeline."""
    
    def __init__(self, 
                 audit_directory: str,
                 monitoring_level: MonitoringLevel = MonitoringLevel.DETAILED,
                 console_level: LogLevel = LogLevel.INFO,
                 file_level: LogLevel = LogLevel.DEBUG,
                 enable_memory_monitoring: bool = True,
                 memory_warning_threshold_mb: float = 1024,
                 memory_critical_threshold_mb: float = 2048):
        """
        Initialize vNext logger.
        
        Args:
            audit_directory: Directory for log files
            monitoring_level: Level of monitoring detail
            console_level: Console logging level
            file_level: File logging level
            enable_memory_monitoring: Whether to enable memory monitoring
            memory_warning_threshold_mb: Memory warning threshold
            memory_critical_threshold_mb: Memory critical threshold
        """
        self.audit_directory = Path(audit_directory)
        self.monitoring_level = monitoring_level
        self.console_level = console_level
        self.file_level = file_level
        
        # Create audit directory
        self.audit_directory.mkdir(parents=True, exist_ok=True)
        
        # Initialize memory monitor
        self.memory_monitor = None
        if enable_memory_monitoring:
            self.memory_monitor = MemoryMonitor(
                warning_threshold_mb=memory_warning_threshold_mb,
                critical_threshold_mb=memory_critical_threshold_mb
            )
        
        # Initialize performance tracker
        self.performance_tracker = PerformanceTracker(self.memory_monitor)
        
        # Setup loggers
        self._setup_loggers()
        
        # Start monitoring
        if self.memory_monitor:
            self.memory_monitor.start_monitoring()
    
    def _setup_loggers(self):
        """Setup logging configuration."""
        # Main logger
        self.logger = logging.getLogger("vnext_pipeline")
        self.logger.setLevel(logging.DEBUG)
        self.logger.handlers.clear()  # Clear any existing handlers
        
        # Console handler
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setLevel(getattr(logging, self.console_level.value))
        console_formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
        console_handler.setFormatter(console_formatter)
        self.logger.addHandler(console_handler)
        
        # File handler for general logs
        log_file = self.audit_directory / "pipeline.log"
        file_handler = logging.FileHandler(log_file, mode='w', encoding='utf-8')
        file_handler.setLevel(getattr(logging, self.file_level.value))
        file_formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(funcName)s:%(lineno)d - %(message)s'
        )
        file_handler.setFormatter(file_formatter)
        self.logger.addHandler(file_handler)
        
        # Debug logger for detailed debugging
        self.debug_logger = logging.getLogger("vnext_debug")
        self.debug_logger.setLevel(logging.DEBUG)
        self.debug_logger.handlers.clear()
        
        debug_file = self.audit_directory / "debug.log"
        debug_handler = logging.FileHandler(debug_file, mode='w', encoding='utf-8')
        debug_handler.setLevel(logging.DEBUG)
        debug_formatter = logging.Formatter(
            '%(asctime)s - %(levelname)s - %(funcName)s:%(lineno)d - %(message)s\n'
            'Thread: %(thread)d - Process: %(process)d\n'
            '%(pathname)s\n'
        )
        debug_handler.setFormatter(debug_formatter)
        self.debug_logger.addHandler(debug_handler)
        
        # Performance logger
        self.perf_logger = logging.getLogger("vnext_performance")
        self.perf_logger.setLevel(logging.INFO)
        self.perf_logger.handlers.clear()
        
        perf_file = self.audit_directory / "performance.log"
        perf_handler = logging.FileHandler(perf_file, mode='w', encoding='utf-8')
        perf_handler.setLevel(logging.INFO)
        perf_formatter = logging.Formatter(
            '%(asctime)s - PERF - %(message)s'
        )
        perf_handler.setFormatter(perf_formatter)
        self.perf_logger.addHandler(perf_handler)
    
    def log_operation_start(self, operation_name: str, **kwargs):
        """Log the start of an operation."""
        if self.monitoring_level in [MonitoringLevel.DETAILED, MonitoringLevel.DEBUG, MonitoringLevel.PERFORMANCE]:
            self.logger.info(f"Starting operation: {operation_name}")
            
            if kwargs:
                self.debug_logger.debug(f"Operation parameters: {kwargs}")
    
    def log_operation_complete(self, operation_name: str, duration_ms: float, success: bool = True, **kwargs):
        """Log the completion of an operation."""
        if self.monitoring_level in [MonitoringLevel.DETAILED, MonitoringLevel.DEBUG, MonitoringLevel.PERFORMANCE]:
            status = "SUCCESS" if success else "FAILED"
            self.logger.info(f"Completed operation: {operation_name} ({status}) - {duration_ms:.2f}ms")
            
            if self.monitoring_level == MonitoringLevel.PERFORMANCE:
                self.perf_logger.info(f"{operation_name}: {duration_ms:.2f}ms - {status}")
            
            if kwargs:
                self.debug_logger.debug(f"Operation results: {kwargs}")
    
    def log_stage_start(self, stage_name: str):
        """Log the start of a pipeline stage."""
        self.logger.info(f"=== Starting pipeline stage: {stage_name} ===")
        
        if self.memory_monitor:
            current_memory = self.memory_monitor.get_current_memory_mb()
            self.perf_logger.info(f"Stage {stage_name} - Memory at start: {current_memory:.1f}MB")
    
    def log_stage_complete(self, stage_name: str, duration_ms: float, success: bool = True):
        """Log the completion of a pipeline stage."""
        status = "SUCCESS" if success else "FAILED"
        self.logger.info(f"=== Completed pipeline stage: {stage_name} ({status}) - {duration_ms:.2f}ms ===")
        
        if self.memory_monitor:
            current_memory = self.memory_monitor.get_current_memory_mb()
            peak_memory = self.memory_monitor.get_peak_memory_mb()
            self.perf_logger.info(f"Stage {stage_name} - Memory at end: {current_memory:.1f}MB, Peak: {peak_memory:.1f}MB")
    
    def log_error(self, error: Exception, context: Optional[Dict[str, Any]] = None):
        """Log an error with full context."""
        self.logger.error(f"Error: {str(error)}")
        
        if context:
            self.logger.error(f"Error context: {context}")
        
        # Full traceback to debug log
        self.debug_logger.error(f"Full error traceback:\n{traceback.format_exc()}")
    
    def log_warning(self, message: str, context: Optional[Dict[str, Any]] = None):
        """Log a warning message."""
        self.logger.warning(message)
        
        if context and self.monitoring_level == MonitoringLevel.DEBUG:
            self.debug_logger.warning(f"Warning context: {context}")
    
    def log_debug(self, message: str, **kwargs):
        """Log debug information."""
        if self.monitoring_level == MonitoringLevel.DEBUG:
            self.debug_logger.debug(message)
            
            if kwargs:
                self.debug_logger.debug(f"Debug data: {kwargs}")
    
    def log_memory_alert(self, alert: MemoryAlert):
        """Log a memory alert."""
        level = logging.WARNING if alert.alert_level == "WARNING" else logging.CRITICAL
        self.logger.log(level, alert.message)
        self.perf_logger.info(f"MEMORY_ALERT - {alert.alert_level}: {alert.message}")
    
    @contextmanager
    def track_operation(self, operation_name: str, **kwargs):
        """Context manager for tracking and logging operations."""
        self.log_operation_start(operation_name, **kwargs)
        
        with self.performance_tracker.track_operation(operation_name, kwargs) as metrics:
            try:
                yield metrics
                self.log_operation_complete(operation_name, metrics.duration_ms or 0, True)
                
            except Exception as e:
                self.log_operation_complete(operation_name, metrics.duration_ms or 0, False)
                self.log_error(e, {"operation": operation_name, "parameters": kwargs})
                raise
    
    @contextmanager
    def track_stage(self, stage_name: str):
        """Context manager for tracking and logging pipeline stages."""
        self.log_stage_start(stage_name)
        
        with self.performance_tracker.track_stage(stage_name) as stage_metrics:
            try:
                yield stage_metrics
                self.log_stage_complete(stage_name, stage_metrics.duration_ms or 0, True)
                
            except Exception as e:
                self.log_stage_complete(stage_name, stage_metrics.duration_ms or 0, False)
                self.log_error(e, {"stage": stage_name})
                raise
    
    def generate_performance_report(self) -> Dict[str, Any]:
        """Generate comprehensive performance report."""
        report = {
            "timestamp": datetime.now().isoformat(),
            "monitoring_level": self.monitoring_level.value,
            "operation_stats": self.performance_tracker.get_operation_stats(),
            "stage_stats": self.performance_tracker.get_stage_stats(),
        }
        
        if self.memory_monitor:
            report["memory_stats"] = {
                "peak_memory_mb": self.memory_monitor.get_peak_memory_mb(),
                "current_memory_mb": self.memory_monitor.get_current_memory_mb(),
                "warning_threshold_mb": self.memory_monitor.warning_threshold_mb,
                "critical_threshold_mb": self.memory_monitor.critical_threshold_mb,
                "alerts": [
                    {
                        "timestamp": alert.timestamp.isoformat(),
                        "level": alert.alert_level,
                        "memory_mb": alert.current_memory_mb,
                        "operation": alert.operation_name,
                        "message": alert.message
                    }
                    for alert in self.memory_monitor.get_alerts()
                ]
            }
        
        return report
    
    def save_performance_report(self):
        """Save performance report to file."""
        report = self.generate_performance_report()
        
        report_file = self.audit_directory / "performance_report.json"
        with open(report_file, 'w', encoding='utf-8') as f:
            json.dump(report, f, indent=2, ensure_ascii=False)
        
        self.logger.info(f"Performance report saved to: {report_file}")
    
    def cleanup(self):
        """Cleanup monitoring resources."""
        if self.memory_monitor:
            self.memory_monitor.stop_monitoring()
        
        # Save final performance report
        try:
            self.save_performance_report()
        except Exception as e:
            self.logger.error(f"Failed to save performance report: {e}")
        
        # Log memory alerts if any
        if self.memory_monitor:
            alerts = self.memory_monitor.get_alerts()
            if alerts:
                self.logger.warning(f"Total memory alerts during execution: {len(alerts)}")
                for alert in alerts[-5:]:  # Log last 5 alerts
                    self.log_memory_alert(alert)
        
        # Close all file handlers to release file locks
        try:
            for handler in self.logger.handlers[:]:
                if hasattr(handler, 'close'):
                    handler.close()
                self.logger.removeHandler(handler)
            
            for handler in self.debug_logger.handlers[:]:
                if hasattr(handler, 'close'):
                    handler.close()
                self.debug_logger.removeHandler(handler)
            
            for handler in self.perf_logger.handlers[:]:
                if hasattr(handler, 'close'):
                    handler.close()
                self.perf_logger.removeHandler(handler)
        except Exception as e:
            # Don't raise exceptions during cleanup
            pass


# Convenience functions for easy integration

def create_vnext_logger(audit_directory: str, 
                       monitoring_level: MonitoringLevel = MonitoringLevel.DETAILED,
                       **kwargs) -> VNextLogger:
    """
    Create a vNext logger with default configuration.
    
    Args:
        audit_directory: Directory for log files
        monitoring_level: Level of monitoring detail
        **kwargs: Additional logger configuration
        
    Returns:
        Configured VNextLogger instance
    """
    return VNextLogger(
        audit_directory=audit_directory,
        monitoring_level=monitoring_level,
        **kwargs
    )


def log_large_document_warning(logger: VNextLogger, document_size_mb: float, 
                              paragraph_count: int, threshold_mb: float = 50):
    """
    Log warning for large document processing.
    
    Args:
        logger: VNext logger instance
        document_size_mb: Document size in MB
        paragraph_count: Number of paragraphs
        threshold_mb: Size threshold for warning
    """
    if document_size_mb > threshold_mb:
        logger.log_warning(
            f"Processing large document: {document_size_mb:.1f}MB with {paragraph_count} paragraphs",
            {
                "document_size_mb": document_size_mb,
                "paragraph_count": paragraph_count,
                "threshold_mb": threshold_mb,
                "recommendation": "Consider monitoring memory usage closely"
            }
        )


def log_complex_document_scenario(logger: VNextLogger, scenario_type: str, 
                                 details: Dict[str, Any]):
    """
    Log complex document scenario for troubleshooting.
    
    Args:
        logger: VNext logger instance
        scenario_type: Type of complex scenario
        details: Scenario details
    """
    logger.log_debug(
        f"Complex document scenario detected: {scenario_type}",
        scenario_type=scenario_type,
        **details
    )