"""
Unit tests for the monitoring system.

Tests comprehensive logging, performance tracking, memory monitoring,
and debug logging functionality.
"""

import os
import sys
import time
import json
import tempfile
import threading
from datetime import datetime, timedelta
from pathlib import Path
from unittest.mock import Mock, patch, MagicMock
import pytest

# Add the parent directory to the path to import vnext modules
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..', '..'))

from autoword.vnext.monitoring import (
    VNextLogger, MemoryMonitor, PerformanceTracker, MonitoringLevel, LogLevel,
    PerformanceMetrics, PipelineStageMetrics, MemoryAlert,
    create_vnext_logger, log_large_document_warning, log_complex_document_scenario
)


class TestMemoryMonitor:
    """Test memory monitoring functionality."""
    
    def test_memory_monitor_initialization(self):
        """Test memory monitor initialization."""
        monitor = MemoryMonitor(
            warning_threshold_mb=512,
            critical_threshold_mb=1024,
            check_interval_seconds=0.5
        )
        
        assert monitor.warning_threshold_mb == 512
        assert monitor.critical_threshold_mb == 1024
        assert monitor.check_interval_seconds == 0.5
        assert not monitor.monitoring_active
        assert monitor.peak_memory_mb == 0.0
        assert len(monitor.alerts) == 0
    
    def test_get_current_memory(self):
        """Test current memory usage retrieval."""
        monitor = MemoryMonitor()
        memory_mb = monitor.get_current_memory_mb()
        
        assert isinstance(memory_mb, float)
        assert memory_mb > 0  # Should have some memory usage
    
    def test_set_current_operation(self):
        """Test setting current operation context."""
        monitor = MemoryMonitor()
        
        monitor.set_current_operation("test_operation")
        assert monitor.current_operation == "test_operation"
        
        monitor.set_current_operation("another_operation")
        assert monitor.current_operation == "another_operation"
    
    def test_start_stop_monitoring(self):
        """Test starting and stopping background monitoring."""
        monitor = MemoryMonitor(check_interval_seconds=0.1)
        
        # Start monitoring
        monitor.start_monitoring()
        assert monitor.monitoring_active
        assert monitor.monitor_thread is not None
        assert monitor.monitor_thread.is_alive()
        
        # Stop monitoring
        monitor.stop_monitoring()
        assert not monitor.monitoring_active
        
        # Wait for thread to finish
        time.sleep(0.2)
        assert not monitor.monitor_thread.is_alive()
    
    def test_memory_alerts(self):
        """Test memory alert generation."""
        monitor = MemoryMonitor(
            warning_threshold_mb=1,  # Very low threshold for testing
            critical_threshold_mb=2,
            check_interval_seconds=0.1
        )
        
        # Mock memory usage to trigger alerts
        with patch.object(monitor, 'get_current_memory_mb', return_value=1.5):
            monitor.start_monitoring()
            monitor.set_current_operation("test_operation")
            
            # Wait for monitoring to detect high memory
            time.sleep(0.3)
            
            monitor.stop_monitoring()
            
            alerts = monitor.get_alerts()
            assert len(alerts) > 0
            
            warning_alerts = [a for a in alerts if a.alert_level == "WARNING"]
            assert len(warning_alerts) > 0
            
            alert = warning_alerts[0]
            assert alert.current_memory_mb == 1.5
            assert alert.operation_name == "test_operation"
            assert "High memory usage" in alert.message
    
    def test_clear_alerts(self):
        """Test clearing memory alerts."""
        monitor = MemoryMonitor()
        
        # Add a mock alert
        alert = MemoryAlert(
            timestamp=datetime.now(),
            current_memory_mb=100.0,
            threshold_mb=50.0,
            operation_name="test",
            alert_level="WARNING",
            message="Test alert"
        )
        monitor.alerts.append(alert)
        
        assert len(monitor.get_alerts()) == 1
        
        monitor.clear_alerts()
        assert len(monitor.get_alerts()) == 0


class TestPerformanceTracker:
    """Test performance tracking functionality."""
    
    def test_performance_tracker_initialization(self):
        """Test performance tracker initialization."""
        tracker = PerformanceTracker()
        
        assert len(tracker.current_metrics) == 0
        assert len(tracker.completed_metrics) == 0
        assert len(tracker.stage_metrics) == 0
        assert len(tracker.completed_stages) == 0
        assert tracker.memory_monitor is None
    
    def test_track_operation_success(self):
        """Test successful operation tracking."""
        tracker = PerformanceTracker()
        
        with tracker.track_operation("test_operation", metadata={"param": "value"}) as metrics:
            time.sleep(0.01)  # Small delay to measure duration
            assert metrics.operation_name == "test_operation"
            assert metrics.metadata["param"] == "value"
            assert metrics.start_time is not None
            assert metrics.end_time is None  # Not completed yet
        
        # After context manager, metrics should be completed
        assert len(tracker.completed_metrics) == 1
        completed = tracker.completed_metrics[0]
        
        assert completed.operation_name == "test_operation"
        assert completed.success is True
        assert completed.end_time is not None
        assert completed.duration_ms is not None
        assert completed.duration_ms > 0
        assert completed.error_message is None
    
    def test_track_operation_failure(self):
        """Test failed operation tracking."""
        tracker = PerformanceTracker()
        
        with pytest.raises(ValueError):
            with tracker.track_operation("failing_operation") as metrics:
                raise ValueError("Test error")
        
        assert len(tracker.completed_metrics) == 1
        completed = tracker.completed_metrics[0]
        
        assert completed.operation_name == "failing_operation"
        assert completed.success is False
        assert completed.error_message == "Test error"
        assert completed.duration_ms is not None
    
    def test_track_stage(self):
        """Test pipeline stage tracking."""
        tracker = PerformanceTracker()
        
        with tracker.track_stage("test_stage") as stage_metrics:
            # Simulate some operations during the stage
            with tracker.track_operation("op1"):
                time.sleep(0.01)
            with tracker.track_operation("op2"):
                time.sleep(0.01)
            
            assert stage_metrics.stage_name == "test_stage"
            assert stage_metrics.start_time is not None
        
        # After stage completion
        assert len(tracker.completed_stages) == 1
        completed_stage = tracker.completed_stages[0]
        
        assert completed_stage.stage_name == "test_stage"
        assert completed_stage.success is True
        assert completed_stage.duration_ms is not None
        assert len(completed_stage.operations) == 2  # Should include ops from the stage
    
    def test_get_operation_stats(self):
        """Test operation statistics generation."""
        tracker = PerformanceTracker()
        
        # Track some operations
        with tracker.track_operation("op1"):
            time.sleep(0.01)
        
        with pytest.raises(ValueError):
            with tracker.track_operation("op2"):
                raise ValueError("Test error")
        
        with tracker.track_operation("op3"):
            time.sleep(0.01)
        
        stats = tracker.get_operation_stats()
        
        assert stats["total_operations"] == 3
        assert stats["successful_operations"] == 2
        assert stats["failed_operations"] == 1
        assert stats["success_rate"] == 2/3
        assert "total_duration_ms" in stats
        assert "average_duration_ms" in stats
    
    def test_get_stage_stats(self):
        """Test stage statistics generation."""
        tracker = PerformanceTracker()
        
        # Track some stages
        with tracker.track_stage("stage1"):
            with tracker.track_operation("op1"):
                time.sleep(0.01)
        
        with tracker.track_stage("stage2"):
            with tracker.track_operation("op2"):
                time.sleep(0.01)
        
        stats = tracker.get_stage_stats()
        
        assert stats["total_stages"] == 2
        assert stats["successful_stages"] == 2
        assert stats["failed_stages"] == 0
        assert stats["success_rate"] == 1.0
        assert "stage_breakdown" in stats
        assert "stage1" in stats["stage_breakdown"]
        assert "stage2" in stats["stage_breakdown"]


class TestVNextLogger:
    """Test comprehensive logging system."""
    
    def test_vnext_logger_initialization(self):
        """Test VNext logger initialization."""
        with tempfile.TemporaryDirectory() as temp_dir:
            logger = VNextLogger(
                audit_directory=temp_dir,
                monitoring_level=MonitoringLevel.DEBUG,
                enable_memory_monitoring=True
            )
            
            assert logger.audit_directory == Path(temp_dir)
            assert logger.monitoring_level == MonitoringLevel.DEBUG
            assert logger.memory_monitor is not None
            assert logger.performance_tracker is not None
            
            # Check that log files are created
            assert (Path(temp_dir) / "pipeline.log").exists()
            assert (Path(temp_dir) / "debug.log").exists()
            assert (Path(temp_dir) / "performance.log").exists()
            
            logger.cleanup()
    
    def test_operation_logging(self):
        """Test operation start/complete logging."""
        with tempfile.TemporaryDirectory() as temp_dir:
            logger = VNextLogger(
                audit_directory=temp_dir,
                monitoring_level=MonitoringLevel.DETAILED
            )
            
            # Test operation logging
            logger.log_operation_start("test_operation", param1="value1", param2="value2")
            logger.log_operation_complete("test_operation", 123.45, success=True, result="success")
            
            # Check that logs were written
            log_file = Path(temp_dir) / "pipeline.log"
            log_content = log_file.read_text(encoding='utf-8')
            
            assert "Starting operation: test_operation" in log_content
            assert "Completed operation: test_operation (SUCCESS) - 123.45ms" in log_content
            
            logger.cleanup()
    
    def test_stage_logging(self):
        """Test pipeline stage logging."""
        with tempfile.TemporaryDirectory() as temp_dir:
            logger = VNextLogger(
                audit_directory=temp_dir,
                monitoring_level=MonitoringLevel.DETAILED
            )
            
            logger.log_stage_start("Extract")
            logger.log_stage_complete("Extract", 1234.56, success=True)
            
            # Check that logs were written
            log_file = Path(temp_dir) / "pipeline.log"
            log_content = log_file.read_text(encoding='utf-8')
            
            assert "Starting pipeline stage: Extract" in log_content
            assert "Completed pipeline stage: Extract (SUCCESS) - 1234.56ms" in log_content
            
            logger.cleanup()
    
    def test_error_logging(self):
        """Test error logging with context."""
        with tempfile.TemporaryDirectory() as temp_dir:
            logger = VNextLogger(
                audit_directory=temp_dir,
                monitoring_level=MonitoringLevel.DEBUG
            )
            
            test_error = ValueError("Test error message")
            context = {"operation": "test_op", "param": "value"}
            
            logger.log_error(test_error, context)
            
            # Check main log
            log_file = Path(temp_dir) / "pipeline.log"
            log_content = log_file.read_text(encoding='utf-8')
            assert "Error: Test error message" in log_content
            assert "Error context: {'operation': 'test_op', 'param': 'value'}" in log_content
            
            # Check debug log for traceback
            debug_file = Path(temp_dir) / "debug.log"
            debug_content = debug_file.read_text(encoding='utf-8')
            assert "Full error traceback:" in debug_content
            
            logger.cleanup()
    
    def test_track_operation_context_manager(self):
        """Test operation tracking context manager."""
        with tempfile.TemporaryDirectory() as temp_dir:
            logger = VNextLogger(
                audit_directory=temp_dir,
                monitoring_level=MonitoringLevel.DETAILED
            )
            
            # Test successful operation
            with logger.track_operation("test_operation", param="value") as metrics:
                time.sleep(0.01)
                assert metrics.operation_name == "test_operation"
            
            # Test failed operation
            with pytest.raises(ValueError):
                with logger.track_operation("failing_operation"):
                    raise ValueError("Test failure")
            
            # Check that both operations were tracked
            stats = logger.performance_tracker.get_operation_stats()
            assert stats["total_operations"] == 2
            assert stats["successful_operations"] == 1
            assert stats["failed_operations"] == 1
            
            logger.cleanup()
    
    def test_track_stage_context_manager(self):
        """Test stage tracking context manager."""
        with tempfile.TemporaryDirectory() as temp_dir:
            logger = VNextLogger(
                audit_directory=temp_dir,
                monitoring_level=MonitoringLevel.DETAILED
            )
            
            with logger.track_stage("test_stage") as stage_metrics:
                with logger.track_operation("op1"):
                    time.sleep(0.01)
                assert stage_metrics.stage_name == "test_stage"
            
            # Check that stage was tracked
            stats = logger.performance_tracker.get_stage_stats()
            assert stats["total_stages"] == 1
            assert stats["successful_stages"] == 1
            assert "test_stage" in stats["stage_breakdown"]
            
            logger.cleanup()
    
    def test_performance_report_generation(self):
        """Test performance report generation and saving."""
        with tempfile.TemporaryDirectory() as temp_dir:
            logger = VNextLogger(
                audit_directory=temp_dir,
                monitoring_level=MonitoringLevel.PERFORMANCE,
                enable_memory_monitoring=True
            )
            
            # Perform some tracked operations
            with logger.track_stage("test_stage"):
                with logger.track_operation("op1"):
                    time.sleep(0.01)
                with logger.track_operation("op2"):
                    time.sleep(0.01)
            
            # Generate report
            report = logger.generate_performance_report()
            
            assert "timestamp" in report
            assert "monitoring_level" in report
            assert report["monitoring_level"] == "PERFORMANCE"
            assert "operation_stats" in report
            assert "stage_stats" in report
            assert "memory_stats" in report
            
            # Save report
            logger.save_performance_report()
            
            report_file = Path(temp_dir) / "performance_report.json"
            assert report_file.exists()
            
            # Load and verify saved report
            with open(report_file, 'r', encoding='utf-8') as f:
                saved_report = json.load(f)
            
            assert saved_report["monitoring_level"] == "PERFORMANCE"
            assert "operation_stats" in saved_report
            
            logger.cleanup()
    
    def test_memory_alert_logging(self):
        """Test memory alert logging."""
        with tempfile.TemporaryDirectory() as temp_dir:
            logger = VNextLogger(
                audit_directory=temp_dir,
                monitoring_level=MonitoringLevel.DEBUG,
                enable_memory_monitoring=False  # Disable automatic monitoring
            )
            
            # Create a mock alert
            alert = MemoryAlert(
                timestamp=datetime.now(),
                current_memory_mb=1024.5,
                threshold_mb=512.0,
                operation_name="test_operation",
                alert_level="WARNING",
                message="High memory usage: 1024.5MB > 512.0MB"
            )
            
            logger.log_memory_alert(alert)
            
            # Check that alert was logged
            log_file = Path(temp_dir) / "pipeline.log"
            log_content = log_file.read_text(encoding='utf-8')
            assert "High memory usage: 1024.5MB > 512.0MB" in log_content
            
            perf_file = Path(temp_dir) / "performance.log"
            perf_content = perf_file.read_text(encoding='utf-8')
            assert "MEMORY_ALERT - WARNING" in perf_content
            
            logger.cleanup()


class TestConvenienceFunctions:
    """Test convenience functions for easy integration."""
    
    def test_create_vnext_logger(self):
        """Test create_vnext_logger convenience function."""
        with tempfile.TemporaryDirectory() as temp_dir:
            logger = create_vnext_logger(
                audit_directory=temp_dir,
                monitoring_level=MonitoringLevel.BASIC,
                enable_memory_monitoring=False
            )
            
            assert isinstance(logger, VNextLogger)
            assert logger.monitoring_level == MonitoringLevel.BASIC
            assert logger.memory_monitor is None
            
            logger.cleanup()
    
    def test_log_large_document_warning(self):
        """Test large document warning logging."""
        with tempfile.TemporaryDirectory() as temp_dir:
            logger = VNextLogger(
                audit_directory=temp_dir,
                monitoring_level=MonitoringLevel.DEBUG
            )
            
            # Test warning for large document
            log_large_document_warning(logger, 75.5, 5000, threshold_mb=50)
            
            # Check that warning was logged
            log_file = Path(temp_dir) / "pipeline.log"
            log_content = log_file.read_text(encoding='utf-8')
            assert "Processing large document: 75.5MB with 5000 paragraphs" in log_content
            
            # Test no warning for small document
            log_large_document_warning(logger, 25.0, 1000, threshold_mb=50)
            
            logger.cleanup()
    
    def test_log_complex_document_scenario(self):
        """Test complex document scenario logging."""
        with tempfile.TemporaryDirectory() as temp_dir:
            logger = VNextLogger(
                audit_directory=temp_dir,
                monitoring_level=MonitoringLevel.DEBUG
            )
            
            scenario_details = {
                "table_count": 25,
                "complexity_level": "high",
                "recommendation": "Use careful processing"
            }
            
            log_complex_document_scenario(logger, "many_tables", scenario_details)
            
            # Check that scenario was logged to debug log
            debug_file = Path(temp_dir) / "debug.log"
            debug_content = debug_file.read_text(encoding='utf-8')
            assert "Complex document scenario detected: many_tables" in debug_content
            
            logger.cleanup()


class TestIntegrationScenarios:
    """Test integration scenarios with realistic usage patterns."""
    
    def test_full_pipeline_monitoring_simulation(self):
        """Test full pipeline monitoring simulation."""
        with tempfile.TemporaryDirectory() as temp_dir:
            logger = VNextLogger(
                audit_directory=temp_dir,
                monitoring_level=MonitoringLevel.DETAILED,
                enable_memory_monitoring=True,
                memory_warning_threshold_mb=1,  # Low threshold for testing
                memory_critical_threshold_mb=2
            )
            
            # Simulate full pipeline execution
            with logger.track_stage("Extract"):
                with logger.track_operation("extractor_initialization"):
                    time.sleep(0.01)
                with logger.track_operation("structure_extraction"):
                    time.sleep(0.02)
                with logger.track_operation("inventory_extraction"):
                    time.sleep(0.01)
            
            with logger.track_stage("Plan"):
                with logger.track_operation("planner_initialization"):
                    time.sleep(0.01)
                with logger.track_operation("llm_plan_generation"):
                    time.sleep(0.05)  # Longer operation
                with logger.track_operation("plan_validation"):
                    time.sleep(0.01)
            
            with logger.track_stage("Execute"):
                with logger.track_operation("executor_initialization"):
                    time.sleep(0.01)
                with logger.track_operation("atomic_operations_execution"):
                    time.sleep(0.03)
            
            with logger.track_stage("Validate"):
                with logger.track_operation("validator_initialization"):
                    time.sleep(0.01)
                with logger.track_operation("validation_assertions"):
                    time.sleep(0.02)
            
            with logger.track_stage("Audit"):
                with logger.track_operation("audit_snapshots"):
                    time.sleep(0.01)
                with logger.track_operation("diff_report_generation"):
                    time.sleep(0.01)
                with logger.track_operation("audit_status_writing"):
                    time.sleep(0.01)
            
            # Generate final report
            report = logger.generate_performance_report()
            
            # Verify comprehensive tracking
            assert report["stage_stats"]["total_stages"] == 5
            assert report["operation_stats"]["total_operations"] == 13
            assert report["stage_stats"]["successful_stages"] == 5
            assert report["operation_stats"]["successful_operations"] == 13
            
            # Verify stage breakdown
            stage_breakdown = report["stage_stats"]["stage_breakdown"]
            expected_stages = ["Extract", "Plan", "Execute", "Validate", "Audit"]
            for stage in expected_stages:
                assert stage in stage_breakdown
                assert stage_breakdown[stage]["success"] is True
                assert stage_breakdown[stage]["duration_ms"] > 0
            
            logger.cleanup()
    
    def test_error_handling_monitoring(self):
        """Test monitoring during error scenarios."""
        with tempfile.TemporaryDirectory() as temp_dir:
            logger = VNextLogger(
                audit_directory=temp_dir,
                monitoring_level=MonitoringLevel.DEBUG
            )
            
            # Simulate pipeline with errors
            with logger.track_stage("Extract"):
                with logger.track_operation("successful_op"):
                    time.sleep(0.01)
                
                # Failed operation
                with pytest.raises(ValueError):
                    with logger.track_operation("failing_op"):
                        raise ValueError("Simulated extraction error")
            
            # Stage should still complete (even with failed operations)
            stats = logger.performance_tracker.get_stage_stats()
            assert stats["total_stages"] == 1
            assert stats["successful_stages"] == 1  # Stage completed despite operation failure
            
            op_stats = logger.performance_tracker.get_operation_stats()
            assert op_stats["total_operations"] == 2
            assert op_stats["successful_operations"] == 1
            assert op_stats["failed_operations"] == 1
            assert op_stats["success_rate"] == 0.5
            
            logger.cleanup()


if __name__ == "__main__":
    pytest.main([__file__, "-v"])