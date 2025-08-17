"""
Tests for the error handling and recovery system.

This module tests comprehensive exception handling, automatic rollback,
NOOP operation logging, security violation detection, and revision handling.
"""

import os
import tempfile
import shutil
import pytest
from pathlib import Path
from unittest.mock import Mock, patch, MagicMock
from datetime import datetime

from autoword.vnext.error_handler import (
    PipelineErrorHandler, WarningsLogger, SecurityValidator, RollbackManager,
    RevisionHandler, ProcessingStatus, RevisionHandlingStrategy, ErrorContext,
    RecoveryResult
)
from autoword.vnext.exceptions import (
    SecurityViolationError, ValidationError, ExecutionError, RollbackError
)


class TestWarningsLogger:
    """Test warnings logger functionality."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.temp_dir = tempfile.mkdtemp()
        self.warnings_log_path = os.path.join(self.temp_dir, "warnings.log")
        self.logger = WarningsLogger(self.warnings_log_path)
    
    def teardown_method(self):
        """Clean up test fixtures."""
        shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    def test_log_noop_operation(self):
        """Test NOOP operation logging."""
        operation_type = "delete_section_by_heading"
        reason = "Heading not found"
        operation_data = {"heading_text": "Missing Section", "level": 1}
        
        self.logger.log_noop_operation(operation_type, reason, operation_data)
        
        # Verify log file was created and contains expected content
        assert os.path.exists(self.warnings_log_path)
        
        with open(self.warnings_log_path, 'r', encoding='utf-8') as f:
            content = f.read()
            assert "NOOP: delete_section_by_heading - Heading not found" in content
            assert "Missing Section" in content
    
    def test_log_font_fallback(self):
        """Test font fallback logging."""
        original_font = "楷体"
        fallback_font = "楷体_GB2312"
        fallback_chain = ["楷体", "楷体_GB2312", "STKaiti"]
        
        self.logger.log_font_fallback(original_font, fallback_font, fallback_chain)
        
        with open(self.warnings_log_path, 'r', encoding='utf-8') as f:
            content = f.read()
            assert "FONT_FALLBACK: '楷体' -> '楷体_GB2312'" in content
            assert "楷体 -> 楷体_GB2312 -> STKaiti" in content
    
    def test_log_localization_fallback(self):
        """Test localization fallback logging."""
        original_style = "Heading 1"
        fallback_style = "标题 1"
        
        self.logger.log_localization_fallback(original_style, fallback_style)
        
        with open(self.warnings_log_path, 'r', encoding='utf-8') as f:
            content = f.read()
            assert "STYLE_FALLBACK: 'Heading 1' -> '标题 1'" in content
    
    def test_log_security_violation(self):
        """Test security violation logging."""
        operation_type = "unauthorized_operation"
        violation_reason = "Not in whitelist"
        security_context = {"user": "test", "operation_data": {"param": "value"}}
        
        self.logger.log_security_violation(operation_type, violation_reason, security_context)
        
        with open(self.warnings_log_path, 'r', encoding='utf-8') as f:
            content = f.read()
            assert "SECURITY_VIOLATION: unauthorized_operation - Not in whitelist" in content
            assert "test" in content
    
    def test_log_revision_handling(self):
        """Test revision handling logging."""
        strategy = RevisionHandlingStrategy.ACCEPT_ALL
        revision_count = 5
        action_taken = "Accepted all 5 revisions"
        
        self.logger.log_revision_handling(strategy, revision_count, action_taken)
        
        with open(self.warnings_log_path, 'r', encoding='utf-8') as f:
            content = f.read()
            assert "REVISION_HANDLING: Strategy=accept_all" in content
            assert "Count=5" in content
            assert "Action=Accepted all 5 revisions" in content


class TestSecurityValidator:
    """Test security validator functionality."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.temp_dir = tempfile.mkdtemp()
        self.warnings_log_path = os.path.join(self.temp_dir, "warnings.log")
        self.warnings_logger = WarningsLogger(self.warnings_log_path)
        self.validator = SecurityValidator(self.warnings_logger)
    
    def teardown_method(self):
        """Clean up test fixtures."""
        shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    def test_validate_whitelisted_operation(self):
        """Test validation of whitelisted operations."""
        operation_type = "delete_section_by_heading"
        operation_data = {"heading_text": "Test", "level": 1}
        
        # Should not raise exception
        result = self.validator.validate_operation(operation_type, operation_data)
        assert result is True
    
    def test_validate_non_whitelisted_operation(self):
        """Test validation of non-whitelisted operations."""
        operation_type = "unauthorized_operation"
        operation_data = {"param": "value"}
        
        with pytest.raises(SecurityViolationError) as exc_info:
            self.validator.validate_operation(operation_type, operation_data)
        
        assert "Unauthorized operation" in str(exc_info.value)
        assert exc_info.value.details['operation_type'] == operation_type
    
    def test_validate_forbidden_pattern(self):
        """Test detection of forbidden patterns."""
        operation_type = "set_style_rule"
        operation_data = {
            "target_style": "Normal",
            "action": "string_replacement",  # Forbidden pattern
            "content": "test"
        }
        
        with pytest.raises(SecurityViolationError) as exc_info:
            self.validator.validate_operation(operation_type, operation_data)
        
        assert "Security violation" in str(exc_info.value)
        assert "string_replacement" in str(exc_info.value)
    
    def test_contains_forbidden_pattern_nested(self):
        """Test detection of forbidden patterns in nested data."""
        data = {
            "level1": {
                "level2": ["item1", "DIRECT_TEXT_MODIFICATION", "item3"]
            }
        }
        
        result = self.validator._contains_forbidden_pattern(data, "direct_text_modification")
        assert result is True
    
    def test_contains_forbidden_pattern_not_found(self):
        """Test when forbidden pattern is not found."""
        data = {
            "operation": "delete_section_by_heading",
            "params": {"heading": "Test"}
        }
        
        result = self.validator._contains_forbidden_pattern(data, "string_replacement")
        assert result is False


class TestRollbackManager:
    """Test rollback manager functionality."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.temp_dir = tempfile.mkdtemp()
        self.warnings_log_path = os.path.join(self.temp_dir, "warnings.log")
        self.warnings_logger = WarningsLogger(self.warnings_log_path)
        self.manager = RollbackManager(self.warnings_logger)
        
        # Create test files
        self.original_docx = os.path.join(self.temp_dir, "original.docx")
        self.modified_docx = os.path.join(self.temp_dir, "modified.docx")
        
        with open(self.original_docx, 'w') as f:
            f.write("original content")
        with open(self.modified_docx, 'w') as f:
            f.write("modified content")
    
    def teardown_method(self):
        """Clean up test fixtures."""
        shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    def test_perform_rollback_success(self):
        """Test successful rollback operation."""
        rollback_reason = "Validation failed"
        
        result = self.manager.perform_rollback(
            self.original_docx, self.modified_docx, rollback_reason
        )
        
        assert result.success is True
        assert result.status == ProcessingStatus.ROLLBACK
        assert result.rollback_performed is True
        assert rollback_reason in result.warnings[0]
        
        # Verify file was rolled back
        with open(self.modified_docx, 'r') as f:
            content = f.read()
            assert content == "original content"
        
        # Verify backup was created
        backup_path = f"{self.modified_docx}.rollback_backup"
        assert os.path.exists(backup_path)
        with open(backup_path, 'r') as f:
            backup_content = f.read()
            assert backup_content == "modified content"
    
    def test_perform_rollback_missing_original(self):
        """Test rollback when original file is missing."""
        missing_original = os.path.join(self.temp_dir, "missing.docx")
        rollback_reason = "Test rollback"
        
        result = self.manager.perform_rollback(
            missing_original, self.modified_docx, rollback_reason
        )
        
        assert result.success is False
        assert result.rollback_performed is False
        assert "not found" in result.errors[0]
    
    def test_perform_rollback_no_modified_file(self):
        """Test rollback when modified file doesn't exist."""
        missing_modified = os.path.join(self.temp_dir, "missing_modified.docx")
        rollback_reason = "Test rollback"
        
        result = self.manager.perform_rollback(
            self.original_docx, missing_modified, rollback_reason
        )
        
        assert result.success is True
        assert result.rollback_performed is True
        
        # Verify file was created with original content
        assert os.path.exists(missing_modified)
        with open(missing_modified, 'r') as f:
            content = f.read()
            assert content == "original content"


class TestRevisionHandler:
    """Test revision handler functionality."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.temp_dir = tempfile.mkdtemp()
        self.warnings_log_path = os.path.join(self.temp_dir, "warnings.log")
        self.warnings_logger = WarningsLogger(self.warnings_log_path)
    
    def teardown_method(self):
        """Clean up test fixtures."""
        shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    def test_handle_revisions_accept_all(self):
        """Test handling revisions with ACCEPT_ALL strategy."""
        handler = RevisionHandler(RevisionHandlingStrategy.ACCEPT_ALL, self.warnings_logger)
        
        # Test strategy property
        assert handler.strategy == RevisionHandlingStrategy.ACCEPT_ALL
        
        # Test that handler is properly initialized
        assert handler.warnings_logger == self.warnings_logger
    
    def test_handle_revisions_reject_all(self):
        """Test handling revisions with REJECT_ALL strategy."""
        handler = RevisionHandler(RevisionHandlingStrategy.REJECT_ALL, self.warnings_logger)
        
        # Test strategy property
        assert handler.strategy == RevisionHandlingStrategy.REJECT_ALL
        
        # Test that handler is properly initialized
        assert handler.warnings_logger == self.warnings_logger
    
    def test_handle_revisions_bypass(self):
        """Test handling revisions with BYPASS strategy."""
        handler = RevisionHandler(RevisionHandlingStrategy.BYPASS, self.warnings_logger)
        
        # Test strategy property
        assert handler.strategy == RevisionHandlingStrategy.BYPASS
        
        # Test that handler is properly initialized
        assert handler.warnings_logger == self.warnings_logger
    
    def test_handle_revisions_fail_on_revisions(self):
        """Test handling revisions with FAIL_ON_REVISIONS strategy."""
        handler = RevisionHandler(RevisionHandlingStrategy.FAIL_ON_REVISIONS, self.warnings_logger)
        
        # Test strategy property
        assert handler.strategy == RevisionHandlingStrategy.FAIL_ON_REVISIONS
        
        # Test that handler is properly initialized
        assert handler.warnings_logger == self.warnings_logger
    
    def test_handle_revisions_no_revisions(self):
        """Test handling document with no revisions."""
        handler = RevisionHandler(RevisionHandlingStrategy.ACCEPT_ALL, self.warnings_logger)
        
        # Test strategy property
        assert handler.strategy == RevisionHandlingStrategy.ACCEPT_ALL
        
        # Test that handler is properly initialized
        assert handler.warnings_logger == self.warnings_logger


class TestPipelineErrorHandler:
    """Test pipeline error handler functionality."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.temp_dir = tempfile.mkdtemp()
        self.audit_dir = os.path.join(self.temp_dir, "audit")
        self.handler = PipelineErrorHandler(
            self.audit_dir, 
            RevisionHandlingStrategy.BYPASS
        )
        
        # Create test files
        self.original_docx = os.path.join(self.temp_dir, "original.docx")
        self.modified_docx = os.path.join(self.temp_dir, "modified.docx")
        
        with open(self.original_docx, 'w') as f:
            f.write("original content")
        with open(self.modified_docx, 'w') as f:
            f.write("modified content")
    
    def teardown_method(self):
        """Clean up test fixtures."""
        shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    def test_initialization(self):
        """Test error handler initialization."""
        assert os.path.exists(self.audit_dir)
        assert os.path.exists(os.path.join(self.audit_dir, "warnings.log"))
        
        # Check warnings.log header
        with open(os.path.join(self.audit_dir, "warnings.log"), 'r', encoding='utf-8') as f:
            content = f.read()
            assert "AutoWord vNext Warnings Log" in content
            assert "bypass" in content.lower()
    
    def test_handle_security_violation(self):
        """Test handling security violation errors."""
        error = SecurityViolationError(
            "Unauthorized operation",
            operation_type="bad_operation",
            security_context={"violation": "test"}
        )
        
        context = ErrorContext(
            pipeline_stage="executor",
            docx_path=self.modified_docx,
            original_docx_path=self.original_docx
        )
        
        result = self.handler.handle_pipeline_error(error, context)
        
        assert result.success is False
        assert result.status == ProcessingStatus.SECURITY_VIOLATION
        assert result.rollback_performed is True
        assert "Security violation" in result.warnings[0]
    
    def test_handle_validation_error(self):
        """Test handling validation errors."""
        error = ValidationError(
            "Validation failed",
            assertion_failures=["Chapter assertion failed"],
            rollback_path=self.original_docx
        )
        
        context = ErrorContext(
            pipeline_stage="validator",
            docx_path=self.modified_docx,
            original_docx_path=self.original_docx
        )
        
        result = self.handler.handle_pipeline_error(error, context)
        
        assert result.success is False
        assert result.status == ProcessingStatus.FAILED_VALIDATION
        assert result.rollback_performed is True
        
        # Check status file was written
        status_file = os.path.join(self.audit_dir, "result.status.txt")
        assert os.path.exists(status_file)
        
        with open(status_file, 'r', encoding='utf-8') as f:
            content = f.read()
            assert "FAILED_VALIDATION" in content
            assert "Validation failed" in content
    
    def test_handle_execution_error(self):
        """Test handling execution errors."""
        error = ExecutionError(
            "Execution failed",
            operation_type="delete_section_by_heading",
            operation_data={"heading": "Test"}
        )
        
        context = ErrorContext(
            pipeline_stage="executor",
            docx_path=self.modified_docx,
            original_docx_path=self.original_docx
        )
        
        result = self.handler.handle_pipeline_error(error, context)
        
        assert result.success is False
        assert result.status == ProcessingStatus.ROLLBACK
        assert result.rollback_performed is True
    
    def test_handle_unexpected_error(self):
        """Test handling unexpected errors."""
        error = ValueError("Unexpected error occurred")
        
        context = ErrorContext(
            pipeline_stage="planner",
            docx_path=self.modified_docx,
            original_docx_path=self.original_docx
        )
        
        result = self.handler.handle_pipeline_error(error, context)
        
        assert result.success is False
        assert result.status == ProcessingStatus.ROLLBACK
        assert result.rollback_performed is True
        assert "Unexpected error" in result.warnings[0]
    
    def test_validate_operation_security(self):
        """Test operation security validation."""
        # Valid operation
        result = self.handler.validate_operation_security(
            "delete_section_by_heading",
            {"heading_text": "Test", "level": 1}
        )
        assert result is True
        
        # Invalid operation
        with pytest.raises(SecurityViolationError):
            self.handler.validate_operation_security(
                "unauthorized_operation",
                {"param": "value"}
            )
    
    def test_log_noop_operation(self):
        """Test NOOP operation logging."""
        self.handler.log_noop_operation(
            "delete_section_by_heading",
            "Heading not found",
            {"heading_text": "Missing"}
        )
        
        warnings_log = os.path.join(self.audit_dir, "warnings.log")
        with open(warnings_log, 'r', encoding='utf-8') as f:
            content = f.read()
            assert "NOOP: delete_section_by_heading" in content
            assert "Heading not found" in content
    
    def test_log_font_fallback(self):
        """Test font fallback logging."""
        self.handler.log_font_fallback(
            "楷体", "楷体_GB2312", ["楷体", "楷体_GB2312", "STKaiti"]
        )
        
        warnings_log = os.path.join(self.audit_dir, "warnings.log")
        with open(warnings_log, 'r', encoding='utf-8') as f:
            content = f.read()
            assert "FONT_FALLBACK" in content
            assert "楷体" in content
    
    def test_log_localization_fallback(self):
        """Test localization fallback logging."""
        self.handler.log_localization_fallback("Heading 1", "标题 1")
        
        warnings_log = os.path.join(self.audit_dir, "warnings.log")
        with open(warnings_log, 'r', encoding='utf-8') as f:
            content = f.read()
            assert "STYLE_FALLBACK" in content
            assert "Heading 1" in content


if __name__ == "__main__":
    pytest.main([__file__])