"""
Comprehensive error handling and recovery system for AutoWord vNext pipeline.

This module implements centralized error handling, automatic rollback functionality,
NOOP operation logging, security violation detection, and revision handling strategies.
"""

import os
import shutil
import logging
import traceback
from typing import Dict, List, Optional, Any, Callable, Union
from pathlib import Path
from datetime import datetime
from enum import Enum
from dataclasses import dataclass

from .exceptions import (
    VNextError, ExtractionError, PlanningError, ExecutionError, 
    ValidationError, AuditError, RollbackError, SecurityViolationError,
    LocalizationError, WhitelistViolationError
)
from .models import StructureV1, PlanV1, OperationResult


logger = logging.getLogger(__name__)


class ProcessingStatus(Enum):
    """Processing status enumeration."""
    SUCCESS = "SUCCESS"
    FAILED_VALIDATION = "FAILED_VALIDATION"
    ROLLBACK = "ROLLBACK"
    SECURITY_VIOLATION = "SECURITY_VIOLATION"
    NOOP_OPERATIONS = "NOOP_OPERATIONS"


class RevisionHandlingStrategy(Enum):
    """Revision handling strategy enumeration."""
    ACCEPT_ALL = "accept_all"
    REJECT_ALL = "reject_all"
    BYPASS = "bypass"
    FAIL_ON_REVISIONS = "fail_on_revisions"


@dataclass
class ErrorContext:
    """Context information for error handling."""
    pipeline_stage: str
    operation_type: Optional[str] = None
    docx_path: Optional[str] = None
    original_docx_path: Optional[str] = None
    audit_directory: Optional[str] = None
    timestamp: Optional[datetime] = None
    additional_data: Optional[Dict[str, Any]] = None


@dataclass
class RecoveryResult:
    """Result of error recovery operations."""
    success: bool
    status: ProcessingStatus
    rollback_performed: bool
    warnings: List[str]
    errors: List[str]
    recovery_path: Optional[str] = None


class WarningsLogger:
    """Centralized warnings logger for NOOP operations and fallbacks."""
    
    def __init__(self, warnings_log_path: str):
        """
        Initialize warnings logger.
        
        Args:
            warnings_log_path: Path to warnings.log file
        """
        self.warnings_log_path = warnings_log_path
        self._ensure_log_directory()
    
    def _ensure_log_directory(self):
        """Ensure the warnings log directory exists."""
        log_dir = Path(self.warnings_log_path).parent
        log_dir.mkdir(parents=True, exist_ok=True)
    
    def log_noop_operation(self, operation_type: str, reason: str, 
                          operation_data: Optional[Dict[str, Any]] = None):
        """
        Log NOOP operation to warnings.log.
        
        Args:
            operation_type: Type of operation that resulted in NOOP
            reason: Reason why operation was NOOP
            operation_data: Optional operation parameters
        """
        timestamp = datetime.now().isoformat()
        message = f"[{timestamp}] NOOP: {operation_type} - {reason}"
        
        if operation_data:
            message += f" | Data: {operation_data}"
        
        self._write_warning(message)
        logger.warning(f"NOOP operation: {operation_type} - {reason}")
    
    def log_font_fallback(self, original_font: str, fallback_font: str, 
                         fallback_chain: List[str]):
        """
        Log font fallback to warnings.log.
        
        Args:
            original_font: Original font that was not available
            fallback_font: Font that was used as fallback
            fallback_chain: Complete fallback chain attempted
        """
        timestamp = datetime.now().isoformat()
        message = (f"[{timestamp}] FONT_FALLBACK: '{original_font}' -> '{fallback_font}' "
                  f"| Chain: {' -> '.join(fallback_chain)}")
        
        self._write_warning(message)
        logger.warning(f"Font fallback: {original_font} -> {fallback_font}")
    
    def log_localization_fallback(self, original_style: str, fallback_style: str):
        """
        Log style localization fallback to warnings.log.
        
        Args:
            original_style: Original style name that was not found
            fallback_style: Style name used as fallback
        """
        timestamp = datetime.now().isoformat()
        message = f"[{timestamp}] STYLE_FALLBACK: '{original_style}' -> '{fallback_style}'"
        
        self._write_warning(message)
        logger.warning(f"Style fallback: {original_style} -> {fallback_style}")
    
    def log_security_violation(self, operation_type: str, violation_reason: str,
                              security_context: Dict[str, Any]):
        """
        Log security violation to warnings.log.
        
        Args:
            operation_type: Type of operation that violated security
            violation_reason: Reason for security violation
            security_context: Security context information
        """
        timestamp = datetime.now().isoformat()
        message = (f"[{timestamp}] SECURITY_VIOLATION: {operation_type} - {violation_reason} "
                  f"| Context: {security_context}")
        
        self._write_warning(message)
        logger.error(f"Security violation: {operation_type} - {violation_reason}")
    
    def log_revision_handling(self, strategy: RevisionHandlingStrategy, 
                             revision_count: int, action_taken: str):
        """
        Log revision handling to warnings.log.
        
        Args:
            strategy: Revision handling strategy used
            revision_count: Number of revisions found
            action_taken: Action taken for revisions
        """
        timestamp = datetime.now().isoformat()
        message = (f"[{timestamp}] REVISION_HANDLING: Strategy={strategy.value}, "
                  f"Count={revision_count}, Action={action_taken}")
        
        self._write_warning(message)
        logger.info(f"Revision handling: {strategy.value} - {action_taken}")
    
    def _write_warning(self, message: str):
        """Write warning message to warnings.log file."""
        try:
            with open(self.warnings_log_path, 'a', encoding='utf-8') as f:
                f.write(message + '\n')
        except Exception as e:
            logger.error(f"Failed to write to warnings.log: {e}")


class SecurityValidator:
    """Security validator for detecting unauthorized operations."""
    
    # Whitelist of allowed atomic operations
    WHITELISTED_OPERATIONS = {
        'delete_section_by_heading',
        'update_toc',
        'delete_toc',
        'set_style_rule',
        'reassign_paragraphs_to_style',
        'clear_direct_formatting'
    }
    
    # Forbidden patterns that indicate security violations
    FORBIDDEN_PATTERNS = [
        'string_replacement',
        'direct_text_modification',
        'toc_text_replacement',
        'content_injection',
        'macro_execution',
        'external_file_access'
    ]
    
    def __init__(self, warnings_logger: WarningsLogger):
        """
        Initialize security validator.
        
        Args:
            warnings_logger: Warnings logger instance
        """
        self.warnings_logger = warnings_logger
    
    def validate_operation(self, operation_type: str, operation_data: Dict[str, Any]) -> bool:
        """
        Validate if operation is authorized.
        
        Args:
            operation_type: Type of operation to validate
            operation_data: Operation parameters
            
        Returns:
            True if operation is authorized, False otherwise
            
        Raises:
            SecurityViolationError: If unauthorized operation is detected
        """
        # Check if operation type is whitelisted
        if operation_type not in self.WHITELISTED_OPERATIONS:
            violation_reason = f"Operation '{operation_type}' not in whitelist"
            self.warnings_logger.log_security_violation(
                operation_type, violation_reason, operation_data
            )
            raise SecurityViolationError(
                f"Unauthorized operation: {operation_type}",
                operation_type=operation_type,
                security_context={
                    'violation_type': 'non_whitelisted_operation',
                    'operation_data': operation_data
                }
            )
        
        # Check for forbidden patterns in operation data
        for pattern in self.FORBIDDEN_PATTERNS:
            if self._contains_forbidden_pattern(operation_data, pattern):
                violation_reason = f"Operation contains forbidden pattern: {pattern}"
                self.warnings_logger.log_security_violation(
                    operation_type, violation_reason, operation_data
                )
                raise SecurityViolationError(
                    f"Security violation in {operation_type}: {pattern}",
                    operation_type=operation_type,
                    security_context={
                        'violation_type': 'forbidden_pattern',
                        'pattern': pattern,
                        'operation_data': operation_data
                    }
                )
        
        return True
    
    def _contains_forbidden_pattern(self, data: Any, pattern: str) -> bool:
        """
        Check if data contains forbidden pattern.
        
        Args:
            data: Data to check
            pattern: Forbidden pattern to look for
            
        Returns:
            True if pattern is found, False otherwise
        """
        if isinstance(data, str):
            return pattern.lower() in data.lower()
        elif isinstance(data, dict):
            return any(self._contains_forbidden_pattern(v, pattern) for v in data.values())
        elif isinstance(data, list):
            return any(self._contains_forbidden_pattern(item, pattern) for item in data)
        return False


class RollbackManager:
    """Manager for automatic rollback functionality."""
    
    def __init__(self, warnings_logger: WarningsLogger):
        """
        Initialize rollback manager.
        
        Args:
            warnings_logger: Warnings logger instance
        """
        self.warnings_logger = warnings_logger
    
    def perform_rollback(self, original_docx_path: str, modified_docx_path: str,
                        rollback_reason: str) -> RecoveryResult:
        """
        Perform automatic rollback to original document.
        
        Args:
            original_docx_path: Path to original DOCX file
            modified_docx_path: Path to modified DOCX file to rollback
            rollback_reason: Reason for rollback
            
        Returns:
            RecoveryResult with rollback status
        """
        try:
            # Verify original file exists
            if not os.path.exists(original_docx_path):
                raise RollbackError(
                    f"Original DOCX file not found: {original_docx_path}",
                    original_docx=original_docx_path,
                    rollback_reason=rollback_reason
                )
            
            # Create backup of modified file before rollback
            backup_path = None
            if os.path.exists(modified_docx_path):
                backup_path = f"{modified_docx_path}.rollback_backup"
                shutil.copy2(modified_docx_path, backup_path)
                logger.info(f"Created rollback backup: {backup_path}")
            
            # Perform rollback by copying original over modified
            shutil.copy2(original_docx_path, modified_docx_path)
            
            # Log rollback
            timestamp = datetime.now().isoformat()
            message = f"[{timestamp}] ROLLBACK: {rollback_reason} | {original_docx_path} -> {modified_docx_path}"
            self.warnings_logger._write_warning(message)
            
            logger.info(f"Rollback successful: {rollback_reason}")
            
            return RecoveryResult(
                success=True,
                status=ProcessingStatus.ROLLBACK,
                rollback_performed=True,
                warnings=[f"Document rolled back due to: {rollback_reason}"],
                errors=[],
                recovery_path=modified_docx_path
            )
            
        except Exception as e:
            error_msg = f"Rollback failed: {str(e)}"
            logger.error(error_msg)
            
            return RecoveryResult(
                success=False,
                status=ProcessingStatus.ROLLBACK,
                rollback_performed=False,
                warnings=[],
                errors=[error_msg],
                recovery_path=None
            )


class RevisionHandler:
    """Handler for documents with track changes."""
    
    def __init__(self, strategy: RevisionHandlingStrategy, warnings_logger: WarningsLogger):
        """
        Initialize revision handler.
        
        Args:
            strategy: Revision handling strategy
            warnings_logger: Warnings logger instance
        """
        self.strategy = strategy
        self.warnings_logger = warnings_logger
    
    def handle_revisions(self, docx_path: str) -> bool:
        """
        Handle revisions in document according to strategy.
        
        Args:
            docx_path: Path to DOCX file with revisions
            
        Returns:
            True if revisions handled successfully, False otherwise
            
        Raises:
            ExecutionError: If revision handling fails
        """
        try:
            try:
                import win32com.client as win32
                import pythoncom
            except ImportError:
                raise ExecutionError(
                    "Win32 COM libraries not available for revision handling",
                    operation_type="revision_handling",
                    operation_data={"strategy": self.strategy.value}
                )
            
            # Initialize COM
            pythoncom.CoInitialize()
            
            try:
                # Open Word application
                word_app = win32.gencache.EnsureDispatch('Word.Application')
                word_app.Visible = False
                
                # Open document
                doc = word_app.Documents.Open(docx_path)
                
                try:
                    # Check if document has revisions
                    revision_count = doc.Revisions.Count
                    
                    if revision_count == 0:
                        self.warnings_logger.log_revision_handling(
                            self.strategy, 0, "No revisions found"
                        )
                        return True
                    
                    # Handle revisions according to strategy
                    if self.strategy == RevisionHandlingStrategy.ACCEPT_ALL:
                        doc.Revisions.AcceptAll()
                        action_taken = f"Accepted all {revision_count} revisions"
                        
                    elif self.strategy == RevisionHandlingStrategy.REJECT_ALL:
                        doc.Revisions.RejectAll()
                        action_taken = f"Rejected all {revision_count} revisions"
                        
                    elif self.strategy == RevisionHandlingStrategy.BYPASS:
                        action_taken = f"Bypassed {revision_count} revisions (left unchanged)"
                        
                    elif self.strategy == RevisionHandlingStrategy.FAIL_ON_REVISIONS:
                        raise ExecutionError(
                            f"Document contains {revision_count} revisions and strategy is FAIL_ON_REVISIONS",
                            operation_type="revision_handling",
                            operation_data={"revision_count": revision_count, "strategy": self.strategy.value}
                        )
                    
                    # Save document if changes were made
                    if self.strategy in [RevisionHandlingStrategy.ACCEPT_ALL, RevisionHandlingStrategy.REJECT_ALL]:
                        doc.Save()
                    
                    self.warnings_logger.log_revision_handling(
                        self.strategy, revision_count, action_taken
                    )
                    
                    logger.info(f"Revision handling completed: {action_taken}")
                    return True
                    
                finally:
                    doc.Close()
                    
            finally:
                word_app.Quit()
                pythoncom.CoUninitialize()
                
        except Exception as e:
            error_msg = f"Revision handling failed: {str(e)}"
            logger.error(error_msg)
            raise ExecutionError(
                error_msg,
                operation_type="revision_handling",
                operation_data={"strategy": self.strategy.value}
            )


class PipelineErrorHandler:
    """Comprehensive error handler for the entire vNext pipeline."""
    
    def __init__(self, audit_directory: str, 
                 revision_strategy: RevisionHandlingStrategy = RevisionHandlingStrategy.BYPASS):
        """
        Initialize pipeline error handler.
        
        Args:
            audit_directory: Directory for audit files and warnings.log
            revision_strategy: Strategy for handling document revisions
        """
        self.audit_directory = Path(audit_directory)
        self.revision_strategy = revision_strategy
        
        # Ensure audit directory exists
        self.audit_directory.mkdir(parents=True, exist_ok=True)
        
        # Initialize components
        warnings_log_path = self.audit_directory / "warnings.log"
        self.warnings_logger = WarningsLogger(str(warnings_log_path))
        self.security_validator = SecurityValidator(self.warnings_logger)
        self.rollback_manager = RollbackManager(self.warnings_logger)
        self.revision_handler = RevisionHandler(revision_strategy, self.warnings_logger)
        
        # Initialize warnings.log with header
        self._initialize_warnings_log()
    
    def _initialize_warnings_log(self):
        """Initialize warnings.log with header information."""
        timestamp = datetime.now().isoformat()
        header = f"""# AutoWord vNext Warnings Log
# Generated: {timestamp}
# Revision Strategy: {self.revision_strategy.value}
# Audit Directory: {self.audit_directory}

"""
        try:
            with open(self.audit_directory / "warnings.log", 'w', encoding='utf-8') as f:
                f.write(header)
        except Exception as e:
            logger.error(f"Failed to initialize warnings.log: {e}")
    
    def handle_pipeline_error(self, error: Exception, context: ErrorContext) -> RecoveryResult:
        """
        Handle pipeline error with appropriate recovery strategy.
        
        Args:
            error: Exception that occurred
            context: Error context information
            
        Returns:
            RecoveryResult with recovery status and actions taken
        """
        logger.error(f"Pipeline error in {context.pipeline_stage}: {str(error)}")
        logger.debug(f"Error traceback: {traceback.format_exc()}")
        
        # Determine recovery strategy based on error type
        if isinstance(error, SecurityViolationError):
            return self._handle_security_violation(error, context)
        elif isinstance(error, ValidationError):
            return self._handle_validation_error(error, context)
        elif isinstance(error, (ExtractionError, PlanningError, ExecutionError, AuditError)):
            return self._handle_pipeline_stage_error(error, context)
        else:
            return self._handle_unexpected_error(error, context)
    
    def _handle_security_violation(self, error: SecurityViolationError, 
                                  context: ErrorContext) -> RecoveryResult:
        """Handle security violation errors."""
        # Log security violation
        self.warnings_logger.log_security_violation(
            error.details.get('operation_type', 'unknown'),
            str(error),
            error.details.get('security_context', {})
        )
        
        # Perform rollback if modified document exists
        rollback_performed = False
        if context.docx_path and context.original_docx_path:
            rollback_result = self.rollback_manager.perform_rollback(
                context.original_docx_path,
                context.docx_path,
                f"Security violation: {str(error)}"
            )
            rollback_performed = rollback_result.rollback_performed
        
        return RecoveryResult(
            success=False,
            status=ProcessingStatus.SECURITY_VIOLATION,
            rollback_performed=rollback_performed,
            warnings=[f"Security violation detected: {str(error)}"],
            errors=[str(error)],
            recovery_path=context.original_docx_path if rollback_performed else None
        )
    
    def _handle_validation_error(self, error: ValidationError, 
                                context: ErrorContext) -> RecoveryResult:
        """Handle validation errors with automatic rollback."""
        # Perform rollback
        rollback_result = self.rollback_manager.perform_rollback(
            context.original_docx_path or error.details.get('rollback_path'),
            context.docx_path,
            f"Validation failed: {str(error)}"
        )
        
        # Write status file
        self._write_status_file(ProcessingStatus.FAILED_VALIDATION, str(error))
        
        return RecoveryResult(
            success=False,
            status=ProcessingStatus.FAILED_VALIDATION,
            rollback_performed=rollback_result.rollback_performed,
            warnings=rollback_result.warnings + [f"Validation failed: {str(error)}"],
            errors=rollback_result.errors + [str(error)],
            recovery_path=rollback_result.recovery_path
        )
    
    def _handle_pipeline_stage_error(self, error: VNextError, 
                                   context: ErrorContext) -> RecoveryResult:
        """Handle errors from pipeline stages."""
        # Perform rollback if we have paths
        rollback_performed = False
        recovery_path = None
        
        if context.docx_path and context.original_docx_path:
            rollback_result = self.rollback_manager.perform_rollback(
                context.original_docx_path,
                context.docx_path,
                f"{context.pipeline_stage} error: {str(error)}"
            )
            rollback_performed = rollback_result.rollback_performed
            recovery_path = rollback_result.recovery_path
        
        # Write status file
        self._write_status_file(ProcessingStatus.ROLLBACK, str(error))
        
        return RecoveryResult(
            success=False,
            status=ProcessingStatus.ROLLBACK,
            rollback_performed=rollback_performed,
            warnings=[f"{context.pipeline_stage} error: {str(error)}"],
            errors=[str(error)],
            recovery_path=recovery_path
        )
    
    def _handle_unexpected_error(self, error: Exception, 
                               context: ErrorContext) -> RecoveryResult:
        """Handle unexpected errors."""
        # Log unexpected error
        logger.error(f"Unexpected error in {context.pipeline_stage}: {str(error)}")
        logger.debug(f"Unexpected error traceback: {traceback.format_exc()}")
        
        # Perform rollback if possible
        rollback_performed = False
        recovery_path = None
        
        if context.docx_path and context.original_docx_path:
            rollback_result = self.rollback_manager.perform_rollback(
                context.original_docx_path,
                context.docx_path,
                f"Unexpected error: {str(error)}"
            )
            rollback_performed = rollback_result.rollback_performed
            recovery_path = rollback_result.recovery_path
        
        # Write status file
        self._write_status_file(ProcessingStatus.ROLLBACK, f"Unexpected error: {str(error)}")
        
        return RecoveryResult(
            success=False,
            status=ProcessingStatus.ROLLBACK,
            rollback_performed=rollback_performed,
            warnings=[f"Unexpected error in {context.pipeline_stage}: {str(error)}"],
            errors=[str(error)],
            recovery_path=recovery_path
        )
    
    def _write_status_file(self, status: ProcessingStatus, details: str):
        """Write result.status.txt file."""
        try:
            status_file = self.audit_directory / "result.status.txt"
            timestamp = datetime.now().isoformat()
            
            content = f"""Status: {status.value}
Timestamp: {timestamp}
Details: {details}
Revision Strategy: {self.revision_strategy.value}
"""
            
            with open(status_file, 'w', encoding='utf-8') as f:
                f.write(content)
                
            logger.info(f"Status file written: {status.value}")
            
        except Exception as e:
            logger.error(f"Failed to write status file: {e}")
    
    def validate_operation_security(self, operation_type: str, 
                                  operation_data: Dict[str, Any]) -> bool:
        """
        Validate operation security.
        
        Args:
            operation_type: Type of operation
            operation_data: Operation parameters
            
        Returns:
            True if operation is secure, False otherwise
            
        Raises:
            SecurityViolationError: If security violation is detected
        """
        return self.security_validator.validate_operation(operation_type, operation_data)
    
    def handle_document_revisions(self, docx_path: str) -> bool:
        """
        Handle document revisions according to configured strategy.
        
        Args:
            docx_path: Path to document with potential revisions
            
        Returns:
            True if revisions handled successfully, False otherwise
        """
        return self.revision_handler.handle_revisions(docx_path)
    
    def log_noop_operation(self, operation_type: str, reason: str, 
                          operation_data: Optional[Dict[str, Any]] = None):
        """Log NOOP operation."""
        self.warnings_logger.log_noop_operation(operation_type, reason, operation_data)
    
    def log_font_fallback(self, original_font: str, fallback_font: str, 
                         fallback_chain: List[str]):
        """Log font fallback."""
        self.warnings_logger.log_font_fallback(original_font, fallback_font, fallback_chain)
    
    def log_localization_fallback(self, original_style: str, fallback_style: str):
        """Log localization fallback."""
        self.warnings_logger.log_localization_fallback(original_style, fallback_style)