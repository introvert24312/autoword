"""
Exception classes for AutoWord vNext pipeline.

This module defines the exception hierarchy for all vNext pipeline errors,
providing structured error handling with detailed context information.
"""

from typing import List, Optional, Dict, Any


class VNextError(Exception):
    """Base exception class for all vNext pipeline errors."""
    
    def __init__(self, message: str, details: Optional[Dict[str, Any]] = None):
        """
        Initialize vNext error.
        
        Args:
            message: Human-readable error message
            details: Optional dictionary with additional error context
        """
        super().__init__(message)
        self.message = message
        self.details = details or {}
        
    def __str__(self) -> str:
        """Return string representation of the error."""
        if self.details:
            return f"{self.message} (Details: {self.details})"
        return self.message


class ExtractionError(VNextError):
    """Exception raised during document structure extraction."""
    
    def __init__(self, message: str, docx_path: Optional[str] = None, 
                 extraction_stage: Optional[str] = None, **kwargs):
        """
        Initialize extraction error.
        
        Args:
            message: Error message
            docx_path: Path to the DOCX file being processed
            extraction_stage: Stage where extraction failed (structure/inventory)
            **kwargs: Additional error details
        """
        details = kwargs
        if docx_path:
            details['docx_path'] = docx_path
        if extraction_stage:
            details['extraction_stage'] = extraction_stage
            
        super().__init__(message, details)


class PlanningError(VNextError):
    """Exception raised during LLM plan generation or validation."""
    
    def __init__(self, message: str, plan_data: Optional[Dict[str, Any]] = None,
                 validation_errors: Optional[List[str]] = None, 
                 llm_response: Optional[str] = None, **kwargs):
        """
        Initialize planning error.
        
        Args:
            message: Error message
            plan_data: Invalid plan data that caused the error
            validation_errors: List of schema validation errors
            llm_response: Raw LLM response that caused the error
            **kwargs: Additional error details
        """
        details = kwargs
        if plan_data:
            details['plan_data'] = plan_data
        if validation_errors:
            details['validation_errors'] = validation_errors
        if llm_response:
            details['llm_response'] = llm_response[:1000]  # Truncate for readability
            
        super().__init__(message, details)


class ExecutionError(VNextError):
    """Exception raised during atomic operation execution."""
    
    def __init__(self, message: str, operation_type: Optional[str] = None,
                 operation_data: Optional[Dict[str, Any]] = None,
                 com_error: Optional[Exception] = None, **kwargs):
        """
        Initialize execution error.
        
        Args:
            message: Error message
            operation_type: Type of atomic operation that failed
            operation_data: Operation parameters that caused the error
            com_error: Underlying COM exception if applicable
            **kwargs: Additional error details
        """
        details = kwargs
        if operation_type:
            details['operation_type'] = operation_type
        if operation_data:
            details['operation_data'] = operation_data
        if com_error:
            details['com_error'] = str(com_error)
            
        super().__init__(message, details)


class ValidationError(VNextError):
    """Exception raised during document validation."""
    
    def __init__(self, message: str, assertion_failures: Optional[List[str]] = None,
                 rollback_path: Optional[str] = None, 
                 validation_stage: Optional[str] = None, **kwargs):
        """
        Initialize validation error.
        
        Args:
            message: Error message
            assertion_failures: List of failed validation assertions
            rollback_path: Path to rollback DOCX file
            validation_stage: Stage where validation failed
            **kwargs: Additional error details
        """
        details = kwargs
        if assertion_failures:
            details['assertion_failures'] = assertion_failures
        if rollback_path:
            details['rollback_path'] = rollback_path
        if validation_stage:
            details['validation_stage'] = validation_stage
            
        super().__init__(message, details)


class AuditError(VNextError):
    """Exception raised during audit trail generation."""
    
    def __init__(self, message: str, audit_directory: Optional[str] = None,
                 audit_stage: Optional[str] = None, **kwargs):
        """
        Initialize audit error.
        
        Args:
            message: Error message
            audit_directory: Path to audit directory
            audit_stage: Stage where audit failed
            **kwargs: Additional error details
        """
        details = kwargs
        if audit_directory:
            details['audit_directory'] = audit_directory
        if audit_stage:
            details['audit_stage'] = audit_stage
            
        super().__init__(message, details)


class SchemaValidationError(PlanningError):
    """Exception raised when JSON schema validation fails."""
    
    def __init__(self, message: str, schema_name: str, 
                 validation_errors: List[str], invalid_data: Dict[str, Any]):
        """
        Initialize schema validation error.
        
        Args:
            message: Error message
            schema_name: Name of the schema that failed validation
            validation_errors: List of specific validation errors
            invalid_data: Data that failed validation
        """
        super().__init__(
            message,
            validation_errors=validation_errors,
            schema_name=schema_name,
            invalid_data=invalid_data
        )


class WhitelistViolationError(PlanningError):
    """Exception raised when plan contains non-whitelisted operations."""
    
    def __init__(self, message: str, invalid_operations: List[str],
                 plan_data: Dict[str, Any]):
        """
        Initialize whitelist violation error.
        
        Args:
            message: Error message
            invalid_operations: List of invalid operation types
            plan_data: Plan data containing invalid operations
        """
        super().__init__(
            message,
            plan_data=plan_data,
            invalid_operations=invalid_operations
        )


class LocalizationError(ExecutionError):
    """Exception raised during localization and font fallback."""
    
    def __init__(self, message: str, style_name: Optional[str] = None,
                 font_name: Optional[str] = None, fallback_chain: Optional[List[str]] = None):
        """
        Initialize localization error.
        
        Args:
            message: Error message
            style_name: Style name that failed localization
            font_name: Font name that failed fallback
            fallback_chain: Font fallback chain that was attempted
        """
        super().__init__(
            message,
            style_name=style_name,
            font_name=font_name,
            fallback_chain=fallback_chain
        )


class RollbackError(ValidationError):
    """Exception raised when document rollback fails."""
    
    def __init__(self, message: str, original_docx: str, 
                 rollback_reason: str, rollback_exception: Optional[Exception] = None):
        """
        Initialize rollback error.
        
        Args:
            message: Error message
            original_docx: Path to original DOCX file
            rollback_reason: Reason for rollback attempt
            rollback_exception: Exception that occurred during rollback
        """
        super().__init__(
            message,
            original_docx=original_docx,
            rollback_reason=rollback_reason,
            rollback_exception=str(rollback_exception) if rollback_exception else None
        )


class SecurityViolationError(ExecutionError):
    """Exception raised when unauthorized operations are attempted."""
    
    def __init__(self, message: str, operation_type: str, 
                 security_context: Dict[str, Any]):
        """
        Initialize security violation error.
        
        Args:
            message: Error message
            operation_type: Type of operation that violated security
            security_context: Security context information
        """
        super().__init__(
            message,
            operation_type=operation_type,
            security_context=security_context
        )