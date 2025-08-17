"""
Runtime constraint enforcement system for AutoWord vNext.

This module implements comprehensive runtime constraints to ensure security and safety:
- Whitelist operation validation that rejects non-approved operations
- String replacement prevention for content and TOC display text modifications
- Word object layer modifications only (no direct text manipulation)
- LLM output validation with JSON schema gateway
- User input sanitization and parameter validation for all operations
"""

import re
import json
import logging
from typing import Dict, List, Optional, Any, Union, Set, Tuple
from pathlib import Path

from .models import (
    PlanV1, AtomicOperationUnion, ValidationResult, ProcessingResult,
    DeleteSectionByHeading, UpdateToc, DeleteToc, SetStyleRule,
    ReassignParagraphsToStyle, ClearDirectFormatting
)
from .schema_validator import SchemaValidator
from .exceptions import SecurityViolationError, SchemaValidationError


logger = logging.getLogger(__name__)


class RuntimeConstraintEnforcer:
    """Enforce runtime constraints for security and safety."""
    
    # Whitelisted operation types - only these are allowed
    WHITELISTED_OPERATIONS: Set[str] = {
        "delete_section_by_heading",
        "update_toc",
        "delete_toc", 
        "set_style_rule",
        "reassign_paragraphs_to_style",
        "clear_direct_formatting"
    }
    
    # Forbidden patterns that indicate string replacement attempts
    FORBIDDEN_STRING_PATTERNS: List[str] = [
        r'\.replace\s*\(',  # .replace() method calls
        r'\.sub\s*\(',      # re.sub() calls
        r'\.gsub\s*\(',     # gsub calls
        r'str\.replace',    # string replace functions
        r'text\s*=\s*["\'].*["\']',  # Direct text assignment
        r'\.Text\s*=',      # Word COM Text property assignment
        r'\.Value\s*=.*["\']',  # Field value assignment with strings
    ]
    
    # Forbidden Word COM properties/methods for direct text manipulation
    FORBIDDEN_WORD_OPERATIONS: Set[str] = {
        'Text',           # Direct text property access
        'Value',          # Field value direct assignment
        'Result',         # Field result direct assignment
        'InsertAfter',    # Text insertion methods
        'InsertBefore',
        'TypeText',
        'TypeParagraph',
        'Paste',
        'PasteSpecial',
    }
    
    # Required authorization operations
    AUTHORIZATION_REQUIRED_OPS: Set[str] = {
        "clear_direct_formatting"
    }
    
    def __init__(self, schema_validator: Optional[SchemaValidator] = None):
        """
        Initialize runtime constraint enforcer.
        
        Args:
            schema_validator: Schema validator instance (creates new if None)
        """
        self.schema_validator = schema_validator or SchemaValidator()
        self._violation_log: List[Dict[str, Any]] = []
    
    def validate_plan_constraints(self, plan: PlanV1) -> ValidationResult:
        """
        Validate plan against all runtime constraints.
        
        Args:
            plan: Execution plan to validate
            
        Returns:
            ValidationResult with detailed constraint violations
        """
        errors = []
        warnings = []
        
        try:
            # 1. Whitelist operation validation
            whitelist_result = self._validate_whitelist_operations(plan)
            errors.extend(whitelist_result.errors)
            warnings.extend(whitelist_result.warnings)
            
            # 2. String replacement prevention
            string_result = self._validate_no_string_replacement(plan)
            errors.extend(string_result.errors)
            warnings.extend(string_result.warnings)
            
            # 3. Word object layer enforcement
            object_result = self._validate_word_object_layer_only(plan)
            errors.extend(object_result.errors)
            warnings.extend(object_result.warnings)
            
            # 4. Authorization requirements
            auth_result = self._validate_authorization_requirements(plan)
            errors.extend(auth_result.errors)
            warnings.extend(auth_result.warnings)
            
            # 5. Parameter validation and sanitization
            param_result = self._validate_and_sanitize_parameters(plan)
            errors.extend(param_result.errors)
            warnings.extend(param_result.warnings)
            
            # Log violations if any
            if errors:
                self._log_security_violation("plan_validation", {
                    "plan_ops_count": len(plan.ops),
                    "violations": errors
                })
            
            return ValidationResult(
                is_valid=len(errors) == 0,
                errors=errors,
                warnings=warnings
            )
            
        except Exception as e:
            logger.error(f"Constraint validation failed: {str(e)}")
            return ValidationResult(
                is_valid=False,
                errors=[f"Constraint validation error: {str(e)}"],
                warnings=warnings
            )
    
    def validate_llm_output(self, llm_output: Union[str, Dict[str, Any]]) -> ValidationResult:
        """
        Validate LLM output with JSON schema gateway.
        
        Args:
            llm_output: Raw LLM output (string or parsed dict)
            
        Returns:
            ValidationResult with schema validation results
        """
        errors = []
        warnings = []
        
        try:
            # Parse JSON if string
            if isinstance(llm_output, str):
                try:
                    parsed_output = json.loads(llm_output)
                except json.JSONDecodeError as e:
                    return ValidationResult(
                        is_valid=False,
                        errors=[f"Invalid JSON from LLM: {str(e)}"],
                        warnings=[]
                    )
            else:
                parsed_output = llm_output
            
            # Validate against plan.v1 schema
            schema_result = self.schema_validator.validate_plan_v1(parsed_output)
            errors.extend(schema_result.errors)
            warnings.extend(schema_result.warnings)
            
            # Additional LLM-specific validations
            if schema_result.is_valid:
                try:
                    plan = PlanV1(**parsed_output)
                    constraint_result = self.validate_plan_constraints(plan)
                    errors.extend(constraint_result.errors)
                    warnings.extend(constraint_result.warnings)
                except Exception as e:
                    errors.append(f"Failed to create plan from LLM output: {str(e)}")
            
            # Check for suspicious content that might indicate DOCX/OOXML generation
            suspicious_result = self._detect_suspicious_llm_content(parsed_output)
            errors.extend(suspicious_result.errors)
            warnings.extend(suspicious_result.warnings)
            
            if errors:
                self._log_security_violation("llm_output_validation", {
                    "output_type": type(llm_output).__name__,
                    "violations": errors
                })
            
            return ValidationResult(
                is_valid=len(errors) == 0,
                errors=errors,
                warnings=warnings
            )
            
        except Exception as e:
            logger.error(f"LLM output validation failed: {str(e)}")
            return ValidationResult(
                is_valid=False,
                errors=[f"LLM validation error: {str(e)}"],
                warnings=warnings
            )
    
    def sanitize_user_input(self, user_input: Dict[str, Any]) -> Tuple[Dict[str, Any], List[str]]:
        """
        Sanitize and validate user input parameters.
        
        Args:
            user_input: Raw user input dictionary
            
        Returns:
            Tuple of (sanitized_input, warnings)
            
        Raises:
            SecurityViolationError: If input contains security violations
        """
        sanitized = {}
        warnings = []
        
        for key, value in user_input.items():
            try:
                sanitized_value, field_warnings = self._sanitize_field(key, value)
                sanitized[key] = sanitized_value
                warnings.extend(field_warnings)
            except SecurityViolationError as e:
                self._log_security_violation("user_input_sanitization", {
                    "field": key,
                    "value_type": type(value).__name__,
                    "violation": str(e)
                })
                raise
        
        return sanitized, warnings
    
    def validate_operation_execution(self, operation: AtomicOperationUnion) -> ValidationResult:
        """
        Validate operation before execution for runtime safety.
        
        Args:
            operation: Operation to validate before execution
            
        Returns:
            ValidationResult with execution safety validation
        """
        errors = []
        warnings = []
        
        try:
            # Check operation type is whitelisted
            if operation.operation_type not in self.WHITELISTED_OPERATIONS:
                errors.append(f"Operation type '{operation.operation_type}' not in whitelist")
            
            # Operation-specific runtime validations
            if isinstance(operation, DeleteSectionByHeading):
                result = self._validate_delete_section_execution(operation)
                errors.extend(result.errors)
                warnings.extend(result.warnings)
                
            elif isinstance(operation, SetStyleRule):
                result = self._validate_set_style_execution(operation)
                errors.extend(result.errors)
                warnings.extend(result.warnings)
                
            elif isinstance(operation, ReassignParagraphsToStyle):
                result = self._validate_reassign_paragraphs_execution(operation)
                errors.extend(result.errors)
                warnings.extend(result.warnings)
                
            elif isinstance(operation, ClearDirectFormatting):
                result = self._validate_clear_formatting_execution(operation)
                errors.extend(result.errors)
                warnings.extend(result.warnings)
            
            if errors:
                self._log_security_violation("operation_execution_validation", {
                    "operation_type": operation.operation_type,
                    "violations": errors
                })
            
            return ValidationResult(
                is_valid=len(errors) == 0,
                errors=errors,
                warnings=warnings
            )
            
        except Exception as e:
            logger.error(f"Operation execution validation failed: {str(e)}")
            return ValidationResult(
                is_valid=False,
                errors=[f"Execution validation error: {str(e)}"],
                warnings=warnings
            )
    
    def get_violation_log(self) -> List[Dict[str, Any]]:
        """
        Get security violation log.
        
        Returns:
            List of security violation records
        """
        return self._violation_log.copy()
    
    def clear_violation_log(self):
        """Clear security violation log."""
        self._violation_log.clear()
    
    # Private validation methods
    
    def _validate_whitelist_operations(self, plan: PlanV1) -> ValidationResult:
        """Validate all operations are in whitelist."""
        errors = []
        warnings = []
        
        for i, operation in enumerate(plan.ops):
            op_type = operation.operation_type
            if op_type not in self.WHITELISTED_OPERATIONS:
                errors.append(
                    f"Operation {i+1}: '{op_type}' is not in whitelist. "
                    f"Allowed operations: {sorted(self.WHITELISTED_OPERATIONS)}"
                )
        
        return ValidationResult(is_valid=len(errors) == 0, errors=errors, warnings=warnings)
    
    def _validate_no_string_replacement(self, plan: PlanV1) -> ValidationResult:
        """Validate no string replacement patterns in operations."""
        errors = []
        warnings = []
        
        # Convert plan to string for pattern matching
        plan_str = json.dumps(plan.model_dump(), indent=2)
        
        for pattern in self.FORBIDDEN_STRING_PATTERNS:
            matches = re.findall(pattern, plan_str, re.IGNORECASE)
            if matches:
                errors.append(
                    f"Forbidden string replacement pattern detected: {pattern}. "
                    f"Matches: {matches[:3]}{'...' if len(matches) > 3 else ''}"
                )
        
        return ValidationResult(is_valid=len(errors) == 0, errors=errors, warnings=warnings)
    
    def _validate_word_object_layer_only(self, plan: PlanV1) -> ValidationResult:
        """Validate operations use Word object layer only."""
        errors = []
        warnings = []
        
        plan_str = json.dumps(plan.model_dump(), indent=2)
        
        for forbidden_op in self.FORBIDDEN_WORD_OPERATIONS:
            # Check for direct property access patterns
            patterns = [
                rf'\.{forbidden_op}\s*=',  # Property assignment
                rf'\.{forbidden_op}\s*\(',  # Method call
                rf'"{forbidden_op}"\s*:',   # JSON property
                rf"'{forbidden_op}'\s*:",   # JSON property with single quotes
            ]
            
            for pattern in patterns:
                if re.search(pattern, plan_str, re.IGNORECASE):
                    errors.append(
                        f"Forbidden Word COM operation detected: {forbidden_op}. "
                        f"Use Word object layer modifications only."
                    )
                    break
        
        return ValidationResult(is_valid=len(errors) == 0, errors=errors, warnings=warnings)
    
    def _validate_authorization_requirements(self, plan: PlanV1) -> ValidationResult:
        """Validate operations requiring authorization have it."""
        errors = []
        warnings = []
        
        for i, operation in enumerate(plan.ops):
            if operation.operation_type in self.AUTHORIZATION_REQUIRED_OPS:
                if isinstance(operation, ClearDirectFormatting):
                    if not getattr(operation, 'authorization_required', False):
                        errors.append(
                            f"Operation {i+1}: {operation.operation_type} requires "
                            f"explicit authorization (authorization_required=True)"
                        )
        
        return ValidationResult(is_valid=len(errors) == 0, errors=errors, warnings=warnings)
    
    def _validate_and_sanitize_parameters(self, plan: PlanV1) -> ValidationResult:
        """Validate and sanitize operation parameters."""
        errors = []
        warnings = []
        
        for i, operation in enumerate(plan.ops):
            try:
                # Validate parameter types and ranges (handled by Pydantic models)
                # Additional custom validations
                
                if isinstance(operation, DeleteSectionByHeading):
                    # Validate heading text is not empty and reasonable length
                    if not operation.heading_text.strip():
                        errors.append(f"Operation {i+1}: heading_text cannot be empty")
                    elif len(operation.heading_text) > 255:
                        errors.append(f"Operation {i+1}: heading_text too long (max 255 chars)")
                    
                    # Check for suspicious regex patterns
                    if operation.match.value == "REGEX":
                        try:
                            re.compile(operation.heading_text)
                        except re.error as e:
                            errors.append(f"Operation {i+1}: invalid regex pattern: {str(e)}")
                
                elif isinstance(operation, SetStyleRule):
                    # Validate style name
                    if not operation.target_style_name.strip():
                        errors.append(f"Operation {i+1}: target_style_name cannot be empty")
                    
                    # Validate font specifications
                    if operation.font:
                        if operation.font.color_hex:
                            if not re.match(r'^#[0-9A-Fa-f]{6}$', operation.font.color_hex):
                                errors.append(f"Operation {i+1}: invalid color_hex format")
                
                elif isinstance(operation, ReassignParagraphsToStyle):
                    # Validate selector is not empty
                    if not operation.selector:
                        errors.append(f"Operation {i+1}: selector cannot be empty")
                    
                    # Validate target style name
                    if not operation.target_style_name.strip():
                        errors.append(f"Operation {i+1}: target_style_name cannot be empty")
                
            except Exception as e:
                errors.append(f"Operation {i+1}: parameter validation error: {str(e)}")
        
        return ValidationResult(is_valid=len(errors) == 0, errors=errors, warnings=warnings)
    
    def _detect_suspicious_llm_content(self, parsed_output: Dict[str, Any]) -> ValidationResult:
        """Detect suspicious content that might indicate DOCX/OOXML generation."""
        errors = []
        warnings = []
        
        # Convert to string for pattern matching
        content_str = json.dumps(parsed_output, indent=2).lower()
        
        # Suspicious patterns that indicate DOCX/OOXML generation attempts
        suspicious_patterns = [
            r'<w:',           # Word XML namespace
            r'<xml',          # XML content
            r'docx',          # DOCX references
            r'ooxml',         # OOXML references
            r'content-type.*word',  # Word content types
            r'application/vnd\.openxmlformats',  # Office Open XML MIME types
            r'word/document\.xml',  # Word document XML
            r'word/styles\.xml',    # Word styles XML
            r'\.rels',        # Relationship files
            r'[Content_Types]\.xml',  # Content types XML
        ]
        
        for pattern in suspicious_patterns:
            if re.search(pattern, content_str):
                errors.append(
                    f"Suspicious content detected that may indicate DOCX/OOXML generation: {pattern}"
                )
        
        # Check for base64 encoded content (might be embedded files)
        base64_pattern = r'[A-Za-z0-9+/]{50,}={0,2}'
        base64_matches = re.findall(base64_pattern, content_str)
        if base64_matches:
            warnings.append(f"Detected {len(base64_matches)} potential base64 encoded content blocks")
        
        return ValidationResult(is_valid=len(errors) == 0, errors=errors, warnings=warnings)
    
    def _sanitize_field(self, field_name: str, value: Any) -> Tuple[Any, List[str]]:
        """Sanitize individual field value."""
        warnings = []
        
        if isinstance(value, str):
            # Remove potentially dangerous characters
            original_value = value
            
            # Remove null bytes
            value = value.replace('\x00', '')
            
            # Remove control characters except common whitespace
            value = re.sub(r'[\x01-\x08\x0B\x0C\x0E-\x1F\x7F]', '', value)
            
            # Limit length for text fields
            if len(value) > 1000:
                value = value[:1000]
                warnings.append(f"Field '{field_name}' truncated to 1000 characters")
            
            # Check for script injection attempts
            script_patterns = [
                r'<script',
                r'javascript:',
                r'vbscript:',
                r'on\w+\s*=',  # Event handlers
                r'eval\s*\(',
                r'exec\s*\(',
            ]
            
            for pattern in script_patterns:
                if re.search(pattern, value, re.IGNORECASE):
                    raise SecurityViolationError(
                        f"Potential script injection in field '{field_name}'",
                        field_name=field_name,
                        security_context={"script_injection": True, "pattern": pattern}
                    )
            
            if value != original_value:
                warnings.append(f"Field '{field_name}' was sanitized")
        
        elif isinstance(value, dict):
            # Recursively sanitize dictionary values
            sanitized_dict = {}
            for k, v in value.items():
                sanitized_key, key_warnings = self._sanitize_field(f"{field_name}.{k}", k)
                sanitized_value, value_warnings = self._sanitize_field(f"{field_name}.{k}", v)
                sanitized_dict[sanitized_key] = sanitized_value
                warnings.extend(key_warnings)
                warnings.extend(value_warnings)
            value = sanitized_dict
        
        elif isinstance(value, list):
            # Sanitize list items
            sanitized_list = []
            for i, item in enumerate(value):
                sanitized_item, item_warnings = self._sanitize_field(f"{field_name}[{i}]", item)
                sanitized_list.append(sanitized_item)
                warnings.extend(item_warnings)
            value = sanitized_list
        
        return value, warnings
    
    def _validate_delete_section_execution(self, operation: DeleteSectionByHeading) -> ValidationResult:
        """Validate delete section operation for execution safety."""
        errors = []
        warnings = []
        
        # Check for potentially dangerous regex patterns
        if operation.match.value == "REGEX":
            dangerous_patterns = [
                r'\.\*',      # Greedy wildcards
                r'\.\+',      # Greedy plus
                r'\{.*,\}',   # Unbounded quantifiers
                r'\(\?\!',    # Negative lookahead (complex)
                r'\(\?\<',    # Lookbehind (complex)
            ]
            
            for pattern in dangerous_patterns:
                if re.search(pattern, operation.heading_text):
                    warnings.append(f"Complex regex pattern detected: {pattern}")
        
        # Validate heading text doesn't contain suspicious content
        if any(char in operation.heading_text for char in ['<', '>', '&', '"', "'"]):
            warnings.append("Heading text contains HTML/XML-like characters")
        
        return ValidationResult(is_valid=len(errors) == 0, errors=errors, warnings=warnings)
    
    def _validate_set_style_execution(self, operation: SetStyleRule) -> ValidationResult:
        """Validate set style operation for execution safety."""
        errors = []
        warnings = []
        
        # Validate font names don't contain suspicious content
        if operation.font:
            for font_field in ['east_asian', 'latin']:
                font_name = getattr(operation.font, font_field, None)
                if font_name and any(char in font_name for char in ['<', '>', '&', '"', "'"]):
                    warnings.append(f"Font name '{font_name}' contains suspicious characters")
        
        return ValidationResult(is_valid=len(errors) == 0, errors=errors, warnings=warnings)
    
    def _validate_reassign_paragraphs_execution(self, operation: ReassignParagraphsToStyle) -> ValidationResult:
        """Validate reassign paragraphs operation for execution safety."""
        errors = []
        warnings = []
        
        # Validate selector doesn't contain dangerous patterns
        selector_str = json.dumps(operation.selector)
        if any(pattern in selector_str.lower() for pattern in ['script', 'eval', 'exec']):
            errors.append("Selector contains potentially dangerous patterns")
        
        return ValidationResult(is_valid=len(errors) == 0, errors=errors, warnings=warnings)
    
    def _validate_clear_formatting_execution(self, operation: ClearDirectFormatting) -> ValidationResult:
        """Validate clear formatting operation for execution safety."""
        errors = []
        warnings = []
        
        # This is a high-risk operation, ensure authorization
        if not operation.authorization_required:
            errors.append("clear_direct_formatting requires explicit authorization")
        
        # Validate range specification if provided
        if operation.range_spec:
            start = operation.range_spec.get('start', 0)
            end = operation.range_spec.get('end', 0)
            
            if start < 0 or end < 0:
                errors.append("Range specification cannot have negative values")
            
            if end > 0 and start >= end:
                errors.append("Range start must be less than end")
        
        return ValidationResult(is_valid=len(errors) == 0, errors=errors, warnings=warnings)
    
    def _log_security_violation(self, violation_type: str, context: Dict[str, Any]):
        """Log security violation for audit trail."""
        violation_record = {
            "timestamp": logger.handlers[0].formatter.formatTime(logger.makeRecord(
                logger.name, logging.WARNING, __file__, 0, "", (), None
            )) if logger.handlers else "unknown",
            "violation_type": violation_type,
            "context": context
        }
        
        self._violation_log.append(violation_record)
        logger.warning(f"Security violation: {violation_type} - {context}")


# Convenience functions for direct constraint validation

def validate_plan_constraints(plan: PlanV1) -> ValidationResult:
    """
    Validate plan against runtime constraints.
    
    Args:
        plan: Plan to validate
        
    Returns:
        ValidationResult with constraint validation results
    """
    enforcer = RuntimeConstraintEnforcer()
    return enforcer.validate_plan_constraints(plan)


def validate_llm_output_constraints(llm_output: Union[str, Dict[str, Any]]) -> ValidationResult:
    """
    Validate LLM output against constraints.
    
    Args:
        llm_output: LLM output to validate
        
    Returns:
        ValidationResult with LLM output validation results
    """
    enforcer = RuntimeConstraintEnforcer()
    return enforcer.validate_llm_output(llm_output)


def sanitize_user_input_safe(user_input: Dict[str, Any]) -> Tuple[Dict[str, Any], List[str]]:
    """
    Safely sanitize user input.
    
    Args:
        user_input: User input to sanitize
        
    Returns:
        Tuple of (sanitized_input, warnings)
        
    Raises:
        SecurityViolationError: If input contains security violations
    """
    enforcer = RuntimeConstraintEnforcer()
    return enforcer.sanitize_user_input(user_input)