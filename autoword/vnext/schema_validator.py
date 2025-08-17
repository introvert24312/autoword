"""
Comprehensive JSON schema validation system for AutoWord vNext.

This module provides centralized schema validation for all vNext data models:
- structure.v1.json validation with all required fields
- plan.v1.json validation with whitelist operation enforcement  
- inventory.full.v1.json validation with OOXML fragment storage
- Detailed error reporting and validation utilities
"""

import json
import jsonschema
from pathlib import Path
from typing import Dict, Any, List, Optional, Union, Type
from pydantic import ValidationError

from .models import (
    StructureV1, PlanV1, InventoryFullV1, ValidationResult,
    AtomicOperationUnion, SchemaVersion
)
from .exceptions import SchemaValidationError


class SchemaValidator:
    """Comprehensive schema validator for all vNext data models."""
    
    def __init__(self, schemas_dir: Optional[Path] = None):
        """
        Initialize schema validator with JSON schema files.
        
        Args:
            schemas_dir: Directory containing JSON schema files
        """
        if schemas_dir is None:
            schemas_dir = Path(__file__).parent / "schemas"
        
        self.schemas_dir = schemas_dir
        self._schemas: Dict[str, Dict[str, Any]] = {}
        self._load_schemas()
    
    def _load_schemas(self) -> None:
        """Load all JSON schema files."""
        schema_files = {
            SchemaVersion.STRUCTURE_V1: "structure.v1.json",
            SchemaVersion.PLAN_V1: "plan.v1.json", 
            SchemaVersion.INVENTORY_FULL_V1: "inventory.full.v1.json"
        }
        
        for schema_version, filename in schema_files.items():
            schema_path = self.schemas_dir / filename
            if not schema_path.exists():
                raise FileNotFoundError(f"Schema file not found: {schema_path}")
            
            try:
                with open(schema_path, 'r', encoding='utf-8') as f:
                    self._schemas[schema_version] = json.load(f)
            except json.JSONDecodeError as e:
                raise SchemaValidationError(
                    f"Invalid JSON in schema file: {schema_path}",
                    schema_name=schema_version,
                    validation_errors=[str(e)],
                    invalid_data={}
                )
    
    def validate_structure_v1(self, data: Union[Dict[str, Any], StructureV1]) -> ValidationResult:
        """
        Validate structure.v1.json with all required fields.
        
        Args:
            data: Structure data to validate (dict or StructureV1 model)
            
        Returns:
            ValidationResult with detailed error information
        """
        if isinstance(data, StructureV1):
            data_dict = data.model_dump()
        else:
            data_dict = data
        
        return self._validate_against_schema(
            data_dict, 
            SchemaVersion.STRUCTURE_V1,
            additional_validations=self._validate_structure_constraints
        )
    
    def validate_plan_v1(self, data: Union[Dict[str, Any], PlanV1]) -> ValidationResult:
        """
        Validate plan.v1.json with whitelist operation enforcement.
        
        Args:
            data: Plan data to validate (dict or PlanV1 model)
            
        Returns:
            ValidationResult with detailed error information
        """
        if isinstance(data, PlanV1):
            data_dict = data.model_dump()
        else:
            data_dict = data
        
        return self._validate_against_schema(
            data_dict,
            SchemaVersion.PLAN_V1,
            additional_validations=self._validate_plan_constraints
        )
    
    def validate_inventory_full_v1(self, data: Union[Dict[str, Any], InventoryFullV1]) -> ValidationResult:
        """
        Validate inventory.full.v1.json with OOXML fragment storage.
        
        Args:
            data: Inventory data to validate (dict or InventoryFullV1 model)
            
        Returns:
            ValidationResult with detailed error information
        """
        if isinstance(data, InventoryFullV1):
            data_dict = data.model_dump()
        else:
            data_dict = data
        
        return self._validate_against_schema(
            data_dict,
            SchemaVersion.INVENTORY_FULL_V1,
            additional_validations=self._validate_inventory_constraints
        )
    
    def validate_pydantic_model(self, data: Dict[str, Any], model_class: Type) -> ValidationResult:
        """
        Validate data against Pydantic model with detailed error reporting.
        
        Args:
            data: Data to validate
            model_class: Pydantic model class to validate against
            
        Returns:
            ValidationResult with detailed error information
        """
        errors = []
        warnings = []
        
        try:
            # Attempt to create model instance
            model_instance = model_class(**data)
            return ValidationResult(is_valid=True, errors=[], warnings=warnings)
            
        except ValidationError as e:
            for error in e.errors():
                field_path = " -> ".join(str(loc) for loc in error['loc'])
                error_msg = f"Field '{field_path}': {error['msg']}"
                if error['type'] == 'value_error':
                    error_msg += f" (input: {error.get('input', 'N/A')})"
                errors.append(error_msg)
            
            return ValidationResult(is_valid=False, errors=errors, warnings=warnings)
        
        except Exception as e:
            errors.append(f"Unexpected validation error: {str(e)}")
            return ValidationResult(is_valid=False, errors=errors, warnings=warnings)
    
    def _validate_against_schema(
        self, 
        data: Dict[str, Any], 
        schema_version: str,
        additional_validations: Optional[callable] = None
    ) -> ValidationResult:
        """
        Validate data against JSON schema with optional additional validations.
        
        Args:
            data: Data to validate
            schema_version: Schema version to validate against
            additional_validations: Optional additional validation function
            
        Returns:
            ValidationResult with detailed error information
        """
        errors = []
        warnings = []
        
        # JSON Schema validation
        try:
            schema = self._schemas[schema_version]
            jsonschema.validate(data, schema)
            
        except jsonschema.ValidationError as e:
            errors.append(f"Schema validation error: {e.message}")
            if e.absolute_path:
                path = " -> ".join(str(p) for p in e.absolute_path)
                errors.append(f"Error location: {path}")
            if e.validator_value:
                errors.append(f"Expected: {e.validator_value}")
            if hasattr(e, 'instance') and e.instance is not None:
                errors.append(f"Actual value: {e.instance}")
                
        except jsonschema.SchemaError as e:
            errors.append(f"Schema definition error: {e.message}")
            
        except Exception as e:
            errors.append(f"Validation error: {str(e)}")
        
        # Additional custom validations
        if additional_validations and not errors:
            try:
                additional_errors, additional_warnings = additional_validations(data)
                errors.extend(additional_errors)
                warnings.extend(additional_warnings)
            except Exception as e:
                errors.append(f"Additional validation error: {str(e)}")
        
        return ValidationResult(
            is_valid=len(errors) == 0,
            errors=errors,
            warnings=warnings
        )
    
    def _validate_structure_constraints(self, data: Dict[str, Any]) -> tuple[List[str], List[str]]:
        """
        Additional validation constraints for structure.v1.json.
        
        Args:
            data: Structure data to validate
            
        Returns:
            Tuple of (errors, warnings)
        """
        errors = []
        warnings = []
        
        # Validate paragraph indexes are sequential and unique
        if 'paragraphs' in data:
            paragraph_indexes = [p.get('index', -1) for p in data['paragraphs']]
            if len(paragraph_indexes) != len(set(paragraph_indexes)):
                errors.append("Paragraph indexes must be unique")
            
            if paragraph_indexes and min(paragraph_indexes) < 0:
                errors.append("Paragraph indexes must be non-negative")
            
            expected_indexes = list(range(len(paragraph_indexes)))
            if sorted(paragraph_indexes) != expected_indexes:
                warnings.append("Paragraph indexes are not sequential starting from 0")
        
        # Validate heading references point to valid paragraphs
        if 'headings' in data and 'paragraphs' in data:
            paragraph_indexes = set(p.get('index', -1) for p in data['paragraphs'])
            for i, heading in enumerate(data['headings']):
                heading_para_idx = heading.get('paragraph_index', -1)
                if heading_para_idx not in paragraph_indexes:
                    errors.append(f"Heading {i}: paragraph_index {heading_para_idx} not found in paragraphs")
        
        # Validate field references point to valid paragraphs
        if 'fields' in data and 'paragraphs' in data:
            paragraph_indexes = set(p.get('index', -1) for p in data['paragraphs'])
            for i, field in enumerate(data['fields']):
                field_para_idx = field.get('paragraph_index', -1)
                if field_para_idx not in paragraph_indexes:
                    errors.append(f"Field {i}: paragraph_index {field_para_idx} not found in paragraphs")
        
        # Validate table references point to valid paragraphs
        if 'tables' in data and 'paragraphs' in data:
            paragraph_indexes = set(p.get('index', -1) for p in data['paragraphs'])
            for i, table in enumerate(data['tables']):
                table_para_idx = table.get('paragraph_index', -1)
                if table_para_idx not in paragraph_indexes:
                    errors.append(f"Table {i}: paragraph_index {table_para_idx} not found in paragraphs")
                
                # Validate cell references if present
                cell_refs = table.get('cell_references', [])
                if cell_refs:
                    for j, cell_ref in enumerate(cell_refs):
                        if cell_ref not in paragraph_indexes:
                            errors.append(f"Table {i}, cell {j}: paragraph_index {cell_ref} not found in paragraphs")
        
        # Validate style consistency
        if 'styles' in data:
            style_names = set()
            for i, style in enumerate(data['styles']):
                style_name = style.get('name', '')
                if style_name in style_names:
                    errors.append(f"Duplicate style name: {style_name}")
                style_names.add(style_name)
                
                # Validate based_on references
                based_on = style.get('based_on')
                if based_on and based_on not in style_names:
                    warnings.append(f"Style '{style_name}' based_on '{based_on}' not found in styles")
        
        return errors, warnings
    
    def _validate_plan_constraints(self, data: Dict[str, Any]) -> tuple[List[str], List[str]]:
        """
        Additional validation constraints for plan.v1.json with whitelist enforcement.
        
        Args:
            data: Plan data to validate
            
        Returns:
            Tuple of (errors, warnings)
        """
        errors = []
        warnings = []
        
        # Whitelisted operation types
        WHITELISTED_OPERATIONS = {
            "delete_section_by_heading",
            "update_toc", 
            "delete_toc",
            "set_style_rule",
            "reassign_paragraphs_to_style",
            "clear_direct_formatting"
        }
        
        if 'ops' in data:
            for i, op in enumerate(data['ops']):
                op_type = op.get('operation_type', '')
                
                # Whitelist validation (redundant with JSON schema but provides clearer error)
                if op_type not in WHITELISTED_OPERATIONS:
                    errors.append(f"Operation {i}: '{op_type}' is not in whitelist of allowed operations: {sorted(WHITELISTED_OPERATIONS)}")
                
                # Operation-specific validations
                if op_type == "delete_section_by_heading":
                    if not op.get('heading_text', '').strip():
                        errors.append(f"Operation {i}: heading_text cannot be empty")
                    
                    level = op.get('level', 0)
                    if not (1 <= level <= 9):
                        errors.append(f"Operation {i}: level must be between 1 and 9")
                
                elif op_type == "set_style_rule":
                    if not op.get('target_style_name', '').strip():
                        errors.append(f"Operation {i}: target_style_name cannot be empty")
                    
                    # Validate font specifications
                    font = op.get('font', {})
                    if font:
                        size_pt = font.get('size_pt')
                        if size_pt is not None and not (1 <= size_pt <= 72):
                            errors.append(f"Operation {i}: font size_pt must be between 1 and 72")
                
                elif op_type == "reassign_paragraphs_to_style":
                    if not op.get('selector'):
                        errors.append(f"Operation {i}: selector cannot be empty")
                    if not op.get('target_style_name', '').strip():
                        errors.append(f"Operation {i}: target_style_name cannot be empty")
                
                elif op_type == "clear_direct_formatting":
                    if not op.get('authorization_required', False):
                        errors.append(f"Operation {i}: clear_direct_formatting requires explicit authorization")
        
        return errors, warnings
    
    def _validate_inventory_constraints(self, data: Dict[str, Any]) -> tuple[List[str], List[str]]:
        """
        Additional validation constraints for inventory.full.v1.json.
        
        Args:
            data: Inventory data to validate
            
        Returns:
            Tuple of (errors, warnings)
        """
        errors = []
        warnings = []
        
        # Validate OOXML fragments are valid strings
        if 'ooxml_fragments' in data:
            for fragment_id, fragment_content in data['ooxml_fragments'].items():
                if not isinstance(fragment_content, str):
                    errors.append(f"OOXML fragment '{fragment_id}': content must be string")
                elif not fragment_content.strip():
                    warnings.append(f"OOXML fragment '{fragment_id}': content is empty")
        
        # Validate media indexes have required fields
        if 'media_indexes' in data:
            for media_path, media_ref in data['media_indexes'].items():
                if not isinstance(media_ref, dict):
                    errors.append(f"Media index '{media_path}': must be object")
                    continue
                
                if not media_ref.get('media_id', '').strip():
                    errors.append(f"Media index '{media_path}': media_id cannot be empty")
                
                if not media_ref.get('content_type', '').strip():
                    errors.append(f"Media index '{media_path}': content_type cannot be empty")
                
                size_bytes = media_ref.get('size_bytes')
                if size_bytes is not None and size_bytes < 0:
                    errors.append(f"Media index '{media_path}': size_bytes cannot be negative")
        
        # Validate paragraph references in content controls, formulas, charts
        reference_collections = [
            ('content_controls', 'Content control'),
            ('formulas', 'Formula'),
            ('charts', 'Chart')
        ]
        
        for collection_name, item_type in reference_collections:
            if collection_name in data:
                for i, item in enumerate(data[collection_name]):
                    if not isinstance(item, dict):
                        errors.append(f"{item_type} {i}: must be object")
                        continue
                    
                    para_idx = item.get('paragraph_index')
                    if para_idx is None or para_idx < 0:
                        errors.append(f"{item_type} {i}: paragraph_index must be non-negative")
                    
                    # Validate required ID fields
                    id_field = f"{collection_name[:-1]}_id"  # Remove 's' and add '_id'
                    if collection_name == 'content_controls':
                        id_field = 'control_id'
                    
                    if not item.get(id_field, '').strip():
                        errors.append(f"{item_type} {i}: {id_field} cannot be empty")
        
        return errors, warnings
    
    def get_schema(self, schema_version: str) -> Dict[str, Any]:
        """
        Get loaded JSON schema by version.
        
        Args:
            schema_version: Schema version to retrieve
            
        Returns:
            JSON schema dictionary
            
        Raises:
            KeyError: If schema version not found
        """
        if schema_version not in self._schemas:
            raise KeyError(f"Schema version '{schema_version}' not found")
        
        return self._schemas[schema_version].copy()
    
    def list_available_schemas(self) -> List[str]:
        """
        List all available schema versions.
        
        Returns:
            List of available schema version strings
        """
        return list(self._schemas.keys())


# Convenience functions for direct validation

def validate_structure(data: Union[Dict[str, Any], StructureV1]) -> ValidationResult:
    """
    Validate structure.v1.json data.
    
    Args:
        data: Structure data to validate
        
    Returns:
        ValidationResult with detailed error information
    """
    validator = SchemaValidator()
    return validator.validate_structure_v1(data)


def validate_plan(data: Union[Dict[str, Any], PlanV1]) -> ValidationResult:
    """
    Validate plan.v1.json data.
    
    Args:
        data: Plan data to validate
        
    Returns:
        ValidationResult with detailed error information
    """
    validator = SchemaValidator()
    return validator.validate_plan_v1(data)


def validate_inventory(data: Union[Dict[str, Any], InventoryFullV1]) -> ValidationResult:
    """
    Validate inventory.full.v1.json data.
    
    Args:
        data: Inventory data to validate
        
    Returns:
        ValidationResult with detailed error information
    """
    validator = SchemaValidator()
    return validator.validate_inventory_full_v1(data)


def validate_with_detailed_errors(
    data: Dict[str, Any], 
    schema_version: str,
    raise_on_error: bool = False
) -> ValidationResult:
    """
    Validate data with detailed error reporting.
    
    Args:
        data: Data to validate
        schema_version: Schema version to validate against
        raise_on_error: Whether to raise SchemaValidationError on validation failure
        
    Returns:
        ValidationResult with detailed error information
        
    Raises:
        SchemaValidationError: If validation fails and raise_on_error is True
    """
    validator = SchemaValidator()
    
    if schema_version == SchemaVersion.STRUCTURE_V1:
        result = validator.validate_structure_v1(data)
    elif schema_version == SchemaVersion.PLAN_V1:
        result = validator.validate_plan_v1(data)
    elif schema_version == SchemaVersion.INVENTORY_FULL_V1:
        result = validator.validate_inventory_full_v1(data)
    else:
        raise ValueError(f"Unknown schema version: {schema_version}")
    
    if not result.is_valid and raise_on_error:
        raise SchemaValidationError(
            f"Validation failed for {schema_version}",
            schema_name=schema_version,
            validation_errors=result.errors,
            invalid_data=data
        )
    
    return result