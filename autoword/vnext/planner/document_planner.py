"""
Document planner for generating execution plans through LLM with strict JSON schema validation.

This module generates plan.v1.json files through LLM integration with whitelist operation enforcement.
"""

import json
import logging
from typing import Dict, Any, List, Optional, Set
from pathlib import Path

import jsonschema
from pydantic import ValidationError as PydanticValidationError

from ..models import (
    StructureV1, PlanV1, ValidationResult, AtomicOperationUnion,
    DeleteSectionByHeading, UpdateToc, DeleteToc, SetStyleRule,
    ReassignParagraphsToStyle, ClearDirectFormatting, ProcessingResult
)
from ..exceptions import PlanningError, SchemaValidationError, WhitelistViolationError
from ..constraints import RuntimeConstraintEnforcer
from ...core.llm_client import LLMClient, ModelType, LLMResponse


logger = logging.getLogger(__name__)


class DocumentPlanner:
    """Generate execution plans through LLM with schema validation."""
    
    # Whitelisted operation types
    WHITELISTED_OPERATIONS: Set[str] = {
        "delete_section_by_heading",
        "update_toc",
        "delete_toc", 
        "set_style_rule",
        "reassign_paragraphs_to_style",
        "clear_direct_formatting"
    }
    
    def __init__(self, llm_client: Optional[LLMClient] = None, 
                 schema_path: Optional[str] = None,
                 constraint_enforcer: Optional[RuntimeConstraintEnforcer] = None):
        """
        Initialize document planner.
        
        Args:
            llm_client: LLM client instance (creates default if None)
            schema_path: Path to plan.v1.json schema file
            constraint_enforcer: Runtime constraint enforcer (creates default if None)
        """
        self.llm_client = llm_client or LLMClient()
        self.constraint_enforcer = constraint_enforcer or RuntimeConstraintEnforcer()
        
        # Load JSON schema for validation
        if schema_path is None:
            schema_path = Path(__file__).parent.parent / "schemas" / "plan.v1.json"
        
        try:
            with open(schema_path, 'r', encoding='utf-8') as f:
                self.plan_schema = json.load(f)
        except Exception as e:
            raise PlanningError(f"Failed to load plan schema: {str(e)}", schema_path=str(schema_path))
    
    def generate_plan(self, structure: StructureV1, user_intent: str) -> PlanV1:
        """
        Generate plan through LLM with schema validation.
        
        Args:
            structure: Document structure
            user_intent: User's intent for document modification
            
        Returns:
            PlanV1: Validated execution plan
            
        Raises:
            PlanningError: If plan generation or validation fails
        """
        try:
            # Build system prompt for LLM
            system_prompt = self._build_system_prompt()
            
            # Build user prompt with structure and intent
            user_prompt = self._build_user_prompt(structure, user_intent)
            
            # Call LLM with JSON retry logic
            logger.info("Generating plan through LLM...")
            response = self.llm_client.call_with_json_retry(
                ModelType.GPT5,
                system_prompt,
                user_prompt,
                max_json_retries=3
            )
            
            if not response.success:
                raise PlanningError(
                    f"LLM call failed: {response.error}",
                    llm_response=response.content
                )
            
            # Validate LLM output with constraint enforcement
            logger.info("Validating LLM output with constraint enforcement...")
            llm_validation = self.constraint_enforcer.validate_llm_output(response.content)
            if not llm_validation.is_valid:
                raise PlanningError(
                    f"LLM output failed constraint validation: {'; '.join(llm_validation.errors)}",
                    llm_response=response.content,
                    validation_errors=llm_validation.errors
                )
            
            # Parse JSON response (already validated by constraint enforcer)
            try:
                plan_data = json.loads(response.content)
            except json.JSONDecodeError as e:
                raise PlanningError(
                    f"Invalid JSON response from LLM: {str(e)}",
                    llm_response=response.content
                )
            
            # Additional schema validation (redundant but thorough)
            schema_validation = self.validate_plan_schema(plan_data)
            if not schema_validation.is_valid:
                raise SchemaValidationError(
                    "Plan failed schema validation",
                    schema_name="plan.v1",
                    validation_errors=schema_validation.errors,
                    invalid_data=plan_data
                )
            
            # Convert to Pydantic model
            try:
                plan = PlanV1(**plan_data)
            except PydanticValidationError as e:
                raise PlanningError(
                    f"Plan data validation failed: {str(e)}",
                    plan_data=plan_data,
                    validation_errors=[str(err) for err in e.errors()]
                )
            
            # Comprehensive constraint validation
            logger.info("Validating plan constraints...")
            constraint_validation = self.constraint_enforcer.validate_plan_constraints(plan)
            if not constraint_validation.is_valid:
                raise WhitelistViolationError(
                    f"Plan failed constraint validation: {'; '.join(constraint_validation.errors)}",
                    invalid_operations=[],
                    plan_data=plan_data,
                    validation_errors=constraint_validation.errors
                )
            
            logger.info(f"Successfully generated plan with {len(plan.ops)} operations")
            return plan
            
        except (PlanningError, SchemaValidationError, WhitelistViolationError):
            raise
        except Exception as e:
            raise PlanningError(f"Unexpected error during plan generation: {str(e)}")
    
    def validate_plan_schema(self, plan_json: dict) -> ValidationResult:
        """
        Validate against plan.v1 schema.
        
        Args:
            plan_json: Plan data to validate
            
        Returns:
            ValidationResult: Validation result with errors/warnings
        """
        errors = []
        warnings = []
        
        try:
            # Validate against JSON schema
            jsonschema.validate(plan_json, self.plan_schema)
            
            # Additional custom validations
            if "ops" in plan_json:
                for i, op in enumerate(plan_json["ops"]):
                    if not isinstance(op, dict):
                        errors.append(f"Operation {i}: Must be an object")
                        continue
                    
                    if "operation_type" not in op:
                        errors.append(f"Operation {i}: Missing 'operation_type' field")
                        continue
                    
                    # Validate operation-specific constraints
                    op_type = op.get("operation_type")
                    validation_errors = self._validate_operation_constraints(op, op_type)
                    errors.extend([f"Operation {i}: {err}" for err in validation_errors])
            
        except jsonschema.ValidationError as e:
            errors.append(f"Schema validation error: {e.message}")
        except Exception as e:
            errors.append(f"Validation error: {str(e)}")
        
        return ValidationResult(
            is_valid=len(errors) == 0,
            errors=errors,
            warnings=warnings
        )
    
    def check_whitelist_compliance(self, plan: PlanV1) -> ValidationResult:
        """
        Ensure only whitelisted operations.
        
        Args:
            plan: Plan to check for whitelist compliance
            
        Returns:
            ValidationResult: Compliance check result
        """
        errors = []
        warnings = []
        
        for i, op in enumerate(plan.ops):
            op_type = op.operation_type
            
            if op_type not in self.WHITELISTED_OPERATIONS:
                errors.append(f"{op_type}: Operation type not whitelisted")
            
            # Additional security checks for specific operations
            if op_type == "clear_direct_formatting":
                if not getattr(op, 'authorization_required', False):
                    errors.append(f"Operation {i}: clear_direct_formatting requires authorization")
        
        return ValidationResult(
            is_valid=len(errors) == 0,
            errors=errors,
            warnings=warnings
        )
    
    def _build_system_prompt(self) -> str:
        """Build system prompt for LLM plan generation."""
        return """You are an expert document automation system that generates execution plans for Word document modifications.

Your task is to analyze a document structure and user intent, then generate a JSON plan with atomic operations.

CRITICAL REQUIREMENTS:
1. Return ONLY valid JSON matching the plan.v1 schema
2. Use ONLY these whitelisted operations:
   - delete_section_by_heading
   - update_toc
   - delete_toc
   - set_style_rule
   - reassign_paragraphs_to_style
   - clear_direct_formatting

3. Each operation must be atomic and safe
4. Never modify content directly - only through Word object layer
5. For Chinese documents, use proper style names (标题 1, 正文, etc.)

OPERATION GUIDELINES:
- delete_section_by_heading: Remove content from heading to next same-level heading
- update_toc: Refresh all TOC fields and page numbers
- delete_toc: Remove table of contents (mode: all/first/last)
- set_style_rule: Modify style definitions (fonts, spacing, etc.)
- reassign_paragraphs_to_style: Change paragraph style assignments
- clear_direct_formatting: Remove direct formatting (requires authorization)

RESPONSE FORMAT:
{
  "schema_version": "plan.v1",
  "ops": [
    {
      "operation_type": "delete_section_by_heading",
      "heading_text": "摘要",
      "level": 1,
      "match": "EXACT",
      "case_sensitive": false
    }
  ]
}

Return only the JSON plan, no explanations or markdown formatting."""

    def _build_user_prompt(self, structure: StructureV1, user_intent: str) -> str:
        """
        Build user prompt with structure and intent.
        
        Args:
            structure: Document structure
            user_intent: User's modification intent
            
        Returns:
            User prompt string
        """
        # Convert structure to JSON for LLM
        structure_json = structure.model_dump_json(indent=2)
        
        return f"""Document Structure:
{structure_json}

User Intent:
{user_intent}

Generate an execution plan to fulfill the user's intent. Analyze the document structure and create atomic operations that will safely modify the document according to the requirements.

Focus on:
1. Identifying sections to delete based on headings
2. TOC management (update/delete as needed)
3. Style modifications for proper formatting
4. Paragraph reassignment for consistent styling

Return the plan as valid JSON following the plan.v1 schema."""

    def _validate_operation_constraints(self, op: Dict[str, Any], op_type: str) -> List[str]:
        """
        Validate operation-specific constraints.
        
        Args:
            op: Operation data
            op_type: Operation type
            
        Returns:
            List of validation errors
        """
        errors = []
        
        if op_type == "delete_section_by_heading":
            if not op.get("heading_text"):
                errors.append("heading_text is required and cannot be empty")
            
            level = op.get("level")
            if not isinstance(level, int) or level < 1 or level > 9:
                errors.append("level must be an integer between 1 and 9")
            
            match_mode = op.get("match", "EXACT")
            if match_mode not in ["EXACT", "CONTAINS", "REGEX"]:
                errors.append("match must be one of: EXACT, CONTAINS, REGEX")
        
        elif op_type == "set_style_rule":
            if not op.get("target_style_name"):
                errors.append("target_style_name is required and cannot be empty")
            
            # Validate font specification if present
            font = op.get("font")
            if font:
                size_pt = font.get("size_pt")
                if size_pt is not None and (not isinstance(size_pt, int) or size_pt < 1 or size_pt > 72):
                    errors.append("font.size_pt must be an integer between 1 and 72")
        
        elif op_type == "reassign_paragraphs_to_style":
            if not op.get("selector"):
                errors.append("selector is required and cannot be empty")
            
            if not op.get("target_style_name"):
                errors.append("target_style_name is required and cannot be empty")
        
        elif op_type == "clear_direct_formatting":
            scope = op.get("scope")
            if scope not in ["document", "selection", "range"]:
                errors.append("scope must be one of: document, selection, range")
            
            if not op.get("authorization_required"):
                errors.append("authorization_required must be true for clear_direct_formatting")
        
        return errors