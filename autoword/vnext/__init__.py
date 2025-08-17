"""
AutoWord vNext - Next generation document processing pipeline.

This module implements a structured "Extract→Plan→Execute→Validate→Audit" 
closed loop system for Word document processing with >99% stability.
"""

__version__ = "2.0.0"
__author__ = "AutoWord Team"

from .models import StructureV1, PlanV1, InventoryFullV1, ProcessingResult
from .exceptions import VNextError, ExtractionError, PlanningError, ExecutionError, ValidationError, AuditError
from .schema_validator import (
    SchemaValidator, validate_structure, validate_plan, validate_inventory,
    validate_with_detailed_errors
)
from .error_handler import (
    PipelineErrorHandler, WarningsLogger, SecurityValidator, RollbackManager,
    RevisionHandler, ProcessingStatus, RevisionHandlingStrategy, ErrorContext,
    RecoveryResult
)
from .pipeline import VNextPipeline, ProgressReporter

__all__ = [
    "StructureV1",
    "PlanV1", 
    "InventoryFullV1",
    "ProcessingResult",
    "VNextError",
    "ExtractionError",
    "PlanningError", 
    "ExecutionError",
    "ValidationError",
    "AuditError",
    "SchemaValidator",
    "validate_structure",
    "validate_plan", 
    "validate_inventory",
    "validate_with_detailed_errors",
    "PipelineErrorHandler",
    "WarningsLogger",
    "SecurityValidator", 
    "RollbackManager",
    "RevisionHandler",
    "ProcessingStatus",
    "RevisionHandlingStrategy",
    "ErrorContext",
    "RecoveryResult",
    "VNextPipeline",
    "ProgressReporter"
]