"""
AutoWord vNext - Next generation document processing pipeline.

This module implements a structured "Extract→Plan→Execute→Validate→Audit" 
closed loop system for Word document processing with >99% stability.
"""

__version__ = "2.0.0"
__author__ = "AutoWord Team"

from .models import StructureV1, PlanV1, InventoryFullV1, ProcessingResult
from .exceptions import VNextError, ExtractionError, PlanningError, ExecutionError, ValidationError, AuditError
from .core import VNextConfig, LLMConfig, LocalizationConfig, ValidationConfig, AuditConfig, ExecutorConfig, CustomLLMClient, load_config, save_config
from .simple_pipeline import SimplePipeline, VNextPipeline

# 保持向后兼容性
try:
    from .schema_validator import (
        SchemaValidator, validate_structure, validate_plan, validate_inventory,
        validate_with_detailed_errors
    )
    from .error_handler import (
        PipelineErrorHandler, WarningsLogger, SecurityValidator, RollbackManager,
        RevisionHandler, ProcessingStatus, RevisionHandlingStrategy, ErrorContext,
        RecoveryResult
    )
    from .pipeline import ProgressReporter
except ImportError:
    # 如果完整版本不可用，使用简化版本
    pass

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
    "VNextConfig",
    "LLMConfig",
    "LocalizationConfig",
    "ValidationConfig", 
    "AuditConfig",
    "ExecutorConfig",
    "CustomLLMClient",
    "load_config",
    "save_config",
    "SimplePipeline",
    "VNextPipeline"
]