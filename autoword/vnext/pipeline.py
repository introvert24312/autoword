"""
Main pipeline orchestrator for AutoWord vNext.

This module implements the VNextPipeline class that orchestrates all five modules
in sequence: Extract→Plan→Execute→Validate→Audit with complete error handling
and rollback capabilities.
"""

import os
import logging
import tempfile
import shutil
from datetime import datetime
from pathlib import Path
from typing import Optional, Dict, Any, Callable

from .models import StructureV1, InventoryFullV1, PlanV1, ProcessingResult
from .exceptions import VNextError, ExtractionError, PlanningError, ExecutionError, ValidationError, AuditError
from .extractor.document_extractor import DocumentExtractor
from .planner.document_planner import DocumentPlanner
from .executor.document_executor import DocumentExecutor
from .validator.document_validator import DocumentValidator
from .auditor.document_auditor import DocumentAuditor
from .error_handler import PipelineErrorHandler, ErrorContext, ProcessingStatus
from .monitoring import VNextLogger, MonitoringLevel, create_vnext_logger, log_large_document_warning, log_complex_document_scenario
from ..core.llm_client import LLMClient


logger = logging.getLogger(__name__)


class ProgressReporter:
    """Progress reporting for pipeline execution."""
    
    def __init__(self, progress_callback: Optional[Callable[[str, int], None]] = None):
        """
        Initialize progress reporter.
        
        Args:
            progress_callback: Optional callback function(stage_name, progress_percent)
        """
        self.progress_callback = progress_callback
        self.current_stage = ""
        self.total_stages = 5
        self.completed_stages = 0
    
    def start_stage(self, stage_name: str):
        """Start a new pipeline stage."""
        self.current_stage = stage_name
        progress_percent = int((self.completed_stages / self.total_stages) * 100)
        logger.info(f"Starting stage: {stage_name} ({progress_percent}%)")
        
        if self.progress_callback:
            self.progress_callback(stage_name, progress_percent)
    
    def complete_stage(self):
        """Complete the current stage."""
        self.completed_stages += 1
        progress_percent = int((self.completed_stages / self.total_stages) * 100)
        logger.info(f"Completed stage: {self.current_stage} ({progress_percent}%)")
        
        if self.progress_callback:
            self.progress_callback(self.current_stage, progress_percent)
    
    def report_substep(self, substep: str):
        """Report a substep within the current stage."""
        logger.debug(f"{self.current_stage}: {substep}")


class VNextPipeline:
    """Main pipeline orchestrator for AutoWord vNext."""
    
    def __init__(self, 
                 llm_client: Optional[LLMClient] = None,
                 base_audit_dir: Optional[str] = None,
                 visible: bool = False,
                 progress_callback: Optional[Callable[[str, int], None]] = None,
                 monitoring_level: MonitoringLevel = MonitoringLevel.DETAILED,
                 enable_memory_monitoring: bool = True,
                 memory_warning_threshold_mb: float = 1024,
                 memory_critical_threshold_mb: float = 2048):
        """
        Initialize vNext pipeline.
        
        Args:
            llm_client: LLM client for plan generation (creates default if None)
            base_audit_dir: Base directory for audit trails
            visible: Whether to show Word application windows
            progress_callback: Optional callback for progress updates
            monitoring_level: Level of monitoring detail
            enable_memory_monitoring: Whether to enable memory monitoring
            memory_warning_threshold_mb: Memory warning threshold in MB
            memory_critical_threshold_mb: Memory critical threshold in MB
        """
        self.llm_client = llm_client
        self.base_audit_dir = base_audit_dir or "./audit_trails"
        self.visible = visible
        self.progress_reporter = ProgressReporter(progress_callback)
        self.monitoring_level = monitoring_level
        self.enable_memory_monitoring = enable_memory_monitoring
        self.memory_warning_threshold_mb = memory_warning_threshold_mb
        self.memory_critical_threshold_mb = memory_critical_threshold_mb
        
        # Initialize components (will be created per run)
        self.extractor: Optional[DocumentExtractor] = None
        self.planner: Optional[DocumentPlanner] = None
        self.executor: Optional[DocumentExecutor] = None
        self.validator: Optional[DocumentValidator] = None
        self.auditor: Optional[DocumentAuditor] = None
        self.error_handler: Optional[PipelineErrorHandler] = None
        self.vnext_logger: Optional[VNextLogger] = None
        
        # Run state
        self.current_audit_dir: Optional[str] = None
        self.original_docx_path: Optional[str] = None
        self.working_docx_path: Optional[str] = None
        self.temp_dir: Optional[str] = None
    
    def process_document(self, docx_path: str, user_intent: str) -> ProcessingResult:
        """
        Process document through complete vNext pipeline.
        
        Args:
            docx_path: Path to input DOCX file
            user_intent: User's intent description for LLM planning
            
        Returns:
            ProcessingResult with status and details
        """
        logger.info(f"Starting vNext pipeline processing for: {docx_path}")
        logger.info(f"User intent: {user_intent}")
        
        try:
            # Setup run environment (includes logger initialization)
            self._setup_run_environment(docx_path)
            
            # Log pipeline start
            self.vnext_logger.log_debug("Pipeline processing started", 
                                      docx_path=docx_path, 
                                      user_intent=user_intent,
                                      monitoring_level=self.monitoring_level.value)
            
            # Stage 1: Extract
            self.progress_reporter.start_stage("Extract")
            with self.vnext_logger.track_stage("Extract"):
                structure, inventory = self._extract_document()
            self.progress_reporter.complete_stage()
            
            # Stage 2: Plan
            self.progress_reporter.start_stage("Plan")
            with self.vnext_logger.track_stage("Plan"):
                plan = self._generate_plan(structure, user_intent)
            self.progress_reporter.complete_stage()
            
            # Stage 3: Execute
            self.progress_reporter.start_stage("Execute")
            with self.vnext_logger.track_stage("Execute"):
                modified_docx_path = self._execute_plan(plan)
            self.progress_reporter.complete_stage()
            
            # Stage 4: Validate
            self.progress_reporter.start_stage("Validate")
            with self.vnext_logger.track_stage("Validate"):
                validation_result = self._validate_modifications(structure, modified_docx_path)
                if not validation_result.is_valid:
                    return self._handle_validation_failure(validation_result)
            self.progress_reporter.complete_stage()
            
            # Stage 5: Audit
            self.progress_reporter.start_stage("Audit")
            with self.vnext_logger.track_stage("Audit"):
                self._create_audit_trail(structure, inventory, plan, modified_docx_path, validation_result)
            self.progress_reporter.complete_stage()
            
            # Success - log final performance metrics
            self.vnext_logger.log_debug("Pipeline processing completed successfully")
            
            logger.info("Pipeline processing completed successfully")
            return ProcessingResult(
                status="SUCCESS",
                message="Document processed successfully",
                audit_directory=self.current_audit_dir,
                warnings=self.error_handler.warnings_logger.get_warnings() if self.error_handler else []
            )
            
        except Exception as e:
            logger.error(f"Pipeline processing failed: {str(e)}")
            if self.vnext_logger:
                self.vnext_logger.log_error(e, {"pipeline_stage": "overall", "docx_path": docx_path})
            return self._handle_pipeline_error(e)
        
        finally:
            self._cleanup_run_environment()
    
    def _setup_run_environment(self, docx_path: str):
        """Setup run environment with timestamped directories and working copies."""
        self.progress_reporter.report_substep("Setting up run environment")
        
        # Validate input file
        if not os.path.exists(docx_path):
            raise ExtractionError(f"Input DOCX file not found: {docx_path}")
        
        self.original_docx_path = os.path.abspath(docx_path)
        
        # Create timestamped audit directory
        self.auditor = DocumentAuditor(self.base_audit_dir)
        self.current_audit_dir = self.auditor.create_audit_directory()
        
        # Initialize comprehensive logging and monitoring
        self.vnext_logger = create_vnext_logger(
            audit_directory=self.current_audit_dir,
            monitoring_level=self.monitoring_level,
            enable_memory_monitoring=self.enable_memory_monitoring,
            memory_warning_threshold_mb=self.memory_warning_threshold_mb,
            memory_critical_threshold_mb=self.memory_critical_threshold_mb
        )
        
        # Initialize error handler
        self.error_handler = PipelineErrorHandler(self.current_audit_dir)
        
        # Create temporary working directory
        self.temp_dir = tempfile.mkdtemp(prefix="vnext_pipeline_")
        
        # Create working copy of DOCX
        working_filename = f"working_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
        self.working_docx_path = os.path.join(self.temp_dir, working_filename)
        shutil.copy2(self.original_docx_path, self.working_docx_path)
        
        # Log document size for large document warnings
        try:
            file_size_mb = os.path.getsize(self.original_docx_path) / (1024 * 1024)
            log_large_document_warning(self.vnext_logger, file_size_mb, 0)  # paragraph count will be updated after extraction
        except Exception:
            pass
        
        self.vnext_logger.log_debug("Run environment setup complete", 
                                   original_docx=self.original_docx_path,
                                   working_docx=self.working_docx_path,
                                   audit_directory=self.current_audit_dir,
                                   temp_directory=self.temp_dir)
        
        logger.debug(f"Run environment setup complete:")
        logger.debug(f"  Original DOCX: {self.original_docx_path}")
        logger.debug(f"  Working DOCX: {self.working_docx_path}")
        logger.debug(f"  Audit directory: {self.current_audit_dir}")
        logger.debug(f"  Temp directory: {self.temp_dir}")
    
    def _extract_document(self) -> tuple[StructureV1, InventoryFullV1]:
        """Extract document structure and inventory."""
        with self.vnext_logger.track_operation("document_extraction") as metrics:
            self.progress_reporter.report_substep("Initializing extractor")
            
            try:
                with self.vnext_logger.track_operation("extractor_initialization"):
                    self.extractor = DocumentExtractor(visible=self.visible)
                
                with self.extractor:
                    self.progress_reporter.report_substep("Extracting document structure")
                    with self.vnext_logger.track_operation("structure_extraction"):
                        structure = self.extractor.extract_structure(self.working_docx_path)
                    
                    self.progress_reporter.report_substep("Extracting document inventory")
                    with self.vnext_logger.track_operation("inventory_extraction"):
                        inventory = self.extractor.extract_inventory(self.working_docx_path)
                    
                    # Log extraction results and check for complex scenarios
                    extraction_stats = {
                        "paragraphs": len(structure.paragraphs),
                        "headings": len(structure.headings),
                        "styles": len(structure.styles),
                        "tables": len(structure.tables),
                        "fields": len(structure.fields),
                        "content_controls": len(inventory.content_controls),
                        "formulas": len(inventory.formulas),
                        "charts": len(inventory.charts),
                        "footnotes": len(inventory.footnotes),
                        "endnotes": len(inventory.endnotes),
                        "cross_references": len(inventory.cross_references)
                    }
                    
                    self.vnext_logger.log_debug("Document extraction completed", **extraction_stats)
                    
                    # Check for complex document scenarios
                    if len(structure.tables) > 10:
                        log_complex_document_scenario(self.vnext_logger, "many_tables", 
                                                    {"table_count": len(structure.tables)})
                    
                    if len(inventory.formulas) > 0:
                        log_complex_document_scenario(self.vnext_logger, "contains_formulas", 
                                                    {"formula_count": len(inventory.formulas)})
                    
                    if len(inventory.charts) > 0:
                        log_complex_document_scenario(self.vnext_logger, "contains_charts", 
                                                    {"chart_count": len(inventory.charts)})
                    
                    if len(inventory.cross_references) > 20:
                        log_complex_document_scenario(self.vnext_logger, "many_cross_references", 
                                                    {"cross_ref_count": len(inventory.cross_references)})
                    
                    # Update large document warning with paragraph count
                    try:
                        file_size_mb = os.path.getsize(self.working_docx_path) / (1024 * 1024)
                        log_large_document_warning(self.vnext_logger, file_size_mb, len(structure.paragraphs))
                    except Exception:
                        pass
                    
                    logger.info(f"Extraction completed: {len(structure.paragraphs)} paragraphs, "
                               f"{len(structure.headings)} headings, {len(structure.styles)} styles")
                    
                    return structure, inventory
                    
            except Exception as e:
                context = ErrorContext(
                    pipeline_stage="Extract",
                    docx_path=self.working_docx_path,
                    original_docx_path=self.original_docx_path
                )
                self.vnext_logger.log_error(e, {"stage": "extraction", "docx_path": self.working_docx_path})
                raise ExtractionError(f"Document extraction failed: {str(e)}", 
                                    docx_path=self.working_docx_path,
                                    extraction_stage="structure_and_inventory") from e
    
    def _generate_plan(self, structure: StructureV1, user_intent: str) -> PlanV1:
        """Generate execution plan through LLM."""
        with self.vnext_logger.track_operation("plan_generation", user_intent=user_intent) as metrics:
            self.progress_reporter.report_substep("Initializing planner")
            
            try:
                with self.vnext_logger.track_operation("planner_initialization"):
                    self.planner = DocumentPlanner(
                        llm_client=self.llm_client,
                        schema_path=None  # Use default schema
                    )
                
                self.progress_reporter.report_substep("Generating plan through LLM")
                with self.vnext_logger.track_operation("llm_plan_generation"):
                    plan = self.planner.generate_plan(structure, user_intent)
                
                self.progress_reporter.report_substep("Validating plan schema and constraints")
                with self.vnext_logger.track_operation("plan_validation"):
                    validation_result = self.planner.validate_plan_schema(plan.model_dump())
                    if not validation_result.is_valid:
                        self.vnext_logger.log_error(
                            PlanningError("Plan validation failed", validation_errors=validation_result.errors),
                            {"validation_errors": validation_result.errors, "plan_data": plan.model_dump()}
                        )
                        raise PlanningError("Plan validation failed", 
                                          validation_errors=validation_result.errors)
                
                # Log plan details
                operation_types = [op.operation_type for op in plan.ops]
                operation_counts = {}
                for op_type in operation_types:
                    operation_counts[op_type] = operation_counts.get(op_type, 0) + 1
                
                self.vnext_logger.log_debug("Plan generation completed", 
                                          total_operations=len(plan.ops),
                                          operation_counts=operation_counts,
                                          user_intent=user_intent)
                
                logger.info(f"Plan generation completed: {len(plan.ops)} operations")
                
                return plan
                
            except Exception as e:
                context = ErrorContext(
                    pipeline_stage="Plan",
                    docx_path=self.working_docx_path,
                    original_docx_path=self.original_docx_path
                )
                self.vnext_logger.log_error(e, {"stage": "planning", "user_intent": user_intent})
                if isinstance(e, PlanningError):
                    raise
                else:
                    raise PlanningError(f"Plan generation failed: {str(e)}") from e
    
    def _execute_plan(self, plan: PlanV1) -> str:
        """Execute plan through atomic operations."""
        with self.vnext_logger.track_operation("plan_execution", operation_count=len(plan.ops)) as metrics:
            self.progress_reporter.report_substep("Initializing executor")
            
            try:
                with self.vnext_logger.track_operation("executor_initialization"):
                    warnings_log_path = os.path.join(self.current_audit_dir, "warnings.log")
                    self.executor = DocumentExecutor(warnings_log_path=warnings_log_path)
                
                self.progress_reporter.report_substep(f"Executing {len(plan.ops)} atomic operations")
                
                # Execute plan on working copy with individual operation tracking
                with self.vnext_logger.track_operation("atomic_operations_execution"):
                    # Log each operation individually for detailed monitoring
                    for i, operation in enumerate(plan.ops):
                        operation_name = f"{operation.operation_type}_{i+1}"
                        with self.vnext_logger.track_operation(operation_name, 
                                                             operation_type=operation.operation_type,
                                                             operation_index=i+1,
                                                             operation_data=operation.model_dump()):
                            pass  # The actual execution happens in executor.execute_plan
                    
                    result_docx_path = self.executor.execute_plan(plan, self.working_docx_path)
                
                # Log execution results
                execution_stats = {
                    "total_operations": len(plan.ops),
                    "result_docx_path": result_docx_path,
                    "warnings_log": warnings_log_path
                }
                
                self.vnext_logger.log_debug("Plan execution completed", **execution_stats)
                
                logger.info(f"Plan execution completed: {len(plan.ops)} operations executed")
                
                return result_docx_path
                
            except Exception as e:
                context = ErrorContext(
                    pipeline_stage="Execute",
                    operation_type=getattr(e, 'operation_type', None),
                    docx_path=self.working_docx_path,
                    original_docx_path=self.original_docx_path
                )
                self.vnext_logger.log_error(e, {"stage": "execution", "operation_count": len(plan.ops)})
                if isinstance(e, ExecutionError):
                    raise
                else:
                    raise ExecutionError(f"Plan execution failed: {str(e)}") from e
    
    def _validate_modifications(self, original_structure: StructureV1, modified_docx_path: str):
        """Validate document modifications against assertions."""
        with self.vnext_logger.track_operation("document_validation") as metrics:
            self.progress_reporter.report_substep("Initializing validator")
            
            try:
                with self.vnext_logger.track_operation("validator_initialization"):
                    self.validator = DocumentValidator(visible=self.visible)
                
                with self.validator:
                    self.progress_reporter.report_substep("Running validation assertions")
                    with self.vnext_logger.track_operation("validation_assertions"):
                        validation_result = self.validator.validate_modifications(
                            original_structure, modified_docx_path
                        )
                    
                    # Log validation results
                    validation_stats = {
                        "is_valid": validation_result.is_valid,
                        "error_count": len(validation_result.errors),
                        "warning_count": len(validation_result.warnings),
                        "errors": validation_result.errors,
                        "warnings": validation_result.warnings
                    }
                    
                    if validation_result.is_valid:
                        self.vnext_logger.log_debug("Validation completed successfully", **validation_stats)
                        logger.info("Validation completed successfully")
                    else:
                        self.vnext_logger.log_warning(f"Validation failed: {len(validation_result.errors)} errors", 
                                                    validation_stats)
                        logger.warning(f"Validation failed: {len(validation_result.errors)} errors")
                        for error in validation_result.errors:
                            logger.warning(f"  - {error}")
                    
                    return validation_result
                    
            except Exception as e:
                context = ErrorContext(
                    pipeline_stage="Validate",
                    docx_path=modified_docx_path,
                    original_docx_path=self.original_docx_path
                )
                self.vnext_logger.log_error(e, {"stage": "validation", "modified_docx": modified_docx_path})
                if isinstance(e, ValidationError):
                    raise
                else:
                    raise ValidationError(f"Document validation failed: {str(e)}") from e
    
    def _create_audit_trail(self, structure: StructureV1, inventory: InventoryFullV1, 
                           plan: PlanV1, modified_docx_path: str, validation_result):
        """Create complete audit trail with snapshots and reports."""
        with self.vnext_logger.track_operation("audit_trail_creation") as metrics:
            self.progress_reporter.report_substep("Creating audit snapshots")
            
            try:
                with self.vnext_logger.track_operation("audit_snapshots"):
                    # Save snapshots
                    self.auditor.save_snapshots(
                        before_docx=self.original_docx_path,
                        after_docx=modified_docx_path,
                        before_structure=structure,
                        after_structure=validation_result.modified_structure,
                        plan=plan
                    )
                
                self.progress_reporter.report_substep("Generating diff report")
                with self.vnext_logger.track_operation("diff_report_generation"):
                    diff_report = self.auditor.generate_diff_report(
                        structure, validation_result.modified_structure
                    )
                
                self.progress_reporter.report_substep("Writing audit status")
                with self.vnext_logger.track_operation("audit_status_writing"):
                    self.auditor.write_status("SUCCESS", "Pipeline completed successfully")
                
                # Log audit trail details
                audit_stats = {
                    "audit_directory": self.current_audit_dir,
                    "before_docx": self.original_docx_path,
                    "after_docx": modified_docx_path,
                    "diff_summary": diff_report.summary if hasattr(diff_report, 'summary') else "Generated"
                }
                
                self.vnext_logger.log_debug("Audit trail created", **audit_stats)
                
                logger.info(f"Audit trail created in: {self.current_audit_dir}")
                
            except Exception as e:
                self.vnext_logger.log_error(e, {"stage": "audit", "audit_directory": self.current_audit_dir})
                if isinstance(e, AuditError):
                    raise
                else:
                    raise AuditError(f"Audit trail creation failed: {str(e)}") from e
    
    def _handle_validation_failure(self, validation_result) -> ProcessingResult:
        """Handle validation failure with rollback."""
        logger.warning("Validation failed, performing rollback")
        
        try:
            # Perform rollback through error handler
            context = ErrorContext(
                pipeline_stage="Validate",
                docx_path=self.working_docx_path,
                original_docx_path=self.original_docx_path
            )
            
            recovery_result = self.error_handler.rollback_manager.perform_rollback(
                self.original_docx_path,
                self.working_docx_path,
                "Validation assertions failed"
            )
            
            # Create audit trail for failed validation
            self.auditor.write_status("FAILED_VALIDATION", 
                                    f"Validation failed: {'; '.join(validation_result.errors)}")
            
            return ProcessingResult(
                status="FAILED_VALIDATION",
                message="Document validation failed, changes rolled back",
                errors=validation_result.errors,
                audit_directory=self.current_audit_dir
            )
            
        except Exception as e:
            logger.error(f"Rollback failed: {str(e)}")
            return ProcessingResult(
                status="ROLLBACK",
                message="Validation failed and rollback encountered errors",
                errors=validation_result.errors + [f"Rollback error: {str(e)}"],
                audit_directory=self.current_audit_dir
            )
    
    def _handle_pipeline_error(self, error: Exception) -> ProcessingResult:
        """Handle pipeline error with comprehensive error handling."""
        logger.error(f"Pipeline error: {str(error)}")
        
        try:
            # Determine pipeline stage from error type
            if isinstance(error, ExtractionError):
                stage = "Extract"
            elif isinstance(error, PlanningError):
                stage = "Plan"
            elif isinstance(error, ExecutionError):
                stage = "Execute"
            elif isinstance(error, ValidationError):
                stage = "Validate"
            elif isinstance(error, AuditError):
                stage = "Audit"
            else:
                stage = "Unknown"
            
            context = ErrorContext(
                pipeline_stage=stage,
                docx_path=self.working_docx_path,
                original_docx_path=self.original_docx_path
            )
            
            # Handle error through error handler
            if self.error_handler:
                recovery_result = self.error_handler.handle_pipeline_error(error, context)
                
                # Create audit trail for error
                if self.auditor:
                    self.auditor.write_status("ROLLBACK", f"Pipeline error: {str(error)}")
                
                return ProcessingResult(
                    status=recovery_result.status.value,
                    message=f"Pipeline error in {stage} stage",
                    errors=[str(error)] + recovery_result.errors,
                    warnings=recovery_result.warnings,
                    audit_directory=self.current_audit_dir
                )
            else:
                # Fallback error handling
                return ProcessingResult(
                    status="ROLLBACK",
                    message=f"Pipeline error in {stage} stage",
                    errors=[str(error)],
                    audit_directory=self.current_audit_dir
                )
                
        except Exception as handler_error:
            logger.error(f"Error handler failed: {str(handler_error)}")
            return ProcessingResult(
                status="ROLLBACK",
                message="Pipeline error and error handler failed",
                errors=[str(error), f"Handler error: {str(handler_error)}"],
                audit_directory=self.current_audit_dir
            )
    
    def _cleanup_run_environment(self):
        """Cleanup temporary files and resources."""
        try:
            # Cleanup monitoring and save final reports
            if self.vnext_logger:
                try:
                    self.vnext_logger.cleanup()
                except Exception as e:
                    logger.warning(f"Failed to cleanup monitoring: {str(e)}")
            
            # Cleanup temp directory
            if self.temp_dir and os.path.exists(self.temp_dir):
                shutil.rmtree(self.temp_dir)
                logger.debug(f"Cleaned up temp directory: {self.temp_dir}")
        except Exception as e:
            logger.warning(f"Failed to cleanup temp directory: {str(e)}")
        
        # Reset run state
        self.current_audit_dir = None
        self.original_docx_path = None
        self.working_docx_path = None
        self.temp_dir = None
        
        # Clear component references
        self.extractor = None
        self.planner = None
        self.executor = None
        self.validator = None
        self.auditor = None
        self.error_handler = None
        self.vnext_logger = None