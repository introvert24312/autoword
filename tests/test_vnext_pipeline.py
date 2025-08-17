"""
Integration tests for VNext Pipeline orchestrator.

This module tests the complete pipeline execution with all five modules
in sequence: Extract→Plan→Execute→Validate→Audit.
"""

import os
import pytest
import tempfile
import shutil
from unittest.mock import Mock, patch, MagicMock, call
from pathlib import Path

from autoword.vnext.pipeline import VNextPipeline, ProgressReporter
from autoword.vnext.models import (
    StructureV1, PlanV1, InventoryFullV1, ProcessingResult, 
    DocumentMetadata, StyleDefinition, ParagraphSkeleton, HeadingReference,
    DeleteSectionByHeading, ValidationResult
)
from autoword.vnext.exceptions import (
    ExtractionError, PlanningError, ExecutionError, ValidationError, AuditError
)
from autoword.core.llm_client import LLMClient


class TestProgressReporter:
    """Test progress reporting functionality."""
    
    def test_progress_reporter_initialization(self):
        """Test progress reporter initialization."""
        callback = Mock()
        reporter = ProgressReporter(callback)
        
        assert reporter.progress_callback == callback
        assert reporter.current_stage == ""
        assert reporter.total_stages == 5
        assert reporter.completed_stages == 0
    
    def test_start_stage(self):
        """Test starting a pipeline stage."""
        callback = Mock()
        reporter = ProgressReporter(callback)
        
        reporter.start_stage("Extract")
        
        assert reporter.current_stage == "Extract"
        callback.assert_called_once_with("Extract", 0)
    
    def test_complete_stage(self):
        """Test completing a pipeline stage."""
        callback = Mock()
        reporter = ProgressReporter(callback)
        
        reporter.start_stage("Extract")
        reporter.complete_stage()
        
        assert reporter.completed_stages == 1
        assert callback.call_count == 2
        callback.assert_has_calls([
            call("Extract", 0),
            call("Extract", 20)
        ])
    
    def test_report_substep(self):
        """Test reporting substeps."""
        reporter = ProgressReporter()
        reporter.current_stage = "Extract"
        
        # Should not raise any exceptions
        reporter.report_substep("Extracting structure")


class TestVNextPipeline:
    """Test VNext pipeline orchestrator."""
    
    def setup_method(self):
        """Setup test environment."""
        self.temp_dir = tempfile.mkdtemp()
        self.test_docx = os.path.join(self.temp_dir, "test.docx")
        self.audit_dir = os.path.join(self.temp_dir, "audit")
        
        # Create a dummy DOCX file
        with open(self.test_docx, 'wb') as f:
            f.write(b"dummy docx content")
        
        self.llm_client = Mock(spec=LLMClient)
        self.pipeline = VNextPipeline(
            llm_client=self.llm_client,
            base_audit_dir=self.audit_dir,
            visible=False
        )
    
    def teardown_method(self):
        """Cleanup test environment."""
        if os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir)
    
    def test_pipeline_initialization(self):
        """Test pipeline initialization."""
        assert self.pipeline.llm_client == self.llm_client
        assert self.pipeline.base_audit_dir == self.audit_dir
        assert self.pipeline.visible is False
        assert isinstance(self.pipeline.progress_reporter, ProgressReporter)
        
        # Components should be None initially
        assert self.pipeline.extractor is None
        assert self.pipeline.planner is None
        assert self.pipeline.executor is None
        assert self.pipeline.validator is None
        assert self.pipeline.auditor is None
        assert self.pipeline.error_handler is None
    
    def test_pipeline_initialization_with_progress_callback(self):
        """Test pipeline initialization with progress callback."""
        callback = Mock()
        pipeline = VNextPipeline(progress_callback=callback)
        
        assert pipeline.progress_reporter.progress_callback == callback
    
    @patch('autoword.vnext.pipeline.DocumentAuditor')
    @patch('autoword.vnext.pipeline.PipelineErrorHandler')
    @patch('tempfile.mkdtemp')
    @patch('shutil.copy2')
    def test_setup_run_environment(self, mock_copy, mock_mkdtemp, mock_error_handler, mock_auditor):
        """Test run environment setup."""
        # Setup mocks
        mock_mkdtemp.return_value = "/tmp/test"
        mock_auditor_instance = Mock()
        mock_auditor_instance.create_audit_directory.return_value = "/audit/run_123"
        mock_auditor.return_value = mock_auditor_instance
        
        mock_error_handler_instance = Mock()
        mock_error_handler.return_value = mock_error_handler_instance
        
        # Execute
        self.pipeline._setup_run_environment(self.test_docx)
        
        # Verify
        assert self.pipeline.original_docx_path == os.path.abspath(self.test_docx)
        assert self.pipeline.current_audit_dir == "/audit/run_123"
        assert self.pipeline.temp_dir == "/tmp/test"
        assert self.pipeline.auditor == mock_auditor_instance
        assert self.pipeline.error_handler == mock_error_handler_instance
        
        mock_auditor.assert_called_once_with(self.audit_dir)
        mock_auditor_instance.create_audit_directory.assert_called_once()
        mock_error_handler.assert_called_once_with("/audit/run_123")
        mock_copy.assert_called_once()
    
    def test_setup_run_environment_missing_file(self):
        """Test setup with missing input file."""
        with pytest.raises(ExtractionError, match="Input DOCX file not found"):
            self.pipeline._setup_run_environment("/nonexistent/file.docx")
    
    @patch('autoword.vnext.pipeline.DocumentExtractor')
    def test_extract_document_success(self, mock_extractor_class):
        """Test successful document extraction."""
        # Setup mocks
        mock_extractor = MagicMock()
        mock_structure = Mock(spec=StructureV1)
        mock_structure.paragraphs = [Mock(), Mock()]
        mock_structure.headings = [Mock()]
        mock_structure.styles = [Mock(), Mock(), Mock()]
        mock_inventory = Mock(spec=InventoryFullV1)
        
        mock_extractor.extract_structure.return_value = mock_structure
        mock_extractor.extract_inventory.return_value = mock_inventory
        mock_extractor.__enter__.return_value = mock_extractor
        mock_extractor.__exit__.return_value = None
        
        mock_extractor_class.return_value = mock_extractor
        
        # Setup pipeline state
        self.pipeline.working_docx_path = self.test_docx
        
        # Execute
        structure, inventory = self.pipeline._extract_document()
        
        # Verify
        assert structure == mock_structure
        assert inventory == mock_inventory
        
        mock_extractor_class.assert_called_once_with(visible=False)
        mock_extractor.extract_structure.assert_called_once_with(self.test_docx)
        mock_extractor.extract_inventory.assert_called_once_with(self.test_docx)
    
    @patch('autoword.vnext.pipeline.DocumentExtractor')
    def test_extract_document_failure(self, mock_extractor_class):
        """Test document extraction failure."""
        # Setup mocks
        mock_extractor = MagicMock()
        mock_extractor.extract_structure.side_effect = Exception("COM error")
        mock_extractor.__enter__.return_value = mock_extractor
        mock_extractor.__exit__.return_value = None
        
        mock_extractor_class.return_value = mock_extractor
        
        # Setup pipeline state
        self.pipeline.working_docx_path = self.test_docx
        
        # Execute and verify
        with pytest.raises(ExtractionError, match="Document extraction failed"):
            self.pipeline._extract_document()
    
    @patch('autoword.vnext.pipeline.DocumentPlanner')
    def test_generate_plan_success(self, mock_planner_class):
        """Test successful plan generation."""
        # Setup mocks
        mock_planner = Mock()
        mock_plan = Mock(spec=PlanV1)
        mock_plan.ops = [Mock()]  # Add ops attribute
        mock_validation_result = Mock()
        mock_validation_result.is_valid = True
        
        mock_planner.generate_plan.return_value = mock_plan
        mock_planner.validate_plan_schema.return_value = mock_validation_result
        mock_planner_class.return_value = mock_planner
        
        mock_structure = Mock(spec=StructureV1)
        user_intent = "Remove abstract and references"
        
        # Execute
        plan = self.pipeline._generate_plan(mock_structure, user_intent)
        
        # Verify
        assert plan == mock_plan
        
        mock_planner_class.assert_called_once_with(
            llm_client=self.llm_client,
            schema_path=None
        )
        mock_planner.generate_plan.assert_called_once_with(mock_structure, user_intent)
        mock_planner.validate_plan_schema.assert_called_once()
    
    @patch('autoword.vnext.pipeline.DocumentPlanner')
    def test_generate_plan_validation_failure(self, mock_planner_class):
        """Test plan generation with validation failure."""
        # Setup mocks
        mock_planner = Mock()
        mock_plan = Mock(spec=PlanV1)
        mock_validation_result = Mock()
        mock_validation_result.is_valid = False
        mock_validation_result.errors = ["Invalid operation type"]
        
        mock_planner.generate_plan.return_value = mock_plan
        mock_planner.validate_plan_schema.return_value = mock_validation_result
        mock_planner_class.return_value = mock_planner
        
        mock_structure = Mock(spec=StructureV1)
        user_intent = "Remove abstract and references"
        
        # Execute and verify
        with pytest.raises(PlanningError, match="Plan validation failed"):
            self.pipeline._generate_plan(mock_structure, user_intent)
    
    @patch('autoword.vnext.pipeline.DocumentExecutor')
    def test_execute_plan_success(self, mock_executor_class):
        """Test successful plan execution."""
        # Setup mocks
        mock_executor = Mock()
        mock_executor.execute_plan.return_value = "/path/to/modified.docx"
        mock_executor_class.return_value = mock_executor
        
        mock_plan = Mock(spec=PlanV1)
        mock_plan.ops = [Mock()]
        
        # Setup pipeline state
        self.pipeline.working_docx_path = self.test_docx
        self.pipeline.current_audit_dir = "/audit/run_123"
        
        # Execute
        result_path = self.pipeline._execute_plan(mock_plan)
        
        # Verify
        assert result_path == "/path/to/modified.docx"
        
        expected_warnings_path = os.path.join("/audit/run_123", "warnings.log")
        mock_executor_class.assert_called_once_with(warnings_log_path=expected_warnings_path)
        mock_executor.execute_plan.assert_called_once_with(mock_plan, self.test_docx)
    
    @patch('autoword.vnext.pipeline.DocumentValidator')
    def test_validate_modifications_success(self, mock_validator_class):
        """Test successful validation."""
        # Setup mocks
        mock_validator = MagicMock()
        mock_validation_result = Mock()
        mock_validation_result.is_valid = True
        
        mock_validator.validate_modifications.return_value = mock_validation_result
        mock_validator.__enter__.return_value = mock_validator
        mock_validator.__exit__.return_value = None
        
        mock_validator_class.return_value = mock_validator
        
        mock_structure = Mock(spec=StructureV1)
        modified_docx_path = "/path/to/modified.docx"
        
        # Execute
        result = self.pipeline._validate_modifications(mock_structure, modified_docx_path)
        
        # Verify
        assert result == mock_validation_result
        
        mock_validator_class.assert_called_once_with(visible=False)
        mock_validator.validate_modifications.assert_called_once_with(
            mock_structure, modified_docx_path
        )
    
    @patch('autoword.vnext.pipeline.DocumentValidator')
    def test_validate_modifications_failure(self, mock_validator_class):
        """Test validation failure."""
        # Setup mocks
        mock_validator = MagicMock()
        mock_validation_result = Mock()
        mock_validation_result.is_valid = False
        mock_validation_result.errors = ["Style assertion failed"]
        
        mock_validator.validate_modifications.return_value = mock_validation_result
        mock_validator.__enter__.return_value = mock_validator
        mock_validator.__exit__.return_value = None
        
        mock_validator_class.return_value = mock_validator
        
        mock_structure = Mock(spec=StructureV1)
        modified_docx_path = "/path/to/modified.docx"
        
        # Execute
        result = self.pipeline._validate_modifications(mock_structure, modified_docx_path)
        
        # Verify
        assert result == mock_validation_result
        assert not result.is_valid
    
    def test_create_audit_trail_success(self):
        """Test successful audit trail creation."""
        # Setup mocks
        mock_auditor = Mock()
        mock_structure = Mock(spec=StructureV1)
        mock_inventory = Mock(spec=InventoryFullV1)
        mock_plan = Mock(spec=PlanV1)
        mock_validation_result = Mock()
        mock_validation_result.modified_structure = Mock(spec=StructureV1)
        
        self.pipeline.auditor = mock_auditor
        self.pipeline.original_docx_path = "/original.docx"
        
        # Execute
        self.pipeline._create_audit_trail(
            mock_structure, mock_inventory, mock_plan, 
            "/modified.docx", mock_validation_result
        )
        
        # Verify
        mock_auditor.save_snapshots.assert_called_once_with(
            before_docx="/original.docx",
            after_docx="/modified.docx",
            before_structure=mock_structure,
            after_structure=mock_validation_result.modified_structure,
            plan=mock_plan
        )
        mock_auditor.generate_diff_report.assert_called_once()
        mock_auditor.write_status.assert_called_once_with("SUCCESS", "Pipeline completed successfully")
    
    def test_handle_validation_failure(self):
        """Test handling validation failure with rollback."""
        # Setup mocks
        mock_error_handler = Mock()
        mock_rollback_manager = Mock()
        mock_recovery_result = Mock()
        mock_recovery_result.rollback_performed = True
        
        mock_rollback_manager.perform_rollback.return_value = mock_recovery_result
        mock_error_handler.rollback_manager = mock_rollback_manager
        
        mock_auditor = Mock()
        
        self.pipeline.error_handler = mock_error_handler
        self.pipeline.auditor = mock_auditor
        self.pipeline.current_audit_dir = "/audit/run_123"
        self.pipeline.working_docx_path = "/working.docx"
        self.pipeline.original_docx_path = "/original.docx"
        
        mock_validation_result = Mock()
        mock_validation_result.errors = ["Style assertion failed", "TOC mismatch"]
        
        # Execute
        result = self.pipeline._handle_validation_failure(mock_validation_result)
        
        # Verify
        assert result.status == "FAILED_VALIDATION"
        assert result.message == "Document validation failed, changes rolled back"
        assert result.errors == ["Style assertion failed", "TOC mismatch"]
        assert result.audit_directory == "/audit/run_123"
        
        mock_rollback_manager.perform_rollback.assert_called_once_with(
            "/original.docx", "/working.docx", "Validation assertions failed"
        )
        mock_auditor.write_status.assert_called_once_with(
            "FAILED_VALIDATION", "Validation failed: Style assertion failed; TOC mismatch"
        )
    
    def test_handle_pipeline_error_extraction(self):
        """Test handling extraction error."""
        error = ExtractionError("COM initialization failed")
        
        # Setup mocks
        mock_error_handler = Mock()
        mock_recovery_result = Mock()
        mock_recovery_result.status.value = "ROLLBACK"
        mock_recovery_result.errors = ["Rollback completed"]
        mock_recovery_result.warnings = ["COM error occurred"]
        
        mock_error_handler.handle_pipeline_error.return_value = mock_recovery_result
        
        mock_auditor = Mock()
        
        self.pipeline.error_handler = mock_error_handler
        self.pipeline.auditor = mock_auditor
        self.pipeline.current_audit_dir = "/audit/run_123"
        
        # Execute
        result = self.pipeline._handle_pipeline_error(error)
        
        # Verify
        assert result.status == "ROLLBACK"
        assert result.message == "Pipeline error in Extract stage"
        assert "COM initialization failed" in result.errors
        assert result.warnings == ["COM error occurred"]
        assert result.audit_directory == "/audit/run_123"
        
        mock_auditor.write_status.assert_called_once_with(
            "ROLLBACK", "Pipeline error: COM initialization failed"
        )
    
    @patch('shutil.rmtree')
    @patch('os.path.exists')
    def test_cleanup_run_environment(self, mock_exists, mock_rmtree):
        """Test cleanup of run environment."""
        # Setup pipeline state
        self.pipeline.temp_dir = "/tmp/test"
        self.pipeline.current_audit_dir = "/audit/run_123"
        self.pipeline.original_docx_path = "/original.docx"
        self.pipeline.working_docx_path = "/working.docx"
        self.pipeline.extractor = Mock()
        
        # Mock exists to return True so cleanup is attempted
        mock_exists.return_value = True
        
        # Execute
        self.pipeline._cleanup_run_environment()
        
        # Verify cleanup
        mock_exists.assert_called_once_with("/tmp/test")
        mock_rmtree.assert_called_once_with("/tmp/test")
        
        # Verify state reset
        assert self.pipeline.current_audit_dir is None
        assert self.pipeline.original_docx_path is None
        assert self.pipeline.working_docx_path is None
        assert self.pipeline.temp_dir is None
        assert self.pipeline.extractor is None
    
    @patch.object(VNextPipeline, '_setup_run_environment')
    @patch.object(VNextPipeline, '_extract_document')
    @patch.object(VNextPipeline, '_generate_plan')
    @patch.object(VNextPipeline, '_execute_plan')
    @patch.object(VNextPipeline, '_validate_modifications')
    @patch.object(VNextPipeline, '_create_audit_trail')
    @patch.object(VNextPipeline, '_cleanup_run_environment')
    def test_process_document_success_integration(self, mock_cleanup, mock_audit, 
                                                mock_validate, mock_execute, 
                                                mock_plan, mock_extract, mock_setup):
        """Test complete successful document processing integration."""
        # Setup mocks
        mock_structure = Mock(spec=StructureV1)
        mock_inventory = Mock(spec=InventoryFullV1)
        mock_plan_obj = Mock(spec=PlanV1)
        mock_validation_result = Mock()
        mock_validation_result.is_valid = True
        
        mock_extract.return_value = (mock_structure, mock_inventory)
        mock_plan.return_value = mock_plan_obj
        mock_execute.return_value = "/modified.docx"
        mock_validate.return_value = mock_validation_result
        
        # Setup pipeline state for audit directory
        self.pipeline.current_audit_dir = "/audit/run_123"
        
        # Execute
        result = self.pipeline.process_document(self.test_docx, "Remove abstract")
        
        # Verify result
        assert result.status == "SUCCESS"
        assert result.message == "Document processed successfully"
        assert result.audit_directory == "/audit/run_123"
        
        # Verify method calls
        mock_setup.assert_called_once_with(self.test_docx)
        mock_extract.assert_called_once()
        mock_plan.assert_called_once_with(mock_structure, "Remove abstract")
        mock_execute.assert_called_once_with(mock_plan_obj)
        mock_validate.assert_called_once_with(mock_structure, "/modified.docx")
        mock_audit.assert_called_once_with(
            mock_structure, mock_inventory, mock_plan_obj, 
            "/modified.docx", mock_validation_result
        )
        mock_cleanup.assert_called_once()
    
    @patch.object(VNextPipeline, '_setup_run_environment')
    @patch.object(VNextPipeline, '_extract_document')
    @patch.object(VNextPipeline, '_cleanup_run_environment')
    def test_process_document_extraction_failure(self, mock_cleanup, mock_extract, mock_setup):
        """Test document processing with extraction failure."""
        # Setup mocks
        extraction_error = ExtractionError("Document corrupted")
        mock_extract.side_effect = extraction_error
        
        # Setup pipeline state
        self.pipeline.current_audit_dir = "/audit/run_123"
        
        # Execute
        result = self.pipeline.process_document(self.test_docx, "Remove abstract")
        
        # Verify result
        assert result.status == "ROLLBACK"
        assert "Pipeline error in Extract stage" in result.message
        assert "Document corrupted" in result.errors
        
        # Verify cleanup was called
        mock_cleanup.assert_called_once()
    
    def test_process_document_with_progress_callback(self):
        """Test document processing with progress callback."""
        callback = Mock()
        pipeline = VNextPipeline(progress_callback=callback)
        
        # Mock all pipeline methods to avoid actual processing
        with patch.multiple(pipeline,
                          _setup_run_environment=Mock(),
                          _extract_document=Mock(return_value=(Mock(), Mock())),
                          _generate_plan=Mock(return_value=Mock()),
                          _execute_plan=Mock(return_value="/modified.docx"),
                          _validate_modifications=Mock(return_value=Mock(is_valid=True)),
                          _create_audit_trail=Mock(),
                          _cleanup_run_environment=Mock()):
            
            pipeline.current_audit_dir = "/audit/run_123"
            result = pipeline.process_document(self.test_docx, "Remove abstract")
        
        # Verify progress callbacks were made
        assert callback.call_count >= 5  # At least one call per stage
        
        # Verify specific stage calls
        stage_calls = [call[0][0] for call in callback.call_args_list]
        assert "Extract" in stage_calls
        assert "Plan" in stage_calls
        assert "Execute" in stage_calls
        assert "Validate" in stage_calls
        assert "Audit" in stage_calls


class TestPipelineIntegrationScenarios:
    """Test various integration scenarios for the pipeline."""
    
    def setup_method(self):
        """Setup test environment."""
        self.temp_dir = tempfile.mkdtemp()
        self.test_docx = os.path.join(self.temp_dir, "test.docx")
        
        # Create a dummy DOCX file
        with open(self.test_docx, 'wb') as f:
            f.write(b"dummy docx content")
    
    def teardown_method(self):
        """Cleanup test environment."""
        if os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir)
    
    def test_pipeline_with_invalid_plan(self):
        """Test pipeline handling of invalid LLM plan."""
        pipeline = VNextPipeline()
        
        with patch.multiple(pipeline,
                          _setup_run_environment=Mock(),
                          _extract_document=Mock(return_value=(Mock(), Mock())),
                          _generate_plan=Mock(side_effect=PlanningError("Invalid plan schema")),
                          _cleanup_run_environment=Mock()):
            
            pipeline.current_audit_dir = "/audit/run_123"
            result = pipeline.process_document(self.test_docx, "Invalid request")
        
        assert result.status == "ROLLBACK"
        assert "Invalid plan schema" in result.errors
    
    def test_pipeline_with_execution_error(self):
        """Test pipeline handling of execution error."""
        pipeline = VNextPipeline()
        
        with patch.multiple(pipeline,
                          _setup_run_environment=Mock(),
                          _extract_document=Mock(return_value=(Mock(), Mock())),
                          _generate_plan=Mock(return_value=Mock()),
                          _execute_plan=Mock(side_effect=ExecutionError("COM operation failed")),
                          _cleanup_run_environment=Mock()):
            
            pipeline.current_audit_dir = "/audit/run_123"
            result = pipeline.process_document(self.test_docx, "Remove abstract")
        
        assert result.status == "ROLLBACK"
        assert "COM operation failed" in result.errors
    
    def test_pipeline_with_validation_failure(self):
        """Test pipeline handling of validation failure."""
        pipeline = VNextPipeline()
        
        mock_validation_result = Mock()
        mock_validation_result.is_valid = False
        mock_validation_result.errors = ["TOC assertion failed"]
        
        with patch.multiple(pipeline,
                          _setup_run_environment=Mock(),
                          _extract_document=Mock(return_value=(Mock(), Mock())),
                          _generate_plan=Mock(return_value=Mock()),
                          _execute_plan=Mock(return_value="/modified.docx"),
                          _validate_modifications=Mock(return_value=mock_validation_result),
                          _handle_validation_failure=Mock(return_value=ProcessingResult(
                              status="FAILED_VALIDATION",
                              message="Validation failed",
                              errors=["TOC assertion failed"]
                          )),
                          _cleanup_run_environment=Mock()):
            
            result = pipeline.process_document(self.test_docx, "Remove abstract")
        
        assert result.status == "FAILED_VALIDATION"
        assert "TOC assertion failed" in result.errors


if __name__ == "__main__":
    pytest.main([__file__, "-v"])