"""
Integration tests for VNextPipeline orchestrator.

This module tests the complete pipeline execution with all five modules
in sequence: Extract→Plan→Execute→Validate→Audit with error handling
and rollback capabilities.
"""

import os
import tempfile
import shutil
import json
import pytest
from pathlib import Path
from unittest.mock import Mock, patch, MagicMock
from datetime import datetime

from autoword.vnext.pipeline import VNextPipeline, ProgressReporter
from autoword.vnext.models import (
    StructureV1, PlanV1, InventoryFullV1, ProcessingResult,
    DocumentMetadata, StyleDefinition, ParagraphSkeleton, HeadingReference,
    DeleteSectionByHeading, UpdateToc, SetStyleRule, FontSpec, ParagraphSpec,
    LineSpacingMode, StyleType
)
from autoword.vnext.exceptions import (
    VNextError, ExtractionError, PlanningError, ExecutionError, 
    ValidationError, AuditError
)
from autoword.core.llm_client import LLMClient


class TestProgressReporter:
    """Test progress reporting functionality."""
    
    def test_progress_reporter_initialization(self):
        """Test progress reporter initialization."""
        callback_calls = []
        
        def mock_callback(stage, progress):
            callback_calls.append((stage, progress))
        
        reporter = ProgressReporter(mock_callback)
        assert reporter.progress_callback == mock_callback
        assert reporter.current_stage == ""
        assert reporter.total_stages == 5
        assert reporter.completed_stages == 0
    
    def test_progress_reporting_flow(self):
        """Test complete progress reporting flow."""
        callback_calls = []
        
        def mock_callback(stage, progress):
            callback_calls.append((stage, progress))
        
        reporter = ProgressReporter(mock_callback)
        
        # Test stage progression
        stages = ["Extract", "Plan", "Execute", "Validate", "Audit"]
        expected_progress = [0, 20, 40, 60, 80, 100]
        
        for i, stage in enumerate(stages):
            reporter.start_stage(stage)
            assert reporter.current_stage == stage
            assert callback_calls[i*2] == (stage, expected_progress[i])
            
            reporter.complete_stage()
            assert reporter.completed_stages == i + 1
            assert callback_calls[i*2 + 1] == (stage, expected_progress[i + 1])
    
    def test_progress_reporter_without_callback(self):
        """Test progress reporter without callback function."""
        reporter = ProgressReporter()
        
        # Should not raise errors
        reporter.start_stage("Test")
        reporter.complete_stage()
        reporter.report_substep("Test substep")


class TestVNextPipelineInitialization:
    """Test VNextPipeline initialization and configuration."""
    
    def test_default_initialization(self):
        """Test pipeline initialization with default parameters."""
        pipeline = VNextPipeline()
        
        assert pipeline.llm_client is None
        assert pipeline.base_audit_dir == "./audit_trails"
        assert pipeline.visible is False
        assert pipeline.progress_reporter is not None
        assert pipeline.current_audit_dir is None
        assert pipeline.original_docx_path is None
        assert pipeline.working_docx_path is None
        assert pipeline.temp_dir is None
    
    def test_custom_initialization(self):
        """Test pipeline initialization with custom parameters."""
        mock_llm = Mock(spec=LLMClient)
        callback_calls = []
        
        def mock_callback(stage, progress):
            callback_calls.append((stage, progress))
        
        pipeline = VNextPipeline(
            llm_client=mock_llm,
            base_audit_dir="/custom/audit",
            visible=True,
            progress_callback=mock_callback
        )
        
        assert pipeline.llm_client == mock_llm
        assert pipeline.base_audit_dir == "/custom/audit"
        assert pipeline.visible is True
        assert pipeline.progress_reporter.progress_callback == mock_callback


class TestVNextPipelineSetup:
    """Test pipeline setup and environment management."""
    
    def test_setup_run_environment_success(self):
        """Test successful run environment setup."""
        with tempfile.TemporaryDirectory() as temp_dir:
            # Create a mock DOCX file
            test_docx = os.path.join(temp_dir, "test.docx")
            with open(test_docx, 'wb') as f:
                f.write(b"Mock DOCX content")
            
            pipeline = VNextPipeline(base_audit_dir=temp_dir)
            
            with patch('autoword.vnext.pipeline.DocumentAuditor') as mock_auditor_class:
                mock_auditor = Mock()
                mock_auditor.create_audit_directory.return_value = os.path.join(temp_dir, "audit_123")
                mock_auditor_class.return_value = mock_auditor
                
                with patch('autoword.vnext.pipeline.create_vnext_logger') as mock_logger:
                    with patch('autoword.vnext.pipeline.PipelineErrorHandler') as mock_error_handler:
                        pipeline._setup_run_environment(test_docx)
                        
                        assert pipeline.original_docx_path == os.path.abspath(test_docx)
                        assert pipeline.current_audit_dir == os.path.join(temp_dir, "audit_123")
                        assert pipeline.temp_dir is not None
                        assert pipeline.working_docx_path is not None
                        assert os.path.exists(pipeline.working_docx_path)
    
    def test_setup_run_environment_missing_file(self):
        """Test setup with missing input file."""
        pipeline = VNextPipeline()
        
        with pytest.raises(ExtractionError, match="Input DOCX file not found"):
            pipeline._setup_run_environment("/nonexistent/file.docx")
    
    def test_cleanup_run_environment(self):
        """Test run environment cleanup."""
        with tempfile.TemporaryDirectory() as temp_dir:
            pipeline = VNextPipeline()
            pipeline.temp_dir = temp_dir
            pipeline.current_audit_dir = "/some/audit/dir"
            pipeline.original_docx_path = "/some/original.docx"
            pipeline.working_docx_path = "/some/working.docx"
            
            # Mock components
            pipeline.vnext_logger = Mock()
            pipeline.extractor = Mock()
            pipeline.planner = Mock()
            
            pipeline._cleanup_run_environment()
            
            # Verify cleanup
            assert pipeline.current_audit_dir is None
            assert pipeline.original_docx_path is None
            assert pipeline.working_docx_path is None
            assert pipeline.temp_dir is None
            assert pipeline.extractor is None
            assert pipeline.planner is None
            assert pipeline.vnext_logger is None


class TestVNextPipelineStages:
    """Test individual pipeline stages."""
    
    def create_mock_structure(self):
        """Create a mock document structure for testing."""
        return StructureV1(
            metadata=DocumentMetadata(
                title="Test Document",
                author="Test Author",
                paragraph_count=10,
                word_count=100
            ),
            styles=[
                StyleDefinition(
                    name="Heading 1",
                    type=StyleType.PARAGRAPH,
                    font=FontSpec(east_asian="楷体", latin="Times New Roman", size_pt=14, bold=True),
                    paragraph=ParagraphSpec(line_spacing_mode=LineSpacingMode.MULTIPLE, line_spacing_value=2.0)
                )
            ],
            paragraphs=[
                ParagraphSkeleton(index=0, style_name="Heading 1", preview_text="Test Heading", is_heading=True, heading_level=1),
                ParagraphSkeleton(index=1, style_name="Normal", preview_text="Test paragraph content")
            ],
            headings=[
                HeadingReference(paragraph_index=0, level=1, text="Test Heading", style_name="Heading 1")
            ]
        )
    
    def create_mock_inventory(self):
        """Create a mock document inventory for testing."""
        return InventoryFullV1(
            ooxml_fragments={"test_fragment": "<w:p>Test</w:p>"},
            media_indexes={},
            content_controls=[],
            formulas=[],
            charts=[]
        )
    
    def create_mock_plan(self):
        """Create a mock execution plan for testing."""
        return PlanV1(
            ops=[
                DeleteSectionByHeading(
                    heading_text="摘要",
                    level=1,
                    match="EXACT",
                    case_sensitive=False
                ),
                UpdateToc(),
                SetStyleRule(
                    target_style_name="Heading 1",
                    font=FontSpec(east_asian="楷体", size_pt=14, bold=True),
                    paragraph=ParagraphSpec(line_spacing_mode=LineSpacingMode.MULTIPLE, line_spacing_value=2.0)
                )
            ]
        )
    
    def test_extract_document_success(self):
        """Test successful document extraction."""
        with tempfile.TemporaryDirectory() as temp_dir:
            pipeline = VNextPipeline()
            pipeline.working_docx_path = os.path.join(temp_dir, "test.docx")
            pipeline.vnext_logger = Mock()
            pipeline.progress_reporter = Mock()
            
            mock_structure = self.create_mock_structure()
            mock_inventory = self.create_mock_inventory()
            
            with patch('autoword.vnext.pipeline.DocumentExtractor') as mock_extractor_class:
                mock_extractor = Mock()
                mock_extractor.__enter__ = Mock(return_value=mock_extractor)
                mock_extractor.__exit__ = Mock(return_value=None)
                mock_extractor.extract_structure.return_value = mock_structure
                mock_extractor.extract_inventory.return_value = mock_inventory
                mock_extractor_class.return_value = mock_extractor
                
                structure, inventory = pipeline._extract_document()
                
                assert structure == mock_structure
                assert inventory == mock_inventory
                mock_extractor.extract_structure.assert_called_once_with(pipeline.working_docx_path)
                mock_extractor.extract_inventory.assert_called_once_with(pipeline.working_docx_path)
    
    def test_extract_document_failure(self):
        """Test document extraction failure."""
        pipeline = VNextPipeline()
        pipeline.working_docx_path = "/test/path.docx"
        pipeline.vnext_logger = Mock()
        pipeline.progress_reporter = Mock()
        
        with patch('autoword.vnext.pipeline.DocumentExtractor') as mock_extractor_class:
            mock_extractor = Mock()
            mock_extractor.__enter__ = Mock(return_value=mock_extractor)
            mock_extractor.__exit__ = Mock(return_value=None)
            mock_extractor.extract_structure.side_effect = Exception("Extraction failed")
            mock_extractor_class.return_value = mock_extractor
            
            with pytest.raises(ExtractionError, match="Document extraction failed"):
                pipeline._extract_document()
    
    def test_generate_plan_success(self):
        """Test successful plan generation."""
        pipeline = VNextPipeline()
        pipeline.vnext_logger = Mock()
        pipeline.progress_reporter = Mock()
        
        mock_structure = self.create_mock_structure()
        mock_plan = self.create_mock_plan()
        user_intent = "Remove abstract and references sections"
        
        with patch('autoword.vnext.pipeline.DocumentPlanner') as mock_planner_class:
            mock_planner = Mock()
            mock_planner.generate_plan.return_value = mock_plan
            mock_planner.validate_plan_schema.return_value = Mock(is_valid=True, errors=[])
            mock_planner_class.return_value = mock_planner
            
            plan = pipeline._generate_plan(mock_structure, user_intent)
            
            assert plan == mock_plan
            mock_planner.generate_plan.assert_called_once_with(mock_structure, user_intent)
            mock_planner.validate_plan_schema.assert_called_once()
    
    def test_generate_plan_validation_failure(self):
        """Test plan generation with validation failure."""
        pipeline = VNextPipeline()
        pipeline.vnext_logger = Mock()
        pipeline.progress_reporter = Mock()
        
        mock_structure = self.create_mock_structure()
        mock_plan = self.create_mock_plan()
        user_intent = "Remove abstract and references sections"
        
        with patch('autoword.vnext.pipeline.DocumentPlanner') as mock_planner_class:
            mock_planner = Mock()
            mock_planner.generate_plan.return_value = mock_plan
            mock_planner.validate_plan_schema.return_value = Mock(
                is_valid=False, 
                errors=["Invalid operation type"]
            )
            mock_planner_class.return_value = mock_planner
            
            with pytest.raises(PlanningError, match="Plan validation failed"):
                pipeline._generate_plan(mock_structure, user_intent)
    
    def test_execute_plan_success(self):
        """Test successful plan execution."""
        with tempfile.TemporaryDirectory() as temp_dir:
            pipeline = VNextPipeline()
            pipeline.current_audit_dir = temp_dir
            pipeline.vnext_logger = Mock()
            pipeline.progress_reporter = Mock()
            
            mock_plan = self.create_mock_plan()
            result_path = os.path.join(temp_dir, "result.docx")
            
            with patch('autoword.vnext.pipeline.DocumentExecutor') as mock_executor_class:
                mock_executor = Mock()
                mock_executor.execute_plan.return_value = result_path
                mock_executor_class.return_value = mock_executor
                
                result = pipeline._execute_plan(mock_plan)
                
                assert result == result_path
                mock_executor.execute_plan.assert_called_once()
    
    def test_execute_plan_failure(self):
        """Test plan execution failure."""
        pipeline = VNextPipeline()
        pipeline.current_audit_dir = "/temp"
        pipeline.vnext_logger = Mock()
        pipeline.progress_reporter = Mock()
        
        mock_plan = self.create_mock_plan()
        
        with patch('autoword.vnext.pipeline.DocumentExecutor') as mock_executor_class:
            mock_executor = Mock()
            mock_executor.execute_plan.side_effect = Exception("Execution failed")
            mock_executor_class.return_value = mock_executor
            
            with pytest.raises(ExecutionError, match="Plan execution failed"):
                pipeline._execute_plan(mock_plan)


class TestVNextPipelineIntegration:
    """Test complete pipeline integration scenarios."""
    
    def create_test_scenario_data(self):
        """Create complete test scenario data."""
        structure = StructureV1(
            metadata=DocumentMetadata(
                title="Test Document",
                author="Test Author",
                paragraph_count=10,
                word_count=100
            ),
            styles=[
                StyleDefinition(
                    name="Heading 1",
                    type=StyleType.PARAGRAPH,
                    font=FontSpec(east_asian="楷体", latin="Times New Roman", size_pt=14, bold=True)
                )
            ],
            paragraphs=[
                ParagraphSkeleton(index=0, style_name="Heading 1", preview_text="摘要", is_heading=True, heading_level=1),
                ParagraphSkeleton(index=1, style_name="Normal", preview_text="Abstract content"),
                ParagraphSkeleton(index=2, style_name="Heading 1", preview_text="Introduction", is_heading=True, heading_level=1)
            ],
            headings=[
                HeadingReference(paragraph_index=0, level=1, text="摘要", style_name="Heading 1"),
                HeadingReference(paragraph_index=2, level=1, text="Introduction", style_name="Heading 1")
            ]
        )
        
        inventory = InventoryFullV1(
            ooxml_fragments={"test_fragment": "<w:p>Test</w:p>"},
            media_indexes={},
            content_controls=[],
            formulas=[],
            charts=[]
        )
        
        plan = PlanV1(
            ops=[
                DeleteSectionByHeading(
                    heading_text="摘要",
                    level=1,
                    match="EXACT",
                    case_sensitive=False
                ),
                UpdateToc()
            ]
        )
        
        return structure, inventory, plan
    
    @patch('autoword.vnext.pipeline.DocumentAuditor')
    @patch('autoword.vnext.pipeline.DocumentValidator')
    @patch('autoword.vnext.pipeline.DocumentExecutor')
    @patch('autoword.vnext.pipeline.DocumentPlanner')
    @patch('autoword.vnext.pipeline.DocumentExtractor')
    @patch('autoword.vnext.pipeline.create_vnext_logger')
    @patch('autoword.vnext.pipeline.PipelineErrorHandler')
    def test_successful_pipeline_execution(self, mock_error_handler, mock_logger, 
                                         mock_extractor_class, mock_planner_class, 
                                         mock_executor_class, mock_validator_class, 
                                         mock_auditor_class):
        """Test complete successful pipeline execution."""
        with tempfile.TemporaryDirectory() as temp_dir:
            # Create test DOCX file
            test_docx = os.path.join(temp_dir, "test.docx")
            with open(test_docx, 'wb') as f:
                f.write(b"Mock DOCX content")
            
            # Setup test data
            structure, inventory, plan = self.create_test_scenario_data()
            modified_docx = os.path.join(temp_dir, "modified.docx")
            audit_dir = os.path.join(temp_dir, "audit_123")
            
            # Mock components
            mock_auditor = Mock()
            mock_auditor.create_audit_directory.return_value = audit_dir
            mock_auditor_class.return_value = mock_auditor
            
            mock_extractor = Mock()
            mock_extractor.__enter__ = Mock(return_value=mock_extractor)
            mock_extractor.__exit__ = Mock(return_value=None)
            mock_extractor.extract_structure.return_value = structure
            mock_extractor.extract_inventory.return_value = inventory
            mock_extractor_class.return_value = mock_extractor
            
            mock_planner = Mock()
            mock_planner.generate_plan.return_value = plan
            mock_planner.validate_plan_schema.return_value = Mock(is_valid=True, errors=[])
            mock_planner_class.return_value = mock_planner
            
            mock_executor = Mock()
            mock_executor.execute_plan.return_value = modified_docx
            mock_executor_class.return_value = mock_executor
            
            mock_validator = Mock()
            mock_validator.__enter__ = Mock(return_value=mock_validator)
            mock_validator.__exit__ = Mock(return_value=None)
            mock_validator.validate_modifications.return_value = Mock(
                is_valid=True, 
                errors=[], 
                warnings=[],
                modified_structure=structure
            )
            mock_validator_class.return_value = mock_validator
            
            mock_vnext_logger = Mock()
            mock_vnext_logger.track_stage.return_value.__enter__ = Mock()
            mock_vnext_logger.track_stage.return_value.__exit__ = Mock()
            mock_vnext_logger.track_operation.return_value.__enter__ = Mock()
            mock_vnext_logger.track_operation.return_value.__exit__ = Mock()
            mock_logger.return_value = mock_vnext_logger
            
            mock_error_handler_instance = Mock()
            mock_error_handler_instance.warnings_logger.get_warnings.return_value = []
            mock_error_handler.return_value = mock_error_handler_instance
            
            # Execute pipeline
            pipeline = VNextPipeline(base_audit_dir=temp_dir)
            result = pipeline.process_document(test_docx, "Remove abstract section")
            
            # Verify result
            assert result.status == "SUCCESS"
            assert result.message == "Document processed successfully"
            assert result.audit_directory == audit_dir
            assert result.errors == []
            
            # Verify all stages were called
            mock_extractor.extract_structure.assert_called_once()
            mock_extractor.extract_inventory.assert_called_once()
            mock_planner.generate_plan.assert_called_once()
            mock_executor.execute_plan.assert_called_once()
            mock_validator.validate_modifications.assert_called_once()
            mock_auditor.save_snapshots.assert_called_once()
            mock_auditor.write_status.assert_called_once_with("SUCCESS", "Pipeline completed successfully")
    
    @patch('autoword.vnext.pipeline.DocumentAuditor')
    @patch('autoword.vnext.pipeline.DocumentValidator')
    @patch('autoword.vnext.pipeline.DocumentExecutor')
    @patch('autoword.vnext.pipeline.DocumentPlanner')
    @patch('autoword.vnext.pipeline.DocumentExtractor')
    @patch('autoword.vnext.pipeline.create_vnext_logger')
    @patch('autoword.vnext.pipeline.PipelineErrorHandler')
    def test_validation_failure_with_rollback(self, mock_error_handler, mock_logger,
                                            mock_extractor_class, mock_planner_class,
                                            mock_executor_class, mock_validator_class,
                                            mock_auditor_class):
        """Test pipeline with validation failure and rollback."""
        with tempfile.TemporaryDirectory() as temp_dir:
            # Create test DOCX file
            test_docx = os.path.join(temp_dir, "test.docx")
            with open(test_docx, 'wb') as f:
                f.write(b"Mock DOCX content")
            
            # Setup test data
            structure, inventory, plan = self.create_test_scenario_data()
            modified_docx = os.path.join(temp_dir, "modified.docx")
            audit_dir = os.path.join(temp_dir, "audit_123")
            
            # Mock components with validation failure
            mock_auditor = Mock()
            mock_auditor.create_audit_directory.return_value = audit_dir
            mock_auditor_class.return_value = mock_auditor
            
            mock_extractor = Mock()
            mock_extractor.__enter__ = Mock(return_value=mock_extractor)
            mock_extractor.__exit__ = Mock(return_value=None)
            mock_extractor.extract_structure.return_value = structure
            mock_extractor.extract_inventory.return_value = inventory
            mock_extractor_class.return_value = mock_extractor
            
            mock_planner = Mock()
            mock_planner.generate_plan.return_value = plan
            mock_planner.validate_plan_schema.return_value = Mock(is_valid=True, errors=[])
            mock_planner_class.return_value = mock_planner
            
            mock_executor = Mock()
            mock_executor.execute_plan.return_value = modified_docx
            mock_executor_class.return_value = mock_executor
            
            mock_validator = Mock()
            mock_validator.__enter__ = Mock(return_value=mock_validator)
            mock_validator.__exit__ = Mock(return_value=None)
            mock_validator.validate_modifications.return_value = Mock(
                is_valid=False,
                errors=["Chapter assertion failed: 摘要 found at level 1"],
                warnings=[],
                modified_structure=structure
            )
            mock_validator_class.return_value = mock_validator
            
            mock_vnext_logger = Mock()
            mock_vnext_logger.track_stage.return_value.__enter__ = Mock()
            mock_vnext_logger.track_stage.return_value.__exit__ = Mock()
            mock_vnext_logger.track_operation.return_value.__enter__ = Mock()
            mock_vnext_logger.track_operation.return_value.__exit__ = Mock()
            mock_logger.return_value = mock_vnext_logger
            
            mock_error_handler_instance = Mock()
            mock_error_handler_instance.rollback_manager.perform_rollback.return_value = Mock(success=True)
            mock_error_handler.return_value = mock_error_handler_instance
            
            # Execute pipeline
            pipeline = VNextPipeline(base_audit_dir=temp_dir)
            result = pipeline.process_document(test_docx, "Remove abstract section")
            
            # Verify result
            assert result.status == "FAILED_VALIDATION"
            assert "Document validation failed" in result.message
            assert "Chapter assertion failed: 摘要 found at level 1" in result.errors
            assert result.audit_directory == audit_dir
            
            # Verify rollback was performed
            mock_error_handler_instance.rollback_manager.perform_rollback.assert_called_once()
            mock_auditor.write_status.assert_called_once_with(
                "FAILED_VALIDATION", 
                "Validation failed: Chapter assertion failed: 摘要 found at level 1"
            )
    
    def test_extraction_error_handling(self):
        """Test pipeline error handling during extraction stage."""
        with tempfile.TemporaryDirectory() as temp_dir:
            # Create test DOCX file
            test_docx = os.path.join(temp_dir, "test.docx")
            with open(test_docx, 'wb') as f:
                f.write(b"Mock DOCX content")
            
            with patch('autoword.vnext.pipeline.DocumentAuditor') as mock_auditor_class:
                mock_auditor = Mock()
                mock_auditor.create_audit_directory.return_value = os.path.join(temp_dir, "audit_123")
                mock_auditor_class.return_value = mock_auditor
                
                with patch('autoword.vnext.pipeline.create_vnext_logger') as mock_logger:
                    with patch('autoword.vnext.pipeline.PipelineErrorHandler') as mock_error_handler:
                        mock_error_handler_instance = Mock()
                        mock_error_handler_instance.handle_pipeline_error.return_value = ProcessingResult(
                            status="ROLLBACK",
                            message="Extraction failed",
                            errors=["Document extraction failed: Test error"]
                        )
                        mock_error_handler.return_value = mock_error_handler_instance
                        
                        with patch('autoword.vnext.pipeline.DocumentExtractor') as mock_extractor_class:
                            mock_extractor = Mock()
                            mock_extractor.__enter__ = Mock(return_value=mock_extractor)
                            mock_extractor.__exit__ = Mock(return_value=None)
                            mock_extractor.extract_structure.side_effect = Exception("Test error")
                            mock_extractor_class.return_value = mock_extractor
                            
                            pipeline = VNextPipeline(base_audit_dir=temp_dir)
                            result = pipeline.process_document(test_docx, "Test intent")
                            
                            assert result.status == "ROLLBACK"
                            assert "Extraction failed" in result.message
                            assert any("Document extraction failed" in error for error in result.errors)


class TestVNextPipelineWithRealScenarios:
    """Test pipeline with realistic scenarios using test data."""
    
    def test_scenario_1_normal_paper_processing(self):
        """Test normal paper processing scenario."""
        scenario_dir = Path("autoword/vnext/test_data/scenario_1_normal_paper")
        
        if not scenario_dir.exists():
            pytest.skip("Test data not available")
        
        # Read user intent
        user_intent_file = scenario_dir / "user_intent.txt"
        if user_intent_file.exists():
            with open(user_intent_file, 'r', encoding='utf-8') as f:
                user_intent = f.read().strip()
        else:
            user_intent = "Remove abstract and references sections, update TOC"
        
        # Read expected plan
        expected_plan_file = scenario_dir / "expected_plan.v1.json"
        if expected_plan_file.exists():
            with open(expected_plan_file, 'r', encoding='utf-8') as f:
                expected_plan_data = json.load(f)
        else:
            pytest.skip("Expected plan not available")
        
        # Mock the pipeline components to return expected data
        with patch('autoword.vnext.pipeline.DocumentExtractor') as mock_extractor_class:
            with patch('autoword.vnext.pipeline.DocumentPlanner') as mock_planner_class:
                with patch('autoword.vnext.pipeline.DocumentExecutor') as mock_executor_class:
                    with patch('autoword.vnext.pipeline.DocumentValidator') as mock_validator_class:
                        with patch('autoword.vnext.pipeline.DocumentAuditor') as mock_auditor_class:
                            # Setup mocks to simulate successful processing
                            self._setup_successful_mocks(
                                mock_extractor_class, mock_planner_class, mock_executor_class,
                                mock_validator_class, mock_auditor_class, expected_plan_data
                            )
                            
                            # Create temporary test file
                            with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as temp_file:
                                temp_file.write(b"Mock DOCX content")
                                temp_docx = temp_file.name
                            
                            try:
                                pipeline = VNextPipeline()
                                result = pipeline.process_document(temp_docx, user_intent)
                                
                                assert result.status == "SUCCESS"
                                assert result.message == "Document processed successfully"
                                
                            finally:
                                os.unlink(temp_docx)
    
    def _setup_successful_mocks(self, mock_extractor_class, mock_planner_class, 
                              mock_executor_class, mock_validator_class, 
                              mock_auditor_class, expected_plan_data):
        """Setup mocks for successful pipeline execution."""
        # Mock structure
        mock_structure = StructureV1(
            metadata=DocumentMetadata(title="Test Document", paragraph_count=10),
            paragraphs=[ParagraphSkeleton(index=0, preview_text="Test content")]
        )
        
        # Mock inventory
        mock_inventory = InventoryFullV1()
        
        # Mock plan from expected data
        mock_plan = PlanV1.model_validate(expected_plan_data)
        
        # Setup extractor
        mock_extractor = Mock()
        mock_extractor.__enter__ = Mock(return_value=mock_extractor)
        mock_extractor.__exit__ = Mock(return_value=None)
        mock_extractor.extract_structure.return_value = mock_structure
        mock_extractor.extract_inventory.return_value = mock_inventory
        mock_extractor_class.return_value = mock_extractor
        
        # Setup planner
        mock_planner = Mock()
        mock_planner.generate_plan.return_value = mock_plan
        mock_planner.validate_plan_schema.return_value = Mock(is_valid=True, errors=[])
        mock_planner_class.return_value = mock_planner
        
        # Setup executor
        mock_executor = Mock()
        mock_executor.execute_plan.return_value = "/temp/modified.docx"
        mock_executor_class.return_value = mock_executor
        
        # Setup validator
        mock_validator = Mock()
        mock_validator.__enter__ = Mock(return_value=mock_validator)
        mock_validator.__exit__ = Mock(return_value=None)
        mock_validator.validate_modifications.return_value = Mock(
            is_valid=True, errors=[], warnings=[], modified_structure=mock_structure
        )
        mock_validator_class.return_value = mock_validator
        
        # Setup auditor
        mock_auditor = Mock()
        mock_auditor.create_audit_directory.return_value = "/temp/audit_123"
        mock_auditor_class.return_value = mock_auditor
        
        # Setup logger and error handler
        with patch('autoword.vnext.pipeline.create_vnext_logger') as mock_logger:
            mock_vnext_logger = Mock()
            mock_vnext_logger.track_stage.return_value.__enter__ = Mock()
            mock_vnext_logger.track_stage.return_value.__exit__ = Mock()
            mock_vnext_logger.track_operation.return_value.__enter__ = Mock()
            mock_vnext_logger.track_operation.return_value.__exit__ = Mock()
            mock_logger.return_value = mock_vnext_logger
            
            with patch('autoword.vnext.pipeline.PipelineErrorHandler') as mock_error_handler:
                mock_error_handler_instance = Mock()
                mock_error_handler_instance.warnings_logger.get_warnings.return_value = []
                mock_error_handler.return_value = mock_error_handler_instance


if __name__ == "__main__":
    pytest.main([__file__, "-v"])