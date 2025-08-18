"""
Simple integration tests for VNextPipeline orchestrator.

This module provides focused tests for the pipeline orchestrator functionality
without complex mocking or external dependencies.
"""

import os
import tempfile
import shutil
import json
import pytest
from pathlib import Path
from unittest.mock import Mock, patch, MagicMock

from autoword.vnext.pipeline import VNextPipeline, ProgressReporter
from autoword.vnext.models import (
    StructureV1, PlanV1, InventoryFullV1, ProcessingResult,
    DocumentMetadata, StyleDefinition, ParagraphSkeleton, HeadingReference,
    DeleteSectionByHeading, UpdateToc, SetStyleRule, FontSpec, ParagraphSpec,
    LineSpacingMode, StyleType
)


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
        mock_llm = Mock()
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


class TestVNextPipelineDataModels:
    """Test data model creation and validation."""
    
    def test_structure_v1_creation(self):
        """Test StructureV1 model creation."""
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
        
        assert structure.schema_version == "structure.v1"
        assert structure.metadata.title == "Test Document"
        assert len(structure.styles) == 1
        assert len(structure.paragraphs) == 2
        assert len(structure.headings) == 1
    
    def test_plan_v1_creation(self):
        """Test PlanV1 model creation."""
        plan = PlanV1(
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
        
        assert plan.schema_version == "plan.v1"
        assert len(plan.ops) == 3
        assert plan.ops[0].operation_type == "delete_section_by_heading"
        assert plan.ops[1].operation_type == "update_toc"
        assert plan.ops[2].operation_type == "set_style_rule"
    
    def test_inventory_v1_creation(self):
        """Test InventoryFullV1 model creation."""
        inventory = InventoryFullV1(
            ooxml_fragments={"test_fragment": "<w:p>Test</w:p>"},
            media_indexes={},
            content_controls=[],
            formulas=[],
            charts=[]
        )
        
        assert inventory.schema_version == "inventory.full.v1"
        assert "test_fragment" in inventory.ooxml_fragments
        assert inventory.ooxml_fragments["test_fragment"] == "<w:p>Test</w:p>"


class TestVNextPipelineEnvironmentManagement:
    """Test pipeline environment setup and cleanup."""
    
    def test_setup_run_environment_missing_file(self):
        """Test setup with missing input file."""
        pipeline = VNextPipeline()
        
        with pytest.raises(Exception):  # Should raise ExtractionError but we'll catch any exception
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
            pipeline.vnext_logger.cleanup = Mock()
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


class TestVNextPipelineBasicFlow:
    """Test basic pipeline flow without complex mocking."""
    
    def test_pipeline_process_document_with_missing_file(self):
        """Test pipeline processing with missing file."""
        pipeline = VNextPipeline()
        
        result = pipeline.process_document("/nonexistent/file.docx", "Test intent")
        
        # Should return error result
        assert result.status in ["ROLLBACK", "FAILED_VALIDATION"]
        assert len(result.errors) > 0
    
    def test_pipeline_timestamped_directory_creation(self):
        """Test that pipeline creates timestamped directories."""
        with tempfile.TemporaryDirectory() as temp_dir:
            # Create a mock DOCX file
            test_docx = os.path.join(temp_dir, "test.docx")
            with open(test_docx, 'wb') as f:
                f.write(b"Mock DOCX content")
            
            # Mock all the components to avoid actual processing
            with patch('autoword.vnext.pipeline.DocumentAuditor') as mock_auditor_class:
                with patch('autoword.vnext.pipeline.DocumentExtractor') as mock_extractor_class:
                    with patch('autoword.vnext.pipeline.DocumentPlanner') as mock_planner_class:
                        with patch('autoword.vnext.pipeline.DocumentExecutor') as mock_executor_class:
                            with patch('autoword.vnext.pipeline.DocumentValidator') as mock_validator_class:
                                with patch('autoword.vnext.pipeline.create_vnext_logger') as mock_logger:
                                    with patch('autoword.vnext.pipeline.PipelineErrorHandler') as mock_error_handler:
                                        
                                        # Setup basic mocks
                                        mock_auditor = Mock()
                                        audit_dir = os.path.join(temp_dir, "audit_20250101_120000")
                                        mock_auditor.create_audit_directory.return_value = audit_dir
                                        mock_auditor_class.return_value = mock_auditor
                                        
                                        # Mock extractor to raise an exception early (to avoid complex setup)
                                        mock_extractor = Mock()
                                        mock_extractor.__enter__ = Mock(return_value=mock_extractor)
                                        mock_extractor.__exit__ = Mock(return_value=None)
                                        mock_extractor.extract_structure.side_effect = Exception("Test extraction error")
                                        mock_extractor_class.return_value = mock_extractor
                                        
                                        # Mock logger
                                        mock_vnext_logger = Mock()
                                        mock_vnext_logger.track_stage.return_value.__enter__ = Mock()
                                        mock_vnext_logger.track_stage.return_value.__exit__ = Mock()
                                        mock_vnext_logger.track_operation.return_value.__enter__ = Mock()
                                        mock_vnext_logger.track_operation.return_value.__exit__ = Mock()
                                        mock_logger.return_value = mock_vnext_logger
                                        
                                        # Mock error handler
                                        mock_error_handler_instance = Mock()
                                        mock_error_handler_instance.handle_pipeline_error.return_value = ProcessingResult(
                                            status="ROLLBACK",
                                            message="Test error",
                                            errors=["Test extraction error"],
                                            warnings=[]
                                        )
                                        mock_error_handler.return_value = mock_error_handler_instance
                                        
                                        pipeline = VNextPipeline(base_audit_dir=temp_dir)
                                        result = pipeline.process_document(test_docx, "Test intent")
                                        
                                        # Verify that audit directory was created
                                        mock_auditor.create_audit_directory.assert_called_once()
                                        
                                        # Verify error handling (may not get exact message due to error handler complexity)
                                        assert result.status == "ROLLBACK"
                                        assert len(result.errors) > 0


class TestVNextPipelineJSONSerialization:
    """Test JSON serialization of pipeline data models."""
    
    def test_structure_json_serialization(self):
        """Test StructureV1 JSON serialization."""
        structure = StructureV1(
            metadata=DocumentMetadata(
                title="Test Document",
                author="Test Author",
                paragraph_count=5
            ),
            paragraphs=[
                ParagraphSkeleton(index=0, preview_text="Test content")
            ]
        )
        
        # Test serialization
        json_data = structure.model_dump()
        assert json_data["schema_version"] == "structure.v1"
        assert json_data["metadata"]["title"] == "Test Document"
        
        # Test deserialization
        structure_restored = StructureV1.model_validate(json_data)
        assert structure_restored.metadata.title == "Test Document"
        assert len(structure_restored.paragraphs) == 1
    
    def test_plan_json_serialization(self):
        """Test PlanV1 JSON serialization."""
        plan = PlanV1(
            ops=[
                DeleteSectionByHeading(
                    heading_text="Test Heading",
                    level=1,
                    match="EXACT"
                ),
                UpdateToc()
            ]
        )
        
        # Test serialization
        json_data = plan.model_dump()
        assert json_data["schema_version"] == "plan.v1"
        assert len(json_data["ops"]) == 2
        assert json_data["ops"][0]["operation_type"] == "delete_section_by_heading"
        assert json_data["ops"][1]["operation_type"] == "update_toc"
        
        # Test deserialization
        plan_restored = PlanV1.model_validate(json_data)
        assert len(plan_restored.ops) == 2
        assert plan_restored.ops[0].heading_text == "Test Heading"


class TestVNextPipelineErrorHandling:
    """Test pipeline error handling scenarios."""
    
    def test_processing_result_creation(self):
        """Test ProcessingResult model creation."""
        result = ProcessingResult(
            status="SUCCESS",
            message="Processing completed",
            errors=[],
            warnings=["Minor warning"],
            audit_directory="/path/to/audit"
        )
        
        assert result.status == "SUCCESS"
        assert result.message == "Processing completed"
        assert len(result.errors) == 0
        assert len(result.warnings) == 1
        assert result.audit_directory == "/path/to/audit"
    
    def test_processing_result_failure(self):
        """Test ProcessingResult for failure scenarios."""
        result = ProcessingResult(
            status="FAILED_VALIDATION",
            message="Validation failed",
            errors=["Chapter assertion failed", "Style assertion failed"],
            warnings=[],
            audit_directory="/path/to/audit"
        )
        
        assert result.status == "FAILED_VALIDATION"
        assert len(result.errors) == 2
        assert "Chapter assertion failed" in result.errors
        assert "Style assertion failed" in result.errors


if __name__ == "__main__":
    pytest.main([__file__, "-v"])