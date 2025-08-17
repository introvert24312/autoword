"""
Tests for DocumentAuditor class.

This module tests the complete audit trail generation functionality
including timestamped directories, snapshots, diff reports, and status logging.
"""

import os
import json
import tempfile
import shutil
from datetime import datetime
from pathlib import Path
from unittest.mock import Mock, patch, mock_open

import pytest

from autoword.vnext.auditor.document_auditor import DocumentAuditor
from autoword.vnext.models import (
    StructureV1, PlanV1, DiffReport, DocumentMetadata, 
    ParagraphSkeleton, HeadingReference, StyleDefinition,
    FontSpec, ParagraphSpec, DeleteSectionByHeading, UpdateToc
)
from autoword.vnext.exceptions import AuditError


class TestDocumentAuditor:
    """Test cases for DocumentAuditor class."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.temp_dir = tempfile.mkdtemp()
        self.auditor = DocumentAuditor(base_audit_dir=self.temp_dir)
        
        # Create sample structures
        self.before_structure = StructureV1(
            metadata=DocumentMetadata(
                title="Test Document",
                author="Test Author",
                creation_time=datetime.now(),
                modified_time=datetime.now(),
                page_count=5,
                paragraph_count=20,
                word_count=500
            ),
            paragraphs=[
                ParagraphSkeleton(index=0, style_name="Normal", preview_text="First paragraph", is_heading=False),
                ParagraphSkeleton(index=1, style_name="Heading 1", preview_text="Introduction", is_heading=True, heading_level=1),
                ParagraphSkeleton(index=2, style_name="Normal", preview_text="Content paragraph", is_heading=False),
                ParagraphSkeleton(index=3, style_name="Heading 1", preview_text="Abstract", is_heading=True, heading_level=1),
                ParagraphSkeleton(index=4, style_name="Normal", preview_text="Abstract content", is_heading=False)
            ],
            headings=[
                HeadingReference(paragraph_index=1, level=1, text="Introduction", style_name="Heading 1"),
                HeadingReference(paragraph_index=3, level=1, text="Abstract", style_name="Heading 1")
            ],
            styles=[
                StyleDefinition(
                    name="Normal",
                    type="paragraph",
                    font=FontSpec(latin="Times New Roman", size_pt=12),
                    paragraph=ParagraphSpec(line_spacing_mode="MULTIPLE", line_spacing_value=2.0)
                ),
                StyleDefinition(
                    name="Heading 1",
                    type="paragraph",
                    font=FontSpec(latin="Times New Roman", size_pt=14, bold=True),
                    paragraph=ParagraphSpec(line_spacing_mode="MULTIPLE", line_spacing_value=2.0)
                )
            ]
        )
        
        self.after_structure = StructureV1(
            metadata=DocumentMetadata(
                title="Test Document",
                author="Test Author",
                creation_time=datetime.now(),
                modified_time=datetime.now(),
                page_count=4,
                paragraph_count=15,
                word_count=400
            ),
            paragraphs=[
                ParagraphSkeleton(index=0, style_name="Normal", preview_text="First paragraph", is_heading=False),
                ParagraphSkeleton(index=1, style_name="Heading 1", preview_text="Introduction", is_heading=True, heading_level=1),
                ParagraphSkeleton(index=2, style_name="Normal", preview_text="Content paragraph", is_heading=False)
                # Note: Abstract section removed (paragraphs 3 and 4)
            ],
            headings=[
                HeadingReference(paragraph_index=1, level=1, text="Introduction", style_name="Heading 1")
                # Note: Abstract heading removed
            ],
            styles=[
                StyleDefinition(
                    name="Normal",
                    type="paragraph",
                    font=FontSpec(latin="Times New Roman", size_pt=12),
                    paragraph=ParagraphSpec(line_spacing_mode="MULTIPLE", line_spacing_value=2.0)
                ),
                StyleDefinition(
                    name="Heading 1",
                    type="paragraph",
                    font=FontSpec(latin="Arial", size_pt=14, bold=True),  # Font changed
                    paragraph=ParagraphSpec(line_spacing_mode="MULTIPLE", line_spacing_value=2.0)
                )
            ]
        )
        
        self.sample_plan = PlanV1(
            ops=[
                DeleteSectionByHeading(
                    heading_text="Abstract",
                    level=1,
                    match="EXACT",
                    case_sensitive=False
                ),
                UpdateToc()
            ]
        )
    
    def teardown_method(self):
        """Clean up test fixtures."""
        if os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir)
    
    def test_create_audit_directory(self):
        """Test audit directory creation with timestamp."""
        audit_dir = self.auditor.create_audit_directory()
        
        # Check directory was created
        assert os.path.exists(audit_dir)
        assert "run_" in audit_dir
        
        # Check subdirectories were created
        assert os.path.exists(os.path.join(audit_dir, "snapshots"))
        assert os.path.exists(os.path.join(audit_dir, "structures"))
        assert os.path.exists(os.path.join(audit_dir, "reports"))
        
        # Check auditor state
        assert self.auditor.current_audit_dir is not None
        assert str(self.auditor.current_audit_dir) == audit_dir
    
    def test_create_audit_directory_failure(self):
        """Test audit directory creation failure handling."""
        # Use invalid path to trigger failure - use Windows-style invalid path
        invalid_path = "Z:\\invalid\\path\\that\\does\\not\\exist" if os.name == 'nt' else "/invalid/path/that/does/not/exist"
        auditor = DocumentAuditor(base_audit_dir=invalid_path)
        
        with pytest.raises(AuditError) as exc_info:
            auditor.create_audit_directory()
        
        assert "Failed to create audit directory" in str(exc_info.value)
        assert exc_info.value.details["audit_stage"] == "directory_creation"
    
    def test_save_snapshots(self):
        """Test saving DOCX and JSON snapshots."""
        # Create audit directory first
        self.auditor.create_audit_directory()
        
        # Create temporary DOCX files
        before_docx = os.path.join(self.temp_dir, "before.docx")
        after_docx = os.path.join(self.temp_dir, "after.docx")
        
        with open(before_docx, 'wb') as f:
            f.write(b"Mock DOCX content before")
        with open(after_docx, 'wb') as f:
            f.write(b"Mock DOCX content after")
        
        # Save snapshots
        self.auditor.save_snapshots(
            before_docx=before_docx,
            after_docx=after_docx,
            before_structure=self.before_structure,
            after_structure=self.after_structure,
            plan=self.sample_plan
        )
        
        audit_dir = self.auditor.current_audit_dir
        
        # Check DOCX files were copied
        assert os.path.exists(audit_dir / "snapshots" / "before.docx")
        assert os.path.exists(audit_dir / "snapshots" / "after.docx")
        
        # Check structure JSON files were created
        assert os.path.exists(audit_dir / "structures" / "structure.before.v1.json")
        assert os.path.exists(audit_dir / "structures" / "structure.after.v1.json")
        
        # Check plan JSON file was created
        assert os.path.exists(audit_dir / "plan.v1.json")
        
        # Verify JSON content
        with open(audit_dir / "structures" / "structure.before.v1.json", 'r', encoding='utf-8') as f:
            before_data = json.load(f)
            assert before_data["schema_version"] == "structure.v1"
            assert len(before_data["paragraphs"]) == 5
        
        with open(audit_dir / "plan.v1.json", 'r', encoding='utf-8') as f:
            plan_data = json.load(f)
            assert plan_data["schema_version"] == "plan.v1"
            assert len(plan_data["ops"]) == 2
    
    def test_save_snapshots_no_audit_directory(self):
        """Test save_snapshots fails without audit directory."""
        with pytest.raises(AuditError) as exc_info:
            self.auditor.save_snapshots(
                before_docx="dummy.docx",
                after_docx="dummy.docx",
                before_structure=self.before_structure,
                after_structure=self.after_structure,
                plan=self.sample_plan
            )
        
        assert "No audit directory created" in str(exc_info.value)
        assert exc_info.value.details["audit_stage"] == "snapshot_saving"
    
    def test_generate_diff_report(self):
        """Test diff report generation."""
        self.auditor.create_audit_directory()
        
        diff_report = self.auditor.generate_diff_report(
            self.before_structure, 
            self.after_structure
        )
        
        # Check diff report structure
        assert isinstance(diff_report, DiffReport)
        
        # Check removed paragraphs (Abstract section)
        assert 3 in diff_report.removed_paragraphs
        assert 4 in diff_report.removed_paragraphs
        
        # Check heading changes (Abstract heading removed)
        assert len(diff_report.heading_changes) == 1
        assert diff_report.heading_changes[0]["type"] == "removed"
        assert diff_report.heading_changes[0]["text"] == "Abstract"
        
        # Check style changes (Heading 1 font changed)
        assert "Heading 1" in diff_report.style_changes
        font_change = diff_report.style_changes["Heading 1"]["font"]
        assert font_change["before"]["latin"] == "Times New Roman"
        assert font_change["after"]["latin"] == "Arial"
        
        # Check summary
        assert "paragraphs removed" in diff_report.summary
        assert "heading changes" in diff_report.summary
        assert "styles changed" in diff_report.summary
        
        # Check diff report was saved
        diff_file = self.auditor.current_audit_dir / "reports" / "diff.report.json"
        assert os.path.exists(diff_file)
    
    def test_generate_diff_report_no_changes(self):
        """Test diff report with no changes."""
        self.auditor.create_audit_directory()
        
        diff_report = self.auditor.generate_diff_report(
            self.before_structure, 
            self.before_structure  # Same structure
        )
        
        assert len(diff_report.added_paragraphs) == 0
        assert len(diff_report.removed_paragraphs) == 0
        assert len(diff_report.modified_paragraphs) == 0
        assert len(diff_report.style_changes) == 0
        assert len(diff_report.heading_changes) == 0
        assert diff_report.summary == "No structural changes detected"
    
    def test_add_warning(self):
        """Test warning accumulation."""
        self.auditor.add_warning("Font fallback: 楷体 -> 楷体_GB2312")
        self.auditor.add_warning("NOOP operation: update_toc (no TOC found)")
        
        assert len(self.auditor.warnings) == 2
        assert "Font fallback" in self.auditor.warnings[0]
        assert "NOOP operation" in self.auditor.warnings[1]
        
        # Check timestamps are included
        for warning in self.auditor.warnings:
            assert ":" in warning  # ISO format timestamp contains colons
    
    def test_write_warnings_log(self):
        """Test warnings log file creation."""
        self.auditor.create_audit_directory()
        
        # Add some warnings
        self.auditor.add_warning("Test warning 1")
        self.auditor.add_warning("Test warning 2")
        
        self.auditor.write_warnings_log()
        
        warnings_file = self.auditor.current_audit_dir / "warnings.log"
        assert os.path.exists(warnings_file)
        
        with open(warnings_file, 'r', encoding='utf-8') as f:
            content = f.read()
            assert "Test warning 1" in content
            assert "Test warning 2" in content
    
    def test_write_warnings_log_empty(self):
        """Test warnings log with no warnings."""
        self.auditor.create_audit_directory()
        
        self.auditor.write_warnings_log()
        
        warnings_file = self.auditor.current_audit_dir / "warnings.log"
        assert os.path.exists(warnings_file)
        
        with open(warnings_file, 'r', encoding='utf-8') as f:
            content = f.read()
            assert "No warnings generated" in content
    
    def test_write_status_success(self):
        """Test status file writing for SUCCESS."""
        self.auditor.create_audit_directory()
        
        self.auditor.write_status("SUCCESS", "Document processed successfully")
        
        status_file = self.auditor.current_audit_dir / "result.status.txt"
        assert os.path.exists(status_file)
        
        with open(status_file, 'r', encoding='utf-8') as f:
            content = f.read()
            assert "STATUS: SUCCESS" in content
            assert "DETAILS: Document processed successfully" in content
            assert "TIMESTAMP:" in content
            assert "AUDIT_DIRECTORY:" in content
            assert "WARNING_COUNT: 0" in content
    
    def test_write_status_failed_validation(self):
        """Test status file writing for FAILED_VALIDATION."""
        self.auditor.create_audit_directory()
        self.auditor.add_warning("Test warning")
        
        self.auditor.write_status("FAILED_VALIDATION", "Chapter assertion failed")
        
        status_file = self.auditor.current_audit_dir / "result.status.txt"
        assert os.path.exists(status_file)
        
        with open(status_file, 'r', encoding='utf-8') as f:
            content = f.read()
            assert "STATUS: FAILED_VALIDATION" in content
            assert "DETAILS: Chapter assertion failed" in content
            assert "WARNING_COUNT: 1" in content
    
    def test_write_status_invalid_status(self):
        """Test status writing with invalid status."""
        self.auditor.create_audit_directory()
        
        with pytest.raises(AuditError) as exc_info:
            self.auditor.write_status("INVALID_STATUS", "Test details")
        
        assert "Invalid status 'INVALID_STATUS'" in str(exc_info.value)
        assert exc_info.value.details["audit_stage"] == "status_writing"
    
    def test_write_status_no_audit_directory(self):
        """Test status writing without audit directory."""
        with pytest.raises(AuditError) as exc_info:
            self.auditor.write_status("SUCCESS", "Test details")
        
        assert "No audit directory created" in str(exc_info.value)
        assert exc_info.value.details["audit_stage"] == "status_writing"
    
    def test_finalize_audit(self):
        """Test complete audit finalization."""
        self.auditor.create_audit_directory()
        self.auditor.add_warning("Test warning")
        
        self.auditor.finalize_audit("SUCCESS", "Processing completed successfully")
        
        # Check both files were created
        warnings_file = self.auditor.current_audit_dir / "warnings.log"
        status_file = self.auditor.current_audit_dir / "result.status.txt"
        
        assert os.path.exists(warnings_file)
        assert os.path.exists(status_file)
        
        # Verify content
        with open(status_file, 'r', encoding='utf-8') as f:
            content = f.read()
            assert "STATUS: SUCCESS" in content
            assert "WARNING_COUNT: 1" in content
    
    def test_get_audit_directory(self):
        """Test audit directory path retrieval."""
        # Initially no directory
        assert self.auditor.get_audit_directory() is None
        
        # After creation
        audit_dir = self.auditor.create_audit_directory()
        assert self.auditor.get_audit_directory() == audit_dir
    
    def test_analyze_style_changes_detailed(self):
        """Test detailed style change analysis."""
        # Create styles with different changes
        before_styles = [
            StyleDefinition(
                name="Normal",
                type="paragraph",
                font=FontSpec(latin="Times New Roman", size_pt=12),
                paragraph=ParagraphSpec(line_spacing_mode="SINGLE", line_spacing_value=1.0)
            ),
            StyleDefinition(
                name="Heading 1",
                type="paragraph",
                font=FontSpec(latin="Arial", size_pt=14, bold=True)
            ),
            StyleDefinition(
                name="ToBeRemoved",
                type="paragraph",
                font=FontSpec(latin="Calibri", size_pt=11)
            )
        ]
        
        after_styles = [
            StyleDefinition(
                name="Normal",
                type="paragraph",
                font=FontSpec(latin="Calibri", size_pt=11),  # Changed font and size
                paragraph=ParagraphSpec(line_spacing_mode="MULTIPLE", line_spacing_value=1.5)  # Changed spacing
            ),
            StyleDefinition(
                name="Heading 1",
                type="paragraph",
                font=FontSpec(latin="Arial", size_pt=14, bold=True)  # No change
            ),
            StyleDefinition(
                name="NewStyle",
                type="paragraph",
                font=FontSpec(latin="Georgia", size_pt=10)
            )
        ]
        
        changes = self.auditor._analyze_style_changes(before_styles, after_styles)
        
        # Check Normal style changes
        assert "Normal" in changes
        normal_changes = changes["Normal"]
        assert "font" in normal_changes
        assert normal_changes["font"]["before"]["latin"] == "Times New Roman"
        assert normal_changes["font"]["after"]["latin"] == "Calibri"
        assert "paragraph" in normal_changes
        
        # Check added style
        assert "NewStyle" in changes
        assert "added" in changes["NewStyle"]
        
        # Check removed style
        assert "ToBeRemoved" in changes
        assert "removed" in changes["ToBeRemoved"]
        
        # Check unchanged style is not in changes
        assert "Heading 1" not in changes
    
    def test_analyze_heading_changes_detailed(self):
        """Test detailed heading change analysis."""
        before_headings = [
            HeadingReference(paragraph_index=1, level=1, text="Introduction", style_name="Heading 1"),
            HeadingReference(paragraph_index=3, level=1, text="Abstract", style_name="Heading 1"),
            HeadingReference(paragraph_index=5, level=2, text="Methods", style_name="Heading 2")
        ]
        
        after_headings = [
            HeadingReference(paragraph_index=1, level=1, text="Introduction", style_name="Heading 1"),
            HeadingReference(paragraph_index=3, level=1, text="Conclusion", style_name="Heading 1"),  # Changed text
            HeadingReference(paragraph_index=7, level=2, text="Results", style_name="Heading 2")  # New heading
        ]
        
        changes = self.auditor._analyze_heading_changes(before_headings, after_headings)
        
        # Should have 3 changes: Abstract removed, Methods removed, Results added, Conclusion added
        assert len(changes) == 4
        
        # Check for specific changes
        removed_changes = [c for c in changes if c["type"] == "removed"]
        added_changes = [c for c in changes if c["type"] == "added"]
        
        assert len(removed_changes) == 2
        assert len(added_changes) == 2
        
        # Check specific removed headings
        removed_texts = [c["text"] for c in removed_changes]
        assert "Abstract" in removed_texts
        assert "Methods" in removed_texts
        
        # Check specific added headings
        added_texts = [c["text"] for c in added_changes]
        assert "Conclusion" in added_texts
        assert "Results" in added_texts


if __name__ == "__main__":
    pytest.main([__file__])