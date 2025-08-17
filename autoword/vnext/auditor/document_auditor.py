"""
Document auditor for complete audit trails and snapshots.

This module creates timestamped audit directories with complete
before/after snapshots and diff reports for full traceability.
"""

import os
import json
import shutil
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Any, Optional

from ..models import StructureV1, PlanV1, DiffReport
from ..exceptions import AuditError


class DocumentAuditor:
    """Create complete audit trails with timestamped snapshots."""
    
    def __init__(self, base_audit_dir: Optional[str] = None):
        """
        Initialize document auditor.
        
        Args:
            base_audit_dir: Base directory for audit trails (defaults to ./audit_trails)
        """
        self.base_audit_dir = Path(base_audit_dir or "./audit_trails")
        self.current_audit_dir: Optional[Path] = None
        self.warnings: List[str] = []
    
    def create_audit_directory(self) -> str:
        """
        Create timestamped run directory.
        
        Returns:
            str: Path to created audit directory
            
        Raises:
            AuditError: If directory creation fails
        """
        try:
            # Create timestamped directory name
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:-3]  # Include milliseconds
            audit_dir_name = f"run_{timestamp}"
            
            # Create full audit directory path
            self.current_audit_dir = self.base_audit_dir / audit_dir_name
            
            # Create directory structure
            self.current_audit_dir.mkdir(parents=True, exist_ok=False)
            
            # Create subdirectories for organized storage
            (self.current_audit_dir / "snapshots").mkdir()
            (self.current_audit_dir / "structures").mkdir()
            (self.current_audit_dir / "reports").mkdir()
            
            return str(self.current_audit_dir)
            
        except Exception as e:
            raise AuditError(
                f"Failed to create audit directory: {str(e)}",
                audit_stage="directory_creation"
            )
    
    def save_snapshots(self, before_docx: str, after_docx: str, 
                      before_structure: StructureV1, after_structure: StructureV1, 
                      plan: PlanV1):
        """
        Save all required audit files.
        
        Args:
            before_docx: Path to original DOCX file
            after_docx: Path to modified DOCX file
            before_structure: Original document structure
            after_structure: Modified document structure
            plan: Execution plan that was applied
            
        Raises:
            AuditError: If snapshot saving fails
        """
        if not self.current_audit_dir:
            raise AuditError(
                "No audit directory created. Call create_audit_directory() first.",
                audit_stage="snapshot_saving"
            )
        
        try:
            # Save DOCX snapshots with fixed names
            before_docx_path = Path(before_docx)
            after_docx_path = Path(after_docx)
            
            if before_docx_path.exists():
                shutil.copy2(before_docx_path, self.current_audit_dir / "snapshots" / "before.docx")
            
            if after_docx_path.exists():
                shutil.copy2(after_docx_path, self.current_audit_dir / "snapshots" / "after.docx")
            
            # Save structure JSON files
            with open(self.current_audit_dir / "structures" / "structure.before.v1.json", 'w', encoding='utf-8') as f:
                json.dump(before_structure.model_dump(), f, indent=2, ensure_ascii=False, default=str)
            
            with open(self.current_audit_dir / "structures" / "structure.after.v1.json", 'w', encoding='utf-8') as f:
                json.dump(after_structure.model_dump(), f, indent=2, ensure_ascii=False, default=str)
            
            # Save execution plan
            with open(self.current_audit_dir / "plan.v1.json", 'w', encoding='utf-8') as f:
                json.dump(plan.model_dump(), f, indent=2, ensure_ascii=False, default=str)
            
        except Exception as e:
            raise AuditError(
                f"Failed to save snapshots: {str(e)}",
                audit_directory=str(self.current_audit_dir),
                audit_stage="snapshot_saving"
            )
    
    def generate_diff_report(self, before_structure: StructureV1, 
                           after_structure: StructureV1) -> DiffReport:
        """
        Generate structural difference report.
        
        Args:
            before_structure: Original document structure
            after_structure: Modified document structure
            
        Returns:
            DiffReport: Detailed difference report
            
        Raises:
            AuditError: If diff generation fails
        """
        try:
            # Create paragraph index mappings
            before_paragraphs = {p.index: p for p in before_structure.paragraphs}
            after_paragraphs = {p.index: p for p in after_structure.paragraphs}
            
            before_indices = set(before_paragraphs.keys())
            after_indices = set(after_paragraphs.keys())
            
            # Find added, removed, and potentially modified paragraphs
            added_paragraphs = list(after_indices - before_indices)
            removed_paragraphs = list(before_indices - after_indices)
            
            # Check for modifications in common paragraphs
            modified_paragraphs = []
            common_indices = before_indices & after_indices
            
            for idx in common_indices:
                before_p = before_paragraphs[idx]
                after_p = after_paragraphs[idx]
                
                # Compare key attributes
                if (before_p.style_name != after_p.style_name or
                    before_p.preview_text != after_p.preview_text or
                    before_p.is_heading != after_p.is_heading or
                    before_p.heading_level != after_p.heading_level):
                    modified_paragraphs.append(idx)
            
            # Analyze style changes
            style_changes = self._analyze_style_changes(before_structure.styles, after_structure.styles)
            
            # Analyze heading changes
            heading_changes = self._analyze_heading_changes(before_structure.headings, after_structure.headings)
            
            # Generate summary
            summary_parts = []
            if added_paragraphs:
                summary_parts.append(f"{len(added_paragraphs)} paragraphs added")
            if removed_paragraphs:
                summary_parts.append(f"{len(removed_paragraphs)} paragraphs removed")
            if modified_paragraphs:
                summary_parts.append(f"{len(modified_paragraphs)} paragraphs modified")
            if style_changes:
                summary_parts.append(f"{len(style_changes)} styles changed")
            if heading_changes:
                summary_parts.append(f"{len(heading_changes)} heading changes")
            
            summary = "; ".join(summary_parts) if summary_parts else "No structural changes detected"
            
            diff_report = DiffReport(
                added_paragraphs=sorted(added_paragraphs),
                removed_paragraphs=sorted(removed_paragraphs),
                modified_paragraphs=sorted(modified_paragraphs),
                style_changes=style_changes,
                heading_changes=heading_changes,
                summary=summary
            )
            
            # Save diff report to audit directory
            if self.current_audit_dir:
                with open(self.current_audit_dir / "reports" / "diff.report.json", 'w', encoding='utf-8') as f:
                    json.dump(diff_report.model_dump(), f, indent=2, ensure_ascii=False, default=str)
            
            return diff_report
            
        except Exception as e:
            raise AuditError(
                f"Failed to generate diff report: {str(e)}",
                audit_directory=str(self.current_audit_dir) if self.current_audit_dir else None,
                audit_stage="diff_generation"
            )
    
    def _analyze_style_changes(self, before_styles: List, after_styles: List) -> Dict[str, Dict[str, Any]]:
        """Analyze changes in style definitions."""
        before_style_map = {s.name: s for s in before_styles}
        after_style_map = {s.name: s for s in after_styles}
        
        style_changes = {}
        
        # Check for modified styles
        for style_name in before_style_map:
            if style_name in after_style_map:
                before_style = before_style_map[style_name]
                after_style = after_style_map[style_name]
                
                changes = {}
                
                # Compare font specifications
                if before_style.font != after_style.font:
                    changes['font'] = {
                        'before': before_style.font.model_dump() if before_style.font else None,
                        'after': after_style.font.model_dump() if after_style.font else None
                    }
                
                # Compare paragraph specifications
                if before_style.paragraph != after_style.paragraph:
                    changes['paragraph'] = {
                        'before': before_style.paragraph.model_dump() if before_style.paragraph else None,
                        'after': after_style.paragraph.model_dump() if after_style.paragraph else None
                    }
                
                if changes:
                    style_changes[style_name] = changes
        
        # Check for added styles
        for style_name in after_style_map:
            if style_name not in before_style_map:
                style_changes[style_name] = {
                    'added': after_style_map[style_name].model_dump()
                }
        
        # Check for removed styles
        for style_name in before_style_map:
            if style_name not in after_style_map:
                style_changes[style_name] = {
                    'removed': before_style_map[style_name].model_dump()
                }
        
        return style_changes
    
    def _analyze_heading_changes(self, before_headings: List, after_headings: List) -> List[Dict[str, Any]]:
        """Analyze changes in heading structure."""
        heading_changes = []
        
        # Create mappings for comparison
        before_heading_map = {(h.paragraph_index, h.level, h.text): h for h in before_headings}
        after_heading_map = {(h.paragraph_index, h.level, h.text): h for h in after_headings}
        
        before_keys = set(before_heading_map.keys())
        after_keys = set(after_heading_map.keys())
        
        # Find added headings
        for key in after_keys - before_keys:
            heading = after_heading_map[key]
            heading_changes.append({
                'type': 'added',
                'paragraph_index': heading.paragraph_index,
                'level': heading.level,
                'text': heading.text,
                'style_name': heading.style_name
            })
        
        # Find removed headings
        for key in before_keys - after_keys:
            heading = before_heading_map[key]
            heading_changes.append({
                'type': 'removed',
                'paragraph_index': heading.paragraph_index,
                'level': heading.level,
                'text': heading.text,
                'style_name': heading.style_name
            })
        
        return heading_changes
    
    def add_warning(self, warning: str):
        """
        Add a warning message to be logged.
        
        Args:
            warning: Warning message to add
        """
        self.warnings.append(f"{datetime.now().isoformat()}: {warning}")
    
    def write_warnings_log(self):
        """Write warnings.log file with all accumulated warnings."""
        if not self.current_audit_dir:
            raise AuditError(
                "No audit directory created. Call create_audit_directory() first.",
                audit_stage="warnings_logging"
            )
        
        try:
            warnings_file = self.current_audit_dir / "warnings.log"
            with open(warnings_file, 'w', encoding='utf-8') as f:
                if self.warnings:
                    f.write("\n".join(self.warnings))
                    f.write("\n")
                else:
                    f.write("No warnings generated during processing.\n")
                    
        except Exception as e:
            raise AuditError(
                f"Failed to write warnings log: {str(e)}",
                audit_directory=str(self.current_audit_dir),
                audit_stage="warnings_logging"
            )
    
    def write_status(self, status: str, details: str):
        """
        Write final status (SUCCESS/ROLLBACK/FAILED_VALIDATION).
        
        Args:
            status: Final processing status
            details: Additional status details
            
        Raises:
            AuditError: If status writing fails
        """
        if not self.current_audit_dir:
            raise AuditError(
                "No audit directory created. Call create_audit_directory() first.",
                audit_stage="status_writing"
            )
        
        # Validate status
        valid_statuses = {"SUCCESS", "ROLLBACK", "FAILED_VALIDATION"}
        if status not in valid_statuses:
            raise AuditError(
                f"Invalid status '{status}'. Must be one of: {', '.join(valid_statuses)}",
                audit_directory=str(self.current_audit_dir),
                audit_stage="status_writing"
            )
        
        try:
            status_file = self.current_audit_dir / "result.status.txt"
            timestamp = datetime.now().isoformat()
            
            with open(status_file, 'w', encoding='utf-8') as f:
                f.write(f"STATUS: {status}\n")
                f.write(f"TIMESTAMP: {timestamp}\n")
                f.write(f"DETAILS: {details}\n")
                
                # Add audit directory info
                f.write(f"AUDIT_DIRECTORY: {self.current_audit_dir}\n")
                
                # Add warning count
                f.write(f"WARNING_COUNT: {len(self.warnings)}\n")
                
        except Exception as e:
            raise AuditError(
                f"Failed to write status file: {str(e)}",
                audit_directory=str(self.current_audit_dir),
                audit_stage="status_writing"
            )
    
    def finalize_audit(self, status: str, details: str):
        """
        Finalize the audit trail by writing warnings log and status.
        
        Args:
            status: Final processing status
            details: Additional status details
            
        Raises:
            AuditError: If finalization fails
        """
        try:
            self.write_warnings_log()
            self.write_status(status, details)
            
        except Exception as e:
            raise AuditError(
                f"Failed to finalize audit trail: {str(e)}",
                audit_directory=str(self.current_audit_dir) if self.current_audit_dir else None,
                audit_stage="finalization"
            )
    
    def get_audit_directory(self) -> Optional[str]:
        """
        Get the current audit directory path.
        
        Returns:
            str: Path to current audit directory, or None if not created
        """
        return str(self.current_audit_dir) if self.current_audit_dir else None