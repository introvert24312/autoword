"""
Auditor module for complete audit trails and snapshots.

This module creates timestamped audit directories with complete
before/after snapshots and diff reports.
"""

from .document_auditor import DocumentAuditor

__all__ = ["DocumentAuditor"]