"""
Extractor module for converting DOCX to structured JSON representations.

This module handles the extraction of document structure and inventory
with zero information loss through the inventory system.
"""

from .document_extractor import DocumentExtractor

__all__ = ["DocumentExtractor"]