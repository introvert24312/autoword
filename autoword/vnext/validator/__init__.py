"""
Validator module for comprehensive validation and rollback capabilities.

This module validates document modifications against strict assertions
and provides rollback functionality on validation failure. It also includes
advanced validation for document integrity, style consistency, cross-reference
validation, accessibility compliance, and quality metrics.
"""

from .document_validator import DocumentValidator

__all__ = ["DocumentValidator"]

# Import advanced validator separately to avoid circular imports
def get_advanced_validator():
    """Get AdvancedValidator class to avoid circular imports."""
    from .advanced_validator import AdvancedValidator
    return AdvancedValidator

def get_quality_metrics():
    """Get QualityMetrics class to avoid circular imports."""
    from .advanced_validator import QualityMetrics
    return QualityMetrics