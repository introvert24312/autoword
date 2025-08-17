"""
Executor module for executing atomic operations through Word COM.

This module performs only whitelisted atomic operations with strict
safety controls and localization support.
"""

from .document_executor import DocumentExecutor

__all__ = ["DocumentExecutor"]