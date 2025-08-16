"""
AutoWord GUI Module
图形用户界面模块
"""

__version__ = "1.0.0"
__author__ = "AutoWord Team"

from .main_window import MainWindow
from .config_manager import ConfigurationManager
from .processor_controller import DocumentProcessorController
from .error_handler import ErrorHandler

__all__ = [
    'MainWindow',
    'ConfigurationManager', 
    'DocumentProcessorController',
    'ErrorHandler'
]