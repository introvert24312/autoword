"""
Localization manager for AutoWord vNext with style aliases and font fallbacks.

This module provides comprehensive localization support including:
- Style name aliases (English ↔ Chinese)
- Font fallback chains for East Asian fonts
- Warnings logging for fallback operations
"""

import logging
from typing import Dict, List, Optional, Set
from pathlib import Path

from .exceptions import LocalizationError


logger = logging.getLogger(__name__)


class LocalizationManager:
    """Manages style aliases and font fallbacks for localization."""
    
    # Style name aliases (English ↔ Chinese)
    STYLE_ALIASES = {
        "Heading 1": "标题 1",
        "Heading 2": "标题 2", 
        "Heading 3": "标题 3",
        "Heading 4": "标题 4",
        "Heading 5": "标题 5",
        "Heading 6": "标题 6",
        "Heading 7": "标题 7",
        "Heading 8": "标题 8",
        "Heading 9": "标题 9",
        "Normal": "正文",
        "Title": "标题",
        "Subtitle": "副标题",
        "Body Text": "正文",
        "Caption": "题注",
        "Header": "页眉",
        "Footer": "页脚"
    }
    
    # Font fallback chains for East Asian fonts
    FONT_FALLBACKS = {
        "楷体": ["楷体", "楷体_GB2312", "STKaiti", "KaiTi"],
        "宋体": ["宋体", "SimSun", "NSimSun", "宋体-简"],
        "黑体": ["黑体", "SimHei", "Microsoft YaHei", "微软雅黑"],
        "仿宋": ["仿宋", "FangSong", "FangSong_GB2312", "仿宋_GB2312"],
        "微软雅黑": ["微软雅黑", "Microsoft YaHei", "Microsoft YaHei UI"],
        "Times New Roman": ["Times New Roman", "Times", "serif"],
        "Arial": ["Arial", "Helvetica", "sans-serif"],
        "Calibri": ["Calibri", "Arial", "sans-serif"]
    }
    
    def __init__(self, warnings_log_path: Optional[str] = None):
        """
        Initialize localization manager.
        
        Args:
            warnings_log_path: Path to warnings.log file for logging fallbacks
        """
        self.warnings_log_path = warnings_log_path
        self._warnings_buffer: List[str] = []
        
    def resolve_style_name(self, style_name: str, doc: object) -> str:
        """
        Resolve style name using aliases with dynamic detection.
        
        Args:
            style_name: Style name to resolve
            doc: Word document COM object
            
        Returns:
            str: Resolved style name that exists in document
        """
        try:
            # Try original name first
            if self._style_exists(style_name, doc):
                return style_name
                
            # Try direct alias lookup (English → Chinese)
            if style_name in self.STYLE_ALIASES:
                alias = self.STYLE_ALIASES[style_name]
                if self._style_exists(alias, doc):
                    self._log_warning(f"Style alias used: {style_name} → {alias}")
                    return alias
                    
            # Try reverse alias lookup (Chinese → English)
            for english, chinese in self.STYLE_ALIASES.items():
                if style_name == chinese and self._style_exists(english, doc):
                    self._log_warning(f"Style alias used: {style_name} → {english}")
                    return english
            
            # Try case-insensitive matching
            for style in doc.Styles:
                try:
                    if style.NameLocal.lower() == style_name.lower():
                        if style.NameLocal != style_name:
                            self._log_warning(f"Style case mismatch: {style_name} → {style.NameLocal}")
                        return style.NameLocal
                except:
                    continue
            
            # No alias found, return original
            self._log_warning(f"Style not found, using original: {style_name}")
            return style_name
            
        except Exception as e:
            logger.warning(f"Error resolving style name '{style_name}': {e}")
            return style_name
    
    def resolve_font_name(self, font_name: str, doc: object) -> str:
        """
        Resolve font name using fallback chain with availability detection.
        
        Args:
            font_name: Font name to resolve
            doc: Word document COM object
            
        Returns:
            str: Resolved font name (first available from fallback chain)
        """
        try:
            # Try original font first
            if self._font_exists(font_name, doc):
                return font_name
                
            # Try fallback chain
            if font_name in self.FONT_FALLBACKS:
                fallback_chain = self.FONT_FALLBACKS[font_name]
                for fallback_font in fallback_chain[1:]:  # Skip first (original)
                    if self._font_exists(fallback_font, doc):
                        self._log_warning(f"Font fallback: {font_name} → {fallback_font}")
                        return fallback_font
                        
                # No fallback found in chain
                self._log_warning(f"Font not available: {font_name} (no fallback found in chain)")
            else:
                # No fallback chain defined
                self._log_warning(f"Font not available: {font_name} (no fallback chain defined)")
            
            # Return original font name as last resort
            return font_name
            
        except Exception as e:
            logger.warning(f"Error resolving font name '{font_name}': {e}")
            return font_name
    
    def detect_document_styles(self, doc: object) -> Dict[str, str]:
        """
        Detect available styles in document and create mapping.
        
        Args:
            doc: Word document COM object
            
        Returns:
            Dict[str, str]: Mapping of standard names to actual style names
        """
        style_mapping = {}
        
        try:
            available_styles = set()
            for style in doc.Styles:
                try:
                    available_styles.add(style.NameLocal)
                except:
                    continue
            
            # Map standard style names to available styles
            for standard_name, localized_name in self.STYLE_ALIASES.items():
                if standard_name in available_styles:
                    style_mapping[standard_name] = standard_name
                elif localized_name in available_styles:
                    style_mapping[standard_name] = localized_name
                    
        except Exception as e:
            logger.warning(f"Error detecting document styles: {e}")
            
        return style_mapping
    
    def get_warnings(self) -> List[str]:
        """
        Get accumulated warnings.
        
        Returns:
            List[str]: List of warning messages
        """
        return self._warnings_buffer.copy()
    
    def clear_warnings(self):
        """Clear accumulated warnings."""
        self._warnings_buffer.clear()
    
    def write_warnings_log(self, additional_warnings: Optional[List[str]] = None):
        """
        Write warnings to warnings.log file.
        
        Args:
            additional_warnings: Additional warnings to include
        """
        if not self.warnings_log_path:
            return
            
        try:
            warnings_to_write = self._warnings_buffer.copy()
            if additional_warnings:
                warnings_to_write.extend(additional_warnings)
                
            if warnings_to_write:
                with open(self.warnings_log_path, 'a', encoding='utf-8') as f:
                    for warning in warnings_to_write:
                        f.write(f"{warning}\n")
                        
                logger.info(f"Wrote {len(warnings_to_write)} warnings to {self.warnings_log_path}")
                
        except Exception as e:
            logger.error(f"Failed to write warnings log: {e}")
    
    def _style_exists(self, style_name: str, doc: object) -> bool:
        """Check if style exists in document."""
        try:
            doc.Styles[style_name]
            return True
        except:
            return False
    
    def _font_exists(self, font_name: str, doc: object) -> bool:
        """
        Check if font is available in system.
        
        This is a simplified implementation. In production, this would
        check system fonts more thoroughly.
        """
        try:
            # For now, we'll use a heuristic approach
            # In a full implementation, this would check system font registry
            common_fonts = {
                "Arial", "Times New Roman", "Calibri", "Helvetica", "serif", "sans-serif",
                "楷体", "宋体", "黑体", "仿宋", "微软雅黑", "SimSun", "SimHei", "Microsoft YaHei",
                "楷体_GB2312", "仿宋_GB2312", "STKaiti", "KaiTi", "FangSong"
            }
            return font_name in common_fonts
        except:
            return False
    
    def _log_warning(self, message: str):
        """Log warning message to buffer and logger."""
        self._warnings_buffer.append(message)
        logger.warning(message)


class FontFallbackChain:
    """Utility class for managing font fallback chains."""
    
    @staticmethod
    def create_east_asian_chain(primary_font: str) -> List[str]:
        """
        Create fallback chain for East Asian fonts.
        
        Args:
            primary_font: Primary font name
            
        Returns:
            List[str]: Fallback chain starting with primary font
        """
        base_chains = {
            "楷体": ["楷体", "楷体_GB2312", "STKaiti", "KaiTi", "serif"],
            "宋体": ["宋体", "SimSun", "NSimSun", "宋体-简", "serif"],
            "黑体": ["黑体", "SimHei", "Microsoft YaHei", "微软雅黑", "sans-serif"],
            "仿宋": ["仿宋", "FangSong", "FangSong_GB2312", "仿宋_GB2312", "serif"]
        }
        
        if primary_font in base_chains:
            return base_chains[primary_font]
        else:
            # Generic fallback for unknown East Asian fonts
            return [primary_font, "SimSun", "Microsoft YaHei", "serif"]
    
    @staticmethod
    def create_latin_chain(primary_font: str) -> List[str]:
        """
        Create fallback chain for Latin fonts.
        
        Args:
            primary_font: Primary font name
            
        Returns:
            List[str]: Fallback chain starting with primary font
        """
        base_chains = {
            "Times New Roman": ["Times New Roman", "Times", "serif"],
            "Arial": ["Arial", "Helvetica", "sans-serif"],
            "Calibri": ["Calibri", "Arial", "Helvetica", "sans-serif"],
            "Helvetica": ["Helvetica", "Arial", "sans-serif"]
        }
        
        if primary_font in base_chains:
            return base_chains[primary_font]
        else:
            # Generic fallback for unknown Latin fonts
            if "serif" in primary_font.lower():
                return [primary_font, "Times New Roman", "serif"]
            else:
                return [primary_font, "Arial", "sans-serif"]


def create_localization_config(config_path: str) -> Dict:
    """
    Create localization configuration file.
    
    Args:
        config_path: Path to save configuration file
        
    Returns:
        Dict: Configuration dictionary
    """
    config = {
        "style_aliases": LocalizationManager.STYLE_ALIASES,
        "font_fallbacks": LocalizationManager.FONT_FALLBACKS,
        "enable_warnings_log": True,
        "case_sensitive_styles": False,
        "auto_detect_styles": True
    }
    
    try:
        import json
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
        logger.info(f"Created localization config: {config_path}")
    except Exception as e:
        logger.error(f"Failed to create localization config: {e}")
        
    return config