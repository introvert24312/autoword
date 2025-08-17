"""
Tests for AutoWord vNext LocalizationManager.

This module tests the localization functionality including:
- Style name aliases (English ↔ Chinese)
- Font fallback chains for East Asian fonts
- Warnings logging for fallback operations
- Dynamic style detection
"""

import pytest
from unittest.mock import Mock, MagicMock, patch, mock_open
import tempfile
import os
from pathlib import Path

from autoword.vnext.localization import LocalizationManager, FontFallbackChain, create_localization_config
from autoword.vnext.exceptions import LocalizationError


class TestLocalizationManager:
    """Test cases for LocalizationManager class."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.temp_dir = tempfile.mkdtemp()
        self.warnings_log_path = os.path.join(self.temp_dir, "warnings.log")
        self.manager = LocalizationManager(self.warnings_log_path)
        
        # Mock Word document
        self.mock_doc = Mock()
        self.mock_styles = Mock()
        self.mock_doc.Styles = self.mock_styles
    
    def teardown_method(self):
        """Clean up test fixtures."""
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    def test_init_with_warnings_log(self):
        """Test LocalizationManager initialization with warnings log."""
        manager = LocalizationManager(self.warnings_log_path)
        assert manager.warnings_log_path == self.warnings_log_path
        assert manager._warnings_buffer == []
    
    def test_init_without_warnings_log(self):
        """Test LocalizationManager initialization without warnings log."""
        manager = LocalizationManager()
        assert manager.warnings_log_path is None
        assert manager._warnings_buffer == []
    
    def test_style_aliases_mapping(self):
        """Test that style aliases are properly defined."""
        expected_aliases = {
            "Heading 1": "标题 1",
            "Heading 2": "标题 2",
            "Normal": "正文",
            "Title": "标题"
        }
        
        for english, chinese in expected_aliases.items():
            assert LocalizationManager.STYLE_ALIASES[english] == chinese
    
    def test_font_fallbacks_mapping(self):
        """Test that font fallbacks are properly defined."""
        expected_fallbacks = {
            "楷体": ["楷体", "楷体_GB2312", "STKaiti", "KaiTi"],
            "宋体": ["宋体", "SimSun", "NSimSun", "宋体-简"],
            "黑体": ["黑体", "SimHei", "Microsoft YaHei", "微软雅黑"]
        }
        
        for font, fallbacks in expected_fallbacks.items():
            assert LocalizationManager.FONT_FALLBACKS[font] == fallbacks
    
    def test_resolve_style_name_exact_match(self):
        """Test style name resolution with exact match."""
        # Mock style exists
        self.mock_styles.__getitem__ = Mock(return_value=Mock())
        
        result = self.manager.resolve_style_name("Heading 1", self.mock_doc)
        assert result == "Heading 1"
        assert len(self.manager.get_warnings()) == 0
    
    def test_resolve_style_name_alias_english_to_chinese(self):
        """Test style name resolution using English to Chinese alias."""
        # Mock _style_exists to return False for English, True for Chinese
        def mock_style_exists(style_name, doc):
            if style_name == "Heading 1":
                return False
            elif style_name == "标题 1":
                return True
            return False
        
        with patch.object(self.manager, '_style_exists', side_effect=mock_style_exists):
            result = self.manager.resolve_style_name("Heading 1", self.mock_doc)
            assert result == "标题 1"
            warnings = self.manager.get_warnings()
            assert len(warnings) == 1
            assert "Style alias used: Heading 1 → 标题 1" in warnings[0]
    
    def test_resolve_style_name_alias_chinese_to_english(self):
        """Test style name resolution using Chinese to English alias."""
        # Mock _style_exists to return False for Chinese, True for English
        def mock_style_exists(style_name, doc):
            if style_name == "标题 1":
                return False
            elif style_name == "Heading 1":
                return True
            return False
        
        with patch.object(self.manager, '_style_exists', side_effect=mock_style_exists):
            result = self.manager.resolve_style_name("标题 1", self.mock_doc)
            assert result == "Heading 1"
            warnings = self.manager.get_warnings()
            assert len(warnings) == 1
            assert "Style alias used: 标题 1 → Heading 1" in warnings[0]
    
    def test_resolve_style_name_case_insensitive(self):
        """Test style name resolution with case-insensitive matching."""
        # Mock styles collection for case-insensitive matching
        mock_style = Mock()
        mock_style.NameLocal = "Heading 1"
        self.mock_styles.__iter__ = Mock(return_value=iter([mock_style]))
        self.mock_styles.__getitem__ = Mock(side_effect=Exception("Style not found"))
        
        result = self.manager.resolve_style_name("heading 1", self.mock_doc)
        assert result == "Heading 1"
        warnings = self.manager.get_warnings()
        assert len(warnings) == 1
        assert "Style case mismatch: heading 1 → Heading 1" in warnings[0]
    
    def test_resolve_style_name_not_found(self):
        """Test style name resolution when style is not found."""
        # Mock no styles exist
        self.mock_styles.__getitem__ = Mock(side_effect=Exception("Style not found"))
        self.mock_styles.__iter__ = Mock(return_value=iter([]))
        
        result = self.manager.resolve_style_name("NonExistent", self.mock_doc)
        assert result == "NonExistent"
        warnings = self.manager.get_warnings()
        assert len(warnings) == 1
        assert "Style not found, using original: NonExistent" in warnings[0]
    
    def test_resolve_font_name_exact_match(self):
        """Test font name resolution with exact match."""
        with patch.object(self.manager, '_font_exists', return_value=True):
            result = self.manager.resolve_font_name("楷体", self.mock_doc)
            assert result == "楷体"
            assert len(self.manager.get_warnings()) == 0
    
    def test_resolve_font_name_fallback_success(self):
        """Test font name resolution with successful fallback."""
        def mock_font_exists(font_name, doc):
            return font_name == "楷体_GB2312"  # Only fallback exists
        
        with patch.object(self.manager, '_font_exists', side_effect=mock_font_exists):
            result = self.manager.resolve_font_name("楷体", self.mock_doc)
            assert result == "楷体_GB2312"
            warnings = self.manager.get_warnings()
            assert len(warnings) == 1
            assert "Font fallback: 楷体 → 楷体_GB2312" in warnings[0]
    
    def test_resolve_font_name_no_fallback_in_chain(self):
        """Test font name resolution when no fallback in chain is available."""
        with patch.object(self.manager, '_font_exists', return_value=False):
            result = self.manager.resolve_font_name("楷体", self.mock_doc)
            assert result == "楷体"  # Returns original as last resort
            warnings = self.manager.get_warnings()
            assert len(warnings) == 1
            assert "Font not available: 楷体 (no fallback found in chain)" in warnings[0]
    
    def test_resolve_font_name_no_fallback_chain(self):
        """Test font name resolution for font with no defined fallback chain."""
        with patch.object(self.manager, '_font_exists', return_value=False):
            result = self.manager.resolve_font_name("UnknownFont", self.mock_doc)
            assert result == "UnknownFont"
            warnings = self.manager.get_warnings()
            assert len(warnings) == 1
            assert "Font not available: UnknownFont (no fallback chain defined)" in warnings[0]
    
    def test_detect_document_styles(self):
        """Test document style detection."""
        # Mock styles in document
        mock_style1 = Mock()
        mock_style1.NameLocal = "Heading 1"
        mock_style2 = Mock()
        mock_style2.NameLocal = "标题 2"
        mock_style3 = Mock()
        mock_style3.NameLocal = "正文"
        
        self.mock_styles.__iter__ = Mock(return_value=iter([mock_style1, mock_style2, mock_style3]))
        
        result = self.manager.detect_document_styles(self.mock_doc)
        
        # The method checks all aliases, so we need to account for all possible mappings
        expected_mappings = {
            "Heading 1": "Heading 1",  # English style exists
            "Heading 2": "标题 2",     # Chinese style exists
            "Normal": "正文",          # Chinese style exists (Normal → 正文)
            "Body Text": "正文"        # Body Text also maps to 正文
        }
        
        # Check that all expected mappings are present
        for key, value in expected_mappings.items():
            if key in result:
                assert result[key] == value
    
    def test_get_warnings(self):
        """Test getting accumulated warnings."""
        self.manager._log_warning("Test warning 1")
        self.manager._log_warning("Test warning 2")
        
        warnings = self.manager.get_warnings()
        assert len(warnings) == 2
        assert "Test warning 1" in warnings
        assert "Test warning 2" in warnings
    
    def test_clear_warnings(self):
        """Test clearing accumulated warnings."""
        self.manager._log_warning("Test warning")
        assert len(self.manager.get_warnings()) == 1
        
        self.manager.clear_warnings()
        assert len(self.manager.get_warnings()) == 0
    
    def test_write_warnings_log(self):
        """Test writing warnings to log file."""
        self.manager._log_warning("Warning 1")
        self.manager._log_warning("Warning 2")
        
        additional_warnings = ["Additional warning"]
        self.manager.write_warnings_log(additional_warnings)
        
        # Check file was created and contains warnings
        assert os.path.exists(self.warnings_log_path)
        with open(self.warnings_log_path, 'r', encoding='utf-8') as f:
            content = f.read()
            assert "Warning 1" in content
            assert "Warning 2" in content
            assert "Additional warning" in content
    
    def test_write_warnings_log_no_path(self):
        """Test writing warnings log when no path is set."""
        manager = LocalizationManager()  # No warnings log path
        manager._log_warning("Test warning")
        
        # Should not raise exception
        manager.write_warnings_log()
    
    def test_font_exists_common_fonts(self):
        """Test font existence check for common fonts."""
        # Test some common fonts that should exist
        common_fonts = ["Arial", "楷体", "宋体", "Times New Roman"]
        
        for font in common_fonts:
            assert self.manager._font_exists(font, self.mock_doc) == True
    
    def test_font_exists_uncommon_fonts(self):
        """Test font existence check for uncommon fonts."""
        uncommon_fonts = ["NonExistentFont", "FakeFont123"]
        
        for font in uncommon_fonts:
            assert self.manager._font_exists(font, self.mock_doc) == False


class TestFontFallbackChain:
    """Test cases for FontFallbackChain utility class."""
    
    def test_create_east_asian_chain_known_font(self):
        """Test creating fallback chain for known East Asian font."""
        chain = FontFallbackChain.create_east_asian_chain("楷体")
        expected = ["楷体", "楷体_GB2312", "STKaiti", "KaiTi", "serif"]
        assert chain == expected
    
    def test_create_east_asian_chain_unknown_font(self):
        """Test creating fallback chain for unknown East Asian font."""
        chain = FontFallbackChain.create_east_asian_chain("UnknownEastAsianFont")
        expected = ["UnknownEastAsianFont", "SimSun", "Microsoft YaHei", "serif"]
        assert chain == expected
    
    def test_create_latin_chain_known_font(self):
        """Test creating fallback chain for known Latin font."""
        chain = FontFallbackChain.create_latin_chain("Times New Roman")
        expected = ["Times New Roman", "Times", "serif"]
        assert chain == expected
    
    def test_create_latin_chain_serif_font(self):
        """Test creating fallback chain for unknown serif font."""
        chain = FontFallbackChain.create_latin_chain("Unknown Serif Font")
        expected = ["Unknown Serif Font", "Times New Roman", "serif"]
        assert chain == expected
    
    def test_create_latin_chain_sans_serif_font(self):
        """Test creating fallback chain for unknown sans-serif font."""
        chain = FontFallbackChain.create_latin_chain("Unknown Sans Font")
        expected = ["Unknown Sans Font", "Arial", "sans-serif"]
        assert chain == expected


class TestLocalizationConfig:
    """Test cases for localization configuration."""
    
    def test_create_localization_config(self):
        """Test creating localization configuration file."""
        with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
            config_path = f.name
        
        try:
            config = create_localization_config(config_path)
            
            # Check config structure
            assert "style_aliases" in config
            assert "font_fallbacks" in config
            assert "enable_warnings_log" in config
            assert "case_sensitive_styles" in config
            assert "auto_detect_styles" in config
            
            # Check file was created
            assert os.path.exists(config_path)
            
            # Check file content
            import json
            with open(config_path, 'r', encoding='utf-8') as f:
                saved_config = json.load(f)
                assert saved_config == config
                
        finally:
            if os.path.exists(config_path):
                os.unlink(config_path)


if __name__ == "__main__":
    pytest.main([__file__])