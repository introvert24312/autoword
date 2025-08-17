"""
Demonstration of AutoWord vNext LocalizationManager functionality.

This script shows how the LocalizationManager handles:
- Style name aliases (English ↔ Chinese)
- Font fallback chains for East Asian fonts
- Warnings logging for fallback operations
- Dynamic style detection
"""

import tempfile
import os
from unittest.mock import Mock

from autoword.vnext.localization import LocalizationManager, FontFallbackChain, create_localization_config


def demo_style_aliases():
    """Demonstrate style alias functionality."""
    print("=== Style Aliases Demo ===")
    
    # Create a temporary warnings log
    with tempfile.NamedTemporaryFile(mode='w', suffix='.log', delete=False) as f:
        warnings_log_path = f.name
    
    try:
        manager = LocalizationManager(warnings_log_path)
        
        # Mock Word document with some styles
        mock_doc = Mock()
        mock_styles = Mock()
        mock_doc.Styles = mock_styles
        
        # Mock style existence - only Chinese styles exist
        def mock_getitem(key):
            chinese_styles = ["标题 1", "标题 2", "正文"]
            if key in chinese_styles:
                return Mock()
            else:
                raise Exception("Style not found")
        
        mock_styles.__getitem__ = mock_getitem
        
        # Mock _style_exists method for proper testing
        def mock_style_exists(style_name, doc):
            chinese_styles = ["标题 1", "标题 2", "正文"]
            return style_name in chinese_styles
        
        manager._style_exists = mock_style_exists
        
        # Test style resolution
        test_styles = ["Heading 1", "Heading 2", "Normal", "NonExistent"]
        
        for style in test_styles:
            resolved = manager.resolve_style_name(style, mock_doc)
            print(f"  {style} → {resolved}")
        
        # Show warnings
        warnings = manager.get_warnings()
        print(f"\nGenerated {len(warnings)} warnings:")
        for warning in warnings:
            print(f"  - {warning}")
        
        # Write warnings to log
        manager.write_warnings_log()
        print(f"\nWarnings written to: {warnings_log_path}")
        
    finally:
        if os.path.exists(warnings_log_path):
            os.unlink(warnings_log_path)


def demo_font_fallbacks():
    """Demonstrate font fallback functionality."""
    print("\n=== Font Fallbacks Demo ===")
    
    manager = LocalizationManager()
    mock_doc = Mock()
    
    # Test font resolution with different availability scenarios
    test_fonts = ["楷体", "宋体", "黑体", "UnknownFont"]
    
    for font in test_fonts:
        resolved = manager.resolve_font_name(font, mock_doc)
        print(f"  {font} → {resolved}")
    
    # Show warnings
    warnings = manager.get_warnings()
    print(f"\nGenerated {len(warnings)} warnings:")
    for warning in warnings:
        print(f"  - {warning}")


def demo_fallback_chains():
    """Demonstrate font fallback chain creation."""
    print("\n=== Font Fallback Chains Demo ===")
    
    # East Asian font chains
    east_asian_fonts = ["楷体", "宋体", "黑体", "UnknownEastAsianFont"]
    
    print("East Asian font chains:")
    for font in east_asian_fonts:
        chain = FontFallbackChain.create_east_asian_chain(font)
        print(f"  {font}: {' → '.join(chain)}")
    
    # Latin font chains
    latin_fonts = ["Times New Roman", "Arial", "Unknown Serif Font", "Unknown Sans Font"]
    
    print("\nLatin font chains:")
    for font in latin_fonts:
        chain = FontFallbackChain.create_latin_chain(font)
        print(f"  {font}: {' → '.join(chain)}")


def demo_document_style_detection():
    """Demonstrate document style detection."""
    print("\n=== Document Style Detection Demo ===")
    
    manager = LocalizationManager()
    
    # Mock Word document with mixed English and Chinese styles
    mock_doc = Mock()
    mock_styles = Mock()
    mock_doc.Styles = mock_styles
    
    # Create mock styles
    mock_style1 = Mock()
    mock_style1.NameLocal = "Heading 1"
    mock_style2 = Mock()
    mock_style2.NameLocal = "标题 2"
    mock_style3 = Mock()
    mock_style3.NameLocal = "正文"
    mock_style4 = Mock()
    mock_style4.NameLocal = "Custom Style"
    
    mock_styles.__iter__ = Mock(return_value=iter([mock_style1, mock_style2, mock_style3, mock_style4]))
    
    # Detect styles
    style_mapping = manager.detect_document_styles(mock_doc)
    
    print("Detected style mappings:")
    for standard_name, actual_name in style_mapping.items():
        print(f"  {standard_name} → {actual_name}")


def demo_configuration():
    """Demonstrate configuration file creation."""
    print("\n=== Configuration Demo ===")
    
    with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
        config_path = f.name
    
    try:
        config = create_localization_config(config_path)
        print(f"Configuration created at: {config_path}")
        print("Configuration keys:", list(config.keys()))
        
        # Show some sample mappings
        print("\nSample style aliases:")
        for i, (english, chinese) in enumerate(config["style_aliases"].items()):
            if i < 5:  # Show first 5
                print(f"  {english} ↔ {chinese}")
        
        print("\nSample font fallbacks:")
        for i, (font, fallbacks) in enumerate(config["font_fallbacks"].items()):
            if i < 3:  # Show first 3
                print(f"  {font}: {' → '.join(fallbacks)}")
                
    finally:
        if os.path.exists(config_path):
            os.unlink(config_path)


def main():
    """Run all localization demos."""
    print("AutoWord vNext LocalizationManager Demo")
    print("=" * 50)
    
    demo_style_aliases()
    demo_font_fallbacks()
    demo_fallback_chains()
    demo_document_style_detection()
    demo_configuration()
    
    print("\n" + "=" * 50)
    print("Demo completed successfully!")


if __name__ == "__main__":
    main()