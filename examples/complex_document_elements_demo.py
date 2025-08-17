"""
Demo script for complex document element handling in AutoWord vNext.

This script demonstrates the enhanced extractor functionality for:
- Tables with cell references to paragraph indexes
- Formula and content control detection with inventory storage
- Chart/SmartArt/OLE object handling as OOXML/binary in inventory
- Footnote/endnote reference preservation in structure extraction
- Cross-reference identification and relationship mapping
"""

import os
import sys
import json
from pathlib import Path

# Add the project root to the Python path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from autoword.vnext.extractor.document_extractor import DocumentExtractor
from autoword.vnext.exceptions import ExtractionError


def demo_complex_document_extraction():
    """Demonstrate complex document element extraction."""
    print("=== AutoWord vNext Complex Document Element Extraction Demo ===\n")
    
    # Note: This demo would work with a real DOCX file
    # For demonstration purposes, we'll show the structure
    
    print("Enhanced Extractor Features:")
    print("1. Tables with cell-to-paragraph mapping")
    print("2. Enhanced formula and equation detection")
    print("3. Comprehensive chart/SmartArt/OLE object handling")
    print("4. Footnote and endnote reference preservation")
    print("5. Cross-reference identification and relationship mapping")
    print("6. Enhanced OOXML fragment extraction with binary object support")
    print()
    
    # Example of what the enhanced extraction would produce
    print("Example Enhanced Table Structure:")
    table_example = {
        "paragraph_index": 5,
        "rows": 3,
        "columns": 4,
        "has_header": True,
        "cell_references": [6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17],
        "cell_paragraph_map": {
            "1,1": [6],
            "1,2": [7],
            "1,3": [8],
            "1,4": [9],
            "2,1": [10, 11],
            "2,2": [12],
            "2,3": [13],
            "2,4": [14],
            "3,1": [15],
            "3,2": [16],
            "3,3": [17],
            "3,4": []
        }
    }
    print(json.dumps(table_example, indent=2))
    print()
    
    print("Example Enhanced Formula Detection:")
    formula_examples = [
        {
            "paragraph_index": 12,
            "formula_id": "eq_0",
            "formula_type": "equation",
            "latex_code": "x^2 + y^2 = z^2"
        },
        {
            "paragraph_index": 25,
            "formula_id": "excel_1",
            "formula_type": "excel_object",
            "latex_code": None
        },
        {
            "paragraph_index": 38,
            "formula_id": "omath_2",
            "formula_type": "omath",
            "latex_code": "∫₀^∞ e^(-x²) dx = √π/2"
        }
    ]
    for formula in formula_examples:
        print(json.dumps(formula, indent=2))
    print()
    
    print("Example Enhanced Chart/Object Detection:")
    chart_examples = [
        {
            "paragraph_index": 15,
            "chart_id": "chart_0",
            "chart_type": "chart",
            "title": "Sales Performance Q1-Q4"
        },
        {
            "paragraph_index": 28,
            "chart_id": "smartart_1",
            "chart_type": "smartart",
            "title": None
        },
        {
            "paragraph_index": 42,
            "chart_id": "picture_2",
            "chart_type": "picture",
            "title": "Company Logo"
        },
        {
            "paragraph_index": 55,
            "chart_id": "embedded_ole_3",
            "chart_type": "embedded_ole",
            "title": None
        }
    ]
    for chart in chart_examples:
        print(json.dumps(chart, indent=2))
    print()
    
    print("Example Footnote References:")
    footnote_examples = [
        {
            "paragraph_index": 8,
            "footnote_id": "footnote_1",
            "reference_mark": "1",
            "text_preview": "According to Smith et al. (2023), this methodology has been proven effective in multiple studies."
        },
        {
            "paragraph_index": 23,
            "footnote_id": "footnote_2",
            "reference_mark": "2",
            "text_preview": "See appendix A for detailed calculations and assumptions used in this analysis."
        }
    ]
    for footnote in footnote_examples:
        print(json.dumps(footnote, indent=2))
    print()
    
    print("Example Cross-Reference Mapping:")
    cross_ref_examples = [
        {
            "source_paragraph_index": 18,
            "target_paragraph_index": 45,
            "reference_type": "bookmark",
            "reference_text": "Figure 3.2",
            "target_id": "_Ref123456789"
        },
        {
            "source_paragraph_index": 32,
            "target_paragraph_index": 12,
            "reference_type": "internal_link",
            "reference_text": "See Section 2.1",
            "target_id": "_Toc987654321"
        },
        {
            "source_paragraph_index": 67,
            "target_paragraph_index": None,
            "reference_type": "hyperlink",
            "reference_text": "Visit our website",
            "target_id": "https://example.com"
        }
    ]
    for cross_ref in cross_ref_examples:
        print(json.dumps(cross_ref, indent=2))
    print()
    
    print("Enhanced OOXML Fragment Storage:")
    ooxml_examples = [
        "word/document.xml - Main document structure",
        "word/charts/chart1.xml - Chart definition",
        "word/charts/chart1.bin - base64:iVBORw0KGgoAAAANSUhEUgAA... (Binary chart data)",
        "word/drawings/drawing1.xml - Drawing markup",
        "word/embeddings/oleObject1.bin - base64:UEsDBBQAAAAIAA... (Embedded object)",
        "word/activeX/activeX1.xml - ActiveX control definition",
        "customXml/item1.xml - Custom XML data",
        "word/vbaProject.bin - base64:TVqQAAMAAAAEAAAA... (VBA project binary)",
        "word/footnotes.xml - Footnote definitions",
        "word/endnotes.xml - Endnote definitions"
    ]
    for example in ooxml_examples:
        print(f"  - {example}")
    print()
    
    print("Benefits of Enhanced Complex Element Handling:")
    print("✓ Zero information loss - All document elements preserved")
    print("✓ Precise relationship mapping - Tables, footnotes, cross-refs linked to paragraphs")
    print("✓ Binary object support - Charts, OLE objects stored as base64 in inventory")
    print("✓ Enhanced formula detection - Equations, OMath, embedded objects")
    print("✓ Complete cross-reference tracking - Internal links, bookmarks, hyperlinks")
    print("✓ Comprehensive OOXML coverage - All document parts extracted and preserved")
    print()


def demo_with_real_document(docx_path: str):
    """Demonstrate extraction with a real document if available."""
    if not os.path.exists(docx_path):
        print(f"Document not found: {docx_path}")
        return
    
    print(f"=== Extracting Complex Elements from: {docx_path} ===\n")
    
    try:
        with DocumentExtractor(visible=False) as extractor:
            print("Extracting structure and inventory...")
            structure, inventory = extractor.process_document(docx_path)
            
            print(f"Structure extracted:")
            print(f"  - {len(structure.paragraphs)} paragraphs")
            print(f"  - {len(structure.headings)} headings")
            print(f"  - {len(structure.tables)} tables")
            print(f"  - {len(structure.fields)} fields")
            print(f"  - {len(structure.styles)} styles")
            print()
            
            print(f"Inventory extracted:")
            print(f"  - {len(inventory.ooxml_fragments)} OOXML fragments")
            print(f"  - {len(inventory.media_indexes)} media files")
            print(f"  - {len(inventory.content_controls)} content controls")
            print(f"  - {len(inventory.formulas)} formulas")
            print(f"  - {len(inventory.charts)} charts/objects")
            print(f"  - {len(inventory.footnotes)} footnotes")
            print(f"  - {len(inventory.endnotes)} endnotes")
            print(f"  - {len(inventory.cross_references)} cross-references")
            print()
            
            # Show enhanced table details
            if structure.tables:
                print("Enhanced Table Details:")
                for i, table in enumerate(structure.tables):
                    print(f"  Table {i+1}:")
                    print(f"    - Position: paragraph {table.paragraph_index}")
                    print(f"    - Dimensions: {table.rows}x{table.columns}")
                    print(f"    - Has header: {table.has_header}")
                    if table.cell_references:
                        print(f"    - Cell paragraphs: {len(table.cell_references)} references")
                    if table.cell_paragraph_map:
                        print(f"    - Cell mapping: {len(table.cell_paragraph_map)} cells mapped")
                print()
            
            # Show formula details
            if inventory.formulas:
                print("Formula Details:")
                for formula in inventory.formulas:
                    print(f"  - {formula.formula_type} at paragraph {formula.paragraph_index}")
                    if formula.latex_code:
                        preview = formula.latex_code[:50] + "..." if len(formula.latex_code) > 50 else formula.latex_code
                        print(f"    Code: {preview}")
                print()
            
            # Show chart/object details
            if inventory.charts:
                print("Chart/Object Details:")
                for chart in inventory.charts:
                    print(f"  - {chart.chart_type} at paragraph {chart.paragraph_index}")
                    if chart.title:
                        print(f"    Title: {chart.title}")
                print()
            
            # Show footnote details
            if inventory.footnotes:
                print("Footnote Details:")
                for footnote in inventory.footnotes:
                    print(f"  - Footnote {footnote.reference_mark} at paragraph {footnote.paragraph_index}")
                    if footnote.text_preview:
                        preview = footnote.text_preview[:50] + "..." if len(footnote.text_preview) > 50 else footnote.text_preview
                        print(f"    Preview: {preview}")
                print()
            
            # Show cross-reference details
            if inventory.cross_references:
                print("Cross-Reference Details:")
                for cross_ref in inventory.cross_references:
                    print(f"  - {cross_ref.reference_type} from paragraph {cross_ref.source_paragraph_index}")
                    if cross_ref.target_paragraph_index is not None:
                        print(f"    Target: paragraph {cross_ref.target_paragraph_index}")
                    print(f"    Text: {cross_ref.reference_text}")
                print()
            
    except ExtractionError as e:
        print(f"Extraction error: {e}")
    except Exception as e:
        print(f"Unexpected error: {e}")


if __name__ == "__main__":
    # Run the demo
    demo_complex_document_extraction()
    
    # If a document path is provided, try to extract from it
    if len(sys.argv) > 1:
        docx_path = sys.argv[1]
        demo_with_real_document(docx_path)
    else:
        print("To test with a real document, run:")
        print("python examples/complex_document_elements_demo.py path/to/your/document.docx")