"""
AutoWord vNext DocumentExtractor Demo

This script demonstrates how to use the DocumentExtractor to extract
structure and inventory from DOCX files.
"""

import os
import json
import logging
from pathlib import Path

from autoword.vnext.extractor.document_extractor import DocumentExtractor
from autoword.vnext.exceptions import ExtractionError

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


def demo_extract_structure(docx_path: str):
    """Demonstrate structure extraction from a DOCX file."""
    print(f"\n=== Extracting Structure from {docx_path} ===")
    
    try:
        with DocumentExtractor(visible=False) as extractor:
            structure = extractor.extract_structure(docx_path)
            
            print(f"Schema Version: {structure.schema_version}")
            print(f"Document Title: {structure.metadata.title}")
            print(f"Author: {structure.metadata.author}")
            print(f"Page Count: {structure.metadata.page_count}")
            print(f"Paragraph Count: {structure.metadata.paragraph_count}")
            print(f"Word Count: {structure.metadata.word_count}")
            print(f"Styles: {len(structure.styles)}")
            print(f"Paragraphs: {len(structure.paragraphs)}")
            print(f"Headings: {len(structure.headings)}")
            print(f"Fields: {len(structure.fields)}")
            print(f"Tables: {len(structure.tables)}")
            
            # Show first few paragraphs
            print("\nFirst 3 paragraphs:")
            for i, para in enumerate(structure.paragraphs[:3]):
                print(f"  {i}: [{para.style_name}] {para.preview_text}")
                if para.is_heading:
                    print(f"      -> Heading Level {para.heading_level}")
            
            # Show headings
            if structure.headings:
                print("\nHeadings:")
                for heading in structure.headings:
                    print(f"  Level {heading.level}: {heading.text}")
            
            # Show styles
            if structure.styles:
                print(f"\nFirst 5 styles:")
                for style in structure.styles[:5]:
                    print(f"  {style.name} ({style.type})")
                    if style.font:
                        print(f"    Font: {style.font.latin} / {style.font.east_asian}, {style.font.size_pt}pt")
            
            return structure
            
    except ExtractionError as e:
        print(f"Extraction failed: {e}")
        if e.details:
            print(f"Details: {e.details}")
        return None


def demo_extract_inventory(docx_path: str):
    """Demonstrate inventory extraction from a DOCX file."""
    print(f"\n=== Extracting Inventory from {docx_path} ===")
    
    try:
        with DocumentExtractor(visible=False) as extractor:
            inventory = extractor.extract_inventory(docx_path)
            
            print(f"Schema Version: {inventory.schema_version}")
            print(f"OOXML Fragments: {len(inventory.ooxml_fragments)}")
            print(f"Media Files: {len(inventory.media_indexes)}")
            print(f"Content Controls: {len(inventory.content_controls)}")
            print(f"Formulas: {len(inventory.formulas)}")
            print(f"Charts: {len(inventory.charts)}")
            
            # Show OOXML fragments
            if inventory.ooxml_fragments:
                print("\nOOXML Fragments:")
                for fragment_name in list(inventory.ooxml_fragments.keys())[:5]:
                    content_length = len(inventory.ooxml_fragments[fragment_name])
                    print(f"  {fragment_name}: {content_length} characters")
            
            # Show media files
            if inventory.media_indexes:
                print("\nMedia Files:")
                for media_path, media_ref in inventory.media_indexes.items():
                    print(f"  {media_path}: {media_ref.content_type}, {media_ref.size_bytes} bytes")
            
            # Show content controls
            if inventory.content_controls:
                print("\nContent Controls:")
                for cc in inventory.content_controls:
                    print(f"  {cc.control_type} at paragraph {cc.paragraph_index}: {cc.title}")
            
            return inventory
            
    except ExtractionError as e:
        print(f"Inventory extraction failed: {e}")
        if e.details:
            print(f"Details: {e.details}")
        return None


def demo_full_processing(docx_path: str):
    """Demonstrate full document processing (structure + inventory)."""
    print(f"\n=== Full Processing of {docx_path} ===")
    
    try:
        with DocumentExtractor(visible=False) as extractor:
            structure, inventory = extractor.process_document(docx_path)
            
            print("✓ Structure and inventory extracted successfully")
            print(f"  Structure: {len(structure.paragraphs)} paragraphs, {len(structure.headings)} headings")
            print(f"  Inventory: {len(inventory.ooxml_fragments)} OOXML fragments, {len(inventory.media_indexes)} media files")
            
            return structure, inventory
            
    except ExtractionError as e:
        print(f"Full processing failed: {e}")
        if e.details:
            print(f"Details: {e.details}")
        return None, None


def save_extraction_results(docx_path: str, output_dir: str = "extraction_output"):
    """Extract and save structure and inventory to JSON files."""
    print(f"\n=== Saving Extraction Results ===")
    
    # Create output directory
    output_path = Path(output_dir)
    output_path.mkdir(exist_ok=True)
    
    # Get base filename
    base_name = Path(docx_path).stem
    
    try:
        with DocumentExtractor(visible=False) as extractor:
            structure, inventory = extractor.process_document(docx_path)
            
            # Save structure
            structure_file = output_path / f"{base_name}_structure.v1.json"
            with open(structure_file, 'w', encoding='utf-8') as f:
                json.dump(structure.model_dump(), f, indent=2, ensure_ascii=False, default=str)
            print(f"✓ Structure saved to: {structure_file}")
            
            # Save inventory
            inventory_file = output_path / f"{base_name}_inventory.full.v1.json"
            with open(inventory_file, 'w', encoding='utf-8') as f:
                json.dump(inventory.model_dump(), f, indent=2, ensure_ascii=False, default=str)
            print(f"✓ Inventory saved to: {inventory_file}")
            
            return structure_file, inventory_file
            
    except ExtractionError as e:
        print(f"Failed to save extraction results: {e}")
        return None, None


def main():
    """Main demo function."""
    print("AutoWord vNext DocumentExtractor Demo")
    print("=" * 50)
    
    # Look for sample DOCX files
    sample_files = []
    
    # Check for common sample file locations
    possible_paths = [
        "test_document.docx",
        "sample.docx",
        "demo.docx",
        "../test_document.docx",
        "examples/sample.docx"
    ]
    
    for path in possible_paths:
        if os.path.exists(path):
            sample_files.append(path)
    
    if not sample_files:
        print("No sample DOCX files found. Please provide a DOCX file path.")
        print("Expected locations:")
        for path in possible_paths:
            print(f"  - {path}")
        return
    
    # Process each sample file
    for docx_path in sample_files:
        print(f"\nProcessing: {docx_path}")
        
        # Demo structure extraction
        structure = demo_extract_structure(docx_path)
        
        # Demo inventory extraction
        inventory = demo_extract_inventory(docx_path)
        
        # Demo full processing
        full_structure, full_inventory = demo_full_processing(docx_path)
        
        # Save results
        if full_structure and full_inventory:
            save_extraction_results(docx_path)
        
        print("-" * 50)
    
    print("\nDemo completed!")


if __name__ == "__main__":
    main()