# AutoWord vNext DocumentExtractor

The DocumentExtractor is the first component in the AutoWord vNext pipeline, responsible for converting DOCX files into structured JSON representations with zero information loss.

## Overview

The DocumentExtractor implements a two-phase extraction process:

1. **Structure Extraction**: Extracts document skeleton and metadata into `structure.v1.json`
2. **Inventory Extraction**: Captures complete inventory including OOXML fragments into `inventory.full.v1.json`

## Features

### Structure Extraction (`structure.v1.json`)
- Document metadata (title, author, creation/modification times, Word version)
- Style definitions with font and paragraph specifications
- Paragraph skeletons with preview text (≤120 characters)
- Heading references with levels and text
- Field references (TOC, page numbers, etc.)
- Table skeletons with basic structure

### Inventory Extraction (`inventory.full.v1.json`)
- OOXML fragments from key document files
- Media file indexes with content types and sizes
- Content control references
- Formula references (equations, etc.)
- Chart and SmartArt references

## Usage

### Basic Usage

```python
from autoword.vnext.extractor.document_extractor import DocumentExtractor

# Extract structure only
with DocumentExtractor(visible=False) as extractor:
    structure = extractor.extract_structure("document.docx")
    print(f"Found {len(structure.paragraphs)} paragraphs")

# Extract inventory only
with DocumentExtractor(visible=False) as extractor:
    inventory = extractor.extract_inventory("document.docx")
    print(f"Found {len(inventory.media_indexes)} media files")

# Extract both structure and inventory
with DocumentExtractor(visible=False) as extractor:
    structure, inventory = extractor.process_document("document.docx")
```

### Context Manager

The DocumentExtractor uses a context manager to properly handle Word COM resources:

```python
with DocumentExtractor(visible=False) as extractor:
    # COM resources are automatically initialized
    structure = extractor.extract_structure("document.docx")
    # COM resources are automatically cleaned up
```

### Error Handling

```python
from autoword.vnext.exceptions import ExtractionError

try:
    with DocumentExtractor() as extractor:
        structure = extractor.extract_structure("document.docx")
except ExtractionError as e:
    print(f"Extraction failed: {e}")
    print(f"Details: {e.details}")
```

## Data Models

### StructureV1
- `schema_version`: Always "structure.v1"
- `metadata`: Document metadata (title, author, etc.)
- `styles`: List of style definitions
- `paragraphs`: List of paragraph skeletons with preview text
- `headings`: List of heading references
- `fields`: List of field references
- `tables`: List of table skeletons

### InventoryFullV1
- `schema_version`: Always "inventory.full.v1"
- `ooxml_fragments`: Dictionary of OOXML content by file path
- `media_indexes`: Dictionary of media references by file path
- `content_controls`: List of content control references
- `formulas`: List of formula references
- `charts`: List of chart/SmartArt references

## Implementation Details

### Word COM Integration
- Uses `win32com.client` for Word automation
- Proper COM resource management with context managers
- Error handling for COM exceptions
- Support for both visible and invisible Word instances

### Zero Information Loss
- Structure extraction captures document skeleton only
- Inventory extraction preserves all sensitive/large objects
- OOXML fragments stored as raw XML strings
- Media files indexed with metadata

### Localization Support
- Handles both English and Chinese style names
- Font fallback support for East Asian fonts
- Proper encoding for Unicode content

### Performance Optimizations
- Efficient paragraph text truncation
- Lazy loading of complex objects
- Minimal COM round-trips
- Proper memory management

## Testing

The DocumentExtractor includes comprehensive unit tests with mock Word COM objects:

```bash
python -m pytest tests/test_vnext_extractor.py -v
```

Test coverage includes:
- Context manager initialization and cleanup
- Structure extraction with various document types
- Inventory extraction with media and complex objects
- Error handling and edge cases
- Mock Word COM object interactions

## Requirements

- Python 3.8+
- `win32com.client` (pywin32)
- `pydantic` for data validation
- Microsoft Word installed on the system

## Limitations

- Requires Microsoft Word to be installed
- Single-threaded due to COM limitations
- Memory usage scales with document size
- Some complex objects may require specialized handling

## Next Steps

After extraction, the structure and inventory are passed to the Planner module for LLM-based plan generation.