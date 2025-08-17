# Structure V1 Schema Documentation

## Overview

The `structure.v1.json` schema defines the complete document structure extracted from DOCX files. This schema represents the document skeleton with metadata, styles, paragraphs, headings, fields, and tables.

## Schema Version

- **Version**: `structure.v1`
- **Purpose**: Document structure representation for planning and validation
- **File Extension**: `.json`

## Root Object

```json
{
  "schema_version": "structure.v1",
  "metadata": { ... },
  "styles": [ ... ],
  "paragraphs": [ ... ],
  "headings": [ ... ],
  "fields": [ ... ],
  "tables": [ ... ]
}
```

## Field Definitions

### schema_version (required)
- **Type**: `string`
- **Value**: `"structure.v1"`
- **Description**: Schema version identifier

### metadata (required)
Document metadata extracted from DOCX properties.

```json
{
  "title": "Document Title",
  "author": "Author Name",
  "subject": "Document Subject",
  "keywords": "keyword1, keyword2",
  "comments": "Document comments",
  "created_time": "2024-01-01T00:00:00Z",
  "modified_time": "2024-01-01T12:00:00Z",
  "word_version": "16.0",
  "page_count": 10,
  "word_count": 1500,
  "character_count": 8000
}
```

### styles (required)
Array of style definitions used in the document.

```json
{
  "name": "Heading 1",
  "type": "paragraph",
  "font": {
    "east_asian": "楷体",
    "latin": "Times New Roman",
    "size_pt": 12,
    "bold": true,
    "italic": false,
    "color_hex": "#000000"
  },
  "paragraph": {
    "line_spacing_mode": "MULTIPLE",
    "line_spacing_value": 2.0,
    "space_before_pt": 12,
    "space_after_pt": 6,
    "alignment": "LEFT"
  }
}
```

#### Style Types
- `paragraph`: Paragraph styles
- `character`: Character styles
- `table`: Table styles
- `linked`: Linked paragraph/character styles

#### Line Spacing Modes
- `SINGLE`: Single line spacing
- `MULTIPLE`: Multiple line spacing (value = multiplier)
- `EXACTLY`: Exact spacing in points

### paragraphs (required)
Array of paragraph skeletons with preview text only.

```json
{
  "index": 0,
  "style_name": "Normal",
  "preview_text": "This is the beginning of the paragraph content...",
  "is_heading": false,
  "level": null,
  "has_direct_formatting": false,
  "in_table": false,
  "table_index": null
}
```

**Important**: Preview text is limited to ≤120 characters to maintain skeleton-only representation.

### headings (required)
Array of heading references for navigation and TOC generation.

```json
{
  "index": 5,
  "level": 1,
  "text": "Chapter 1: Introduction",
  "style_name": "Heading 1",
  "paragraph_index": 5,
  "page_number": 2,
  "outline_level": 1
}
```

### fields (required)
Array of field references (TOC, page numbers, cross-references, etc.).

```json
{
  "index": 0,
  "type": "TOC",
  "code": "TOC \\o \"1-3\" \\h \\z \\u",
  "result": "Table of Contents display text",
  "paragraph_index": 2,
  "is_locked": false,
  "needs_update": false
}
```

#### Field Types
- `TOC`: Table of Contents
- `PAGE`: Page number
- `REF`: Cross-reference
- `HYPERLINK`: Hyperlink
- `DATE`: Date/time
- `FILENAME`: File name

### tables (optional)
Array of table skeletons with structure information.

```json
{
  "index": 0,
  "paragraph_index": 10,
  "rows": 3,
  "columns": 4,
  "has_header": true,
  "style_name": "Table Grid",
  "cell_references": [
    {
      "row": 0,
      "column": 0,
      "paragraph_indexes": [11, 12]
    }
  ]
}
```

## Usage Examples

### Basic Document Structure
```json
{
  "schema_version": "structure.v1",
  "metadata": {
    "title": "Research Paper",
    "author": "John Doe",
    "created_time": "2024-01-01T00:00:00Z",
    "modified_time": "2024-01-01T12:00:00Z",
    "page_count": 5,
    "word_count": 1200
  },
  "styles": [
    {
      "name": "Heading 1",
      "type": "paragraph",
      "font": {
        "east_asian": "楷体",
        "latin": "Times New Roman",
        "size_pt": 12,
        "bold": true
      }
    }
  ],
  "paragraphs": [
    {
      "index": 0,
      "style_name": "Heading 1",
      "preview_text": "Introduction",
      "is_heading": true,
      "level": 1
    }
  ],
  "headings": [
    {
      "index": 0,
      "level": 1,
      "text": "Introduction",
      "paragraph_index": 0
    }
  ],
  "fields": [],
  "tables": []
}
```

## Validation Rules

1. **schema_version** must be exactly `"structure.v1"`
2. **metadata** must contain at least `title`, `created_time`, `modified_time`
3. **paragraphs** preview_text must be ≤120 characters
4. **headings** must reference valid paragraph_index values
5. **styles** must have valid font specifications
6. **tables** cell_references must point to valid paragraph indexes

## Error Handling

Invalid structure.v1.json files will be rejected with detailed error messages:

- Missing required fields
- Invalid schema version
- Preview text exceeding 120 characters
- Invalid cross-references between objects
- Malformed style specifications

## Related Schemas

- `plan.v1.json`: Execution plan based on structure
- `inventory.full.v1.json`: Complete document inventory