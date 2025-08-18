# AutoWord vNext Schema Reference

## Overview

AutoWord vNext uses strict JSON schemas to ensure data integrity and validate all data interchange between pipeline modules. This document provides comprehensive reference for all schemas used in the system.

## Schema Validation System

### Schema Files Location
```
autoword/vnext/schemas/
├── structure.v1.json          # Document structure schema
├── plan.v1.json               # Execution plan schema
├── inventory.full.v1.json     # Document inventory schema
├── config.schema.json         # Configuration schema
└── audit.schema.json          # Audit trail schema
```

### Validation Process
1. **Input Validation**: All input data validated against schemas
2. **Output Validation**: All generated data validated before use
3. **Runtime Validation**: Continuous validation during processing
4. **Error Reporting**: Detailed validation error messages

## Structure V1 Schema

### Purpose
Represents the complete document structure with metadata, styles, paragraphs, headings, fields, and tables.

### Schema Definition
```json
{
  "$schema": "http://json-schema.org/draft-07/schema#",
  "$id": "structure.v1.json",
  "title": "Document Structure V1",
  "description": "Complete document structure representation",
  "type": "object",
  "required": ["schema_version", "metadata", "styles", "paragraphs", "headings"],
  "properties": {
    "schema_version": {
      "type": "string",
      "const": "structure.v1"
    },
    "metadata": {
      "$ref": "#/definitions/DocumentMetadata"
    },
    "styles": {
      "type": "array",
      "items": {"$ref": "#/definitions/StyleDefinition"}
    },
    "paragraphs": {
      "type": "array",
      "items": {"$ref": "#/definitions/ParagraphSkeleton"}
    },
    "headings": {
      "type": "array",
      "items": {"$ref": "#/definitions/HeadingReference"}
    },
    "fields": {
      "type": "array",
      "items": {"$ref": "#/definitions/FieldReference"}
    },
    "tables": {
      "type": "array",
      "items": {"$ref": "#/definitions/TableSkeleton"}
    }
  }
}
```

### Data Model Definitions

#### DocumentMetadata
```json
{
  "DocumentMetadata": {
    "type": "object",
    "required": ["title", "author", "created_time", "modified_time", "word_version"],
    "properties": {
      "title": {"type": "string"},
      "author": {"type": "string"},
      "created_time": {"type": "string", "format": "date-time"},
      "modified_time": {"type": "string", "format": "date-time"},
      "word_version": {"type": "string"},
      "page_count": {"type": "integer", "minimum": 1},
      "paragraph_count": {"type": "integer", "minimum": 0},
      "word_count": {"type": "integer", "minimum": 0}
    }
  }
}
```

#### StyleDefinition
```json
{
  "StyleDefinition": {
    "type": "object",
    "required": ["name", "type", "font", "paragraph"],
    "properties": {
      "name": {"type": "string"},
      "type": {
        "type": "string",
        "enum": ["paragraph", "character", "table", "linked"]
      },
      "font": {"$ref": "#/definitions/FontSpec"},
      "paragraph": {"$ref": "#/definitions/ParagraphSpec"},
      "is_builtin": {"type": "boolean"},
      "is_modified": {"type": "boolean"}
    }
  }
}
```

#### FontSpec
```json
{
  "FontSpec": {
    "type": "object",
    "required": ["east_asian", "latin", "size_pt"],
    "properties": {
      "east_asian": {"type": "string"},
      "latin": {"type": "string"},
      "size_pt": {"type": "number", "minimum": 6, "maximum": 72},
      "bold": {"type": "boolean"},
      "italic": {"type": "boolean"},
      "underline": {"type": "boolean"},
      "color_hex": {
        "type": "string",
        "pattern": "^#[0-9A-Fa-f]{6}$"
      }
    }
  }
}
```

#### ParagraphSpec
```json
{
  "ParagraphSpec": {
    "type": "object",
    "properties": {
      "line_spacing_mode": {
        "type": "string",
        "enum": ["SINGLE", "MULTIPLE", "EXACTLY"]
      },
      "line_spacing_value": {"type": "number", "minimum": 0.5, "maximum": 10.0},
      "space_before_pt": {"type": "number", "minimum": 0, "maximum": 100},
      "space_after_pt": {"type": "number", "minimum": 0, "maximum": 100},
      "alignment": {
        "type": "string",
        "enum": ["LEFT", "CENTER", "RIGHT", "JUSTIFY"]
      },
      "indent_left_pt": {"type": "number"},
      "indent_right_pt": {"type": "number"},
      "indent_first_line_pt": {"type": "number"}
    }
  }
}
```

#### HeadingReference
```json
{
  "HeadingReference": {
    "type": "object",
    "required": ["text", "level", "style_name", "paragraph_index"],
    "properties": {
      "text": {"type": "string", "maxLength": 200},
      "level": {"type": "integer", "minimum": 1, "maximum": 9},
      "style_name": {"type": "string"},
      "paragraph_index": {"type": "integer", "minimum": 0},
      "page_number": {"type": "integer", "minimum": 1},
      "in_table": {"type": "boolean"},
      "table_index": {"type": "integer", "minimum": 0}
    }
  }
}
```

#### ParagraphSkeleton
```json
{
  "ParagraphSkeleton": {
    "type": "object",
    "required": ["index", "style_name", "preview_text"],
    "properties": {
      "index": {"type": "integer", "minimum": 0},
      "style_name": {"type": "string"},
      "preview_text": {"type": "string", "maxLength": 120},
      "is_heading": {"type": "boolean"},
      "heading_level": {"type": "integer", "minimum": 1, "maximum": 9},
      "page_number": {"type": "integer", "minimum": 1}
    }
  }
}
```

### Example Structure V1 Document
```json
{
  "schema_version": "structure.v1",
  "metadata": {
    "title": "Research Paper",
    "author": "John Doe",
    "created_time": "2024-01-15T10:30:00Z",
    "modified_time": "2024-01-16T14:45:00Z",
    "word_version": "16.0",
    "page_count": 15,
    "paragraph_count": 120,
    "word_count": 3500
  },
  "styles": [
    {
      "name": "Heading 1",
      "type": "paragraph",
      "font": {
        "east_asian": "楷体",
        "latin": "Times New Roman",
        "size_pt": 12,
        "bold": true,
        "italic": false,
        "underline": false,
        "color_hex": "#000000"
      },
      "paragraph": {
        "line_spacing_mode": "MULTIPLE",
        "line_spacing_value": 2.0,
        "space_before_pt": 12,
        "space_after_pt": 6,
        "alignment": "LEFT"
      },
      "is_builtin": true,
      "is_modified": true
    }
  ],
  "paragraphs": [
    {
      "index": 0,
      "style_name": "Heading 1",
      "preview_text": "Introduction",
      "is_heading": true,
      "heading_level": 1,
      "page_number": 1
    }
  ],
  "headings": [
    {
      "text": "Introduction",
      "level": 1,
      "style_name": "Heading 1",
      "paragraph_index": 0,
      "page_number": 1,
      "in_table": false
    }
  ],
  "fields": [],
  "tables": []
}
```

## Plan V1 Schema

### Purpose
Represents LLM-generated execution plans with whitelisted atomic operations.

### Schema Definition
```json
{
  "$schema": "http://json-schema.org/draft-07/schema#",
  "$id": "plan.v1.json",
  "title": "Execution Plan V1",
  "description": "LLM-generated execution plan with atomic operations",
  "type": "object",
  "required": ["schema_version", "ops"],
  "properties": {
    "schema_version": {
      "type": "string",
      "const": "plan.v1"
    },
    "ops": {
      "type": "array",
      "items": {"$ref": "#/definitions/AtomicOperation"}
    }
  }
}
```

### Atomic Operation Definitions

#### Base AtomicOperation
```json
{
  "AtomicOperation": {
    "type": "object",
    "required": ["operation"],
    "properties": {
      "operation": {
        "type": "string",
        "enum": [
          "delete_section_by_heading",
          "update_toc",
          "delete_toc",
          "set_style_rule",
          "reassign_paragraphs_to_style",
          "clear_direct_formatting"
        ]
      }
    },
    "oneOf": [
      {"$ref": "#/definitions/DeleteSectionByHeading"},
      {"$ref": "#/definitions/UpdateToc"},
      {"$ref": "#/definitions/DeleteToc"},
      {"$ref": "#/definitions/SetStyleRule"},
      {"$ref": "#/definitions/ReassignParagraphsToStyle"},
      {"$ref": "#/definitions/ClearDirectFormatting"}
    ]
  }
}
```

#### DeleteSectionByHeading
```json
{
  "DeleteSectionByHeading": {
    "type": "object",
    "required": ["operation", "heading_text", "level", "match"],
    "properties": {
      "operation": {"const": "delete_section_by_heading"},
      "heading_text": {"type": "string", "minLength": 1},
      "level": {"type": "integer", "minimum": 1, "maximum": 9},
      "match": {
        "type": "string",
        "enum": ["EXACT", "CONTAINS", "REGEX"]
      },
      "case_sensitive": {"type": "boolean", "default": false},
      "occurrence_index": {"type": "integer", "minimum": 1}
    },
    "additionalProperties": false
  }
}
```

#### SetStyleRule
```json
{
  "SetStyleRule": {
    "type": "object",
    "required": ["operation", "target_style"],
    "properties": {
      "operation": {"const": "set_style_rule"},
      "target_style": {"type": "string"},
      "font_east_asian": {"type": "string"},
      "font_latin": {"type": "string"},
      "font_size_pt": {"type": "number", "minimum": 6, "maximum": 72},
      "font_bold": {"type": "boolean"},
      "font_italic": {"type": "boolean"},
      "font_color_hex": {
        "type": "string",
        "pattern": "^#[0-9A-Fa-f]{6}$"
      },
      "line_spacing_mode": {
        "type": "string",
        "enum": ["SINGLE", "MULTIPLE", "EXACTLY"]
      },
      "line_spacing_value": {"type": "number", "minimum": 0.5, "maximum": 10.0},
      "space_before_pt": {"type": "number", "minimum": 0, "maximum": 100},
      "space_after_pt": {"type": "number", "minimum": 0, "maximum": 100},
      "alignment": {
        "type": "string",
        "enum": ["LEFT", "CENTER", "RIGHT", "JUSTIFY"]
      }
    },
    "additionalProperties": false
  }
}
```

### Example Plan V1 Document
```json
{
  "schema_version": "plan.v1",
  "ops": [
    {
      "operation": "delete_section_by_heading",
      "heading_text": "Abstract",
      "level": 1,
      "match": "EXACT",
      "case_sensitive": false
    },
    {
      "operation": "delete_section_by_heading",
      "heading_text": "References",
      "level": 1,
      "match": "EXACT",
      "case_sensitive": false
    },
    {
      "operation": "update_toc"
    },
    {
      "operation": "set_style_rule",
      "target_style": "Heading 1",
      "font_east_asian": "楷体",
      "font_size_pt": 12,
      "font_bold": true,
      "line_spacing_mode": "MULTIPLE",
      "line_spacing_value": 2.0
    }
  ]
}
```

## Inventory Full V1 Schema

### Purpose
Complete document inventory with OOXML fragments and media references.

### Schema Definition
```json
{
  "$schema": "http://json-schema.org/draft-07/schema#",
  "$id": "inventory.full.v1.json",
  "title": "Document Inventory Full V1",
  "description": "Complete document inventory with OOXML fragments",
  "type": "object",
  "required": ["schema_version", "ooxml_fragments", "media_indexes"],
  "properties": {
    "schema_version": {
      "type": "string",
      "const": "inventory.full.v1"
    },
    "ooxml_fragments": {
      "type": "object",
      "patternProperties": {
        "^[a-zA-Z0-9_-]+$": {"type": "string"}
      }
    },
    "media_indexes": {
      "type": "object",
      "patternProperties": {
        "^[a-zA-Z0-9_-]+$": {"$ref": "#/definitions/MediaReference"}
      }
    },
    "content_controls": {
      "type": "array",
      "items": {"$ref": "#/definitions/ContentControlReference"}
    },
    "formulas": {
      "type": "array",
      "items": {"$ref": "#/definitions/FormulaReference"}
    },
    "charts": {
      "type": "array",
      "items": {"$ref": "#/definitions/ChartReference"}
    }
  }
}
```

### Reference Definitions

#### MediaReference
```json
{
  "MediaReference": {
    "type": "object",
    "required": ["media_id", "filename", "content_type", "size_bytes"],
    "properties": {
      "media_id": {"type": "string"},
      "filename": {"type": "string"},
      "content_type": {"type": "string"},
      "size_bytes": {"type": "integer", "minimum": 0},
      "embedded": {"type": "boolean"}
    }
  }
}
```

## Configuration Schema

### Purpose
Validate configuration files and runtime parameters.

### Schema Definition
```json
{
  "$schema": "http://json-schema.org/draft-07/schema#",
  "$id": "config.schema.json",
  "title": "AutoWord vNext Configuration",
  "description": "Configuration schema for AutoWord vNext",
  "type": "object",
  "properties": {
    "model": {
      "type": "string",
      "enum": ["gpt4", "gpt35", "claude37", "claude3"]
    },
    "temperature": {
      "type": "number",
      "minimum": 0.0,
      "maximum": 1.0
    },
    "audit_dir": {"type": "string"},
    "visible": {"type": "boolean"},
    "verbose": {"type": "boolean"},
    "monitoring_level": {
      "type": "string",
      "enum": ["basic", "detailed", "debug", "performance"]
    },
    "enable_audit_trail": {"type": "boolean"},
    "enable_rollback": {"type": "boolean"},
    "max_execution_time_seconds": {
      "type": "integer",
      "minimum": 30,
      "maximum": 3600
    },
    "localization": {
      "type": "object",
      "properties": {
        "style_aliases": {
          "type": "object",
          "patternProperties": {
            ".*": {"type": "string"}
          }
        },
        "font_fallbacks": {
          "type": "object",
          "patternProperties": {
            ".*": {
              "type": "array",
              "items": {"type": "string"}
            }
          }
        }
      }
    }
  }
}
```

## Schema Validation API

### Validation Functions

```python
from autoword.vnext.schema_validator import SchemaValidator

validator = SchemaValidator()

# Validate structure
result = validator.validate_structure(structure_data)
if not result.is_valid:
    print(f"Validation errors: {result.errors}")

# Validate plan
result = validator.validate_plan(plan_data)
if not result.is_valid:
    print(f"Plan validation failed: {result.errors}")

# Validate configuration
result = validator.validate_config(config_data)
if not result.is_valid:
    print(f"Configuration errors: {result.errors}")
```

### Validation Result
```python
@dataclass
class ValidationResult:
    is_valid: bool
    errors: List[str]
    warnings: List[str]
    schema_version: str
```

## Schema Evolution

### Version Management
- **Semantic Versioning**: Major.Minor format (e.g., v1.0, v1.1, v2.0)
- **Backward Compatibility**: Minor versions maintain compatibility
- **Migration Support**: Automatic migration between compatible versions

### Schema Updates
1. **New Fields**: Added as optional with defaults
2. **Field Changes**: Deprecated old, add new with migration
3. **Breaking Changes**: New major version with migration path

### Migration Example
```python
from autoword.vnext.schema_migrator import SchemaMigrator

migrator = SchemaMigrator()

# Migrate structure from v1.0 to v1.1
migrated_structure = migrator.migrate_structure(
    old_structure, 
    from_version="v1.0", 
    to_version="v1.1"
)
```

## Best Practices

### Schema Design
1. **Required Fields**: Mark essential fields as required
2. **Validation Rules**: Use appropriate constraints and patterns
3. **Documentation**: Include descriptions for all fields
4. **Examples**: Provide complete example documents

### Validation Strategy
1. **Early Validation**: Validate at input boundaries
2. **Comprehensive Checking**: Validate all data transformations
3. **Error Reporting**: Provide detailed, actionable error messages
4. **Performance**: Optimize validation for large documents

### Error Handling
1. **Graceful Degradation**: Continue processing when possible
2. **Clear Messages**: Provide specific validation error details
3. **Context Information**: Include field paths and values
4. **Recovery Suggestions**: Suggest fixes for common errors

This schema reference provides complete documentation for all JSON schemas used in AutoWord vNext, ensuring data integrity and enabling reliable document processing.