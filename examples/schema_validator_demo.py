"""
Schema Validator Demo - AutoWord vNext

This demo showcases the comprehensive JSON schema validation system for AutoWord vNext.
Demonstrates validation of structure.v1.json, plan.v1.json, and inventory.full.v1.json
with detailed error reporting and whitelist operation enforcement.
"""

import json
from datetime import datetime
from pathlib import Path

from autoword.vnext.schema_validator import (
    SchemaValidator, validate_structure, validate_plan, validate_inventory,
    validate_with_detailed_errors
)
from autoword.vnext.models import (
    StructureV1, PlanV1, InventoryFullV1, DocumentMetadata, StyleDefinition,
    ParagraphSkeleton, HeadingReference, DeleteSectionByHeading, UpdateToc,
    SetStyleRule, FontSpec, ParagraphSpec, SchemaVersion
)
from autoword.vnext.exceptions import SchemaValidationError


def demo_structure_validation():
    """Demonstrate structure.v1.json validation."""
    print("=" * 60)
    print("STRUCTURE VALIDATION DEMO")
    print("=" * 60)
    
    # Valid structure data
    valid_structure = {
        "schema_version": "structure.v1",
        "metadata": {
            "title": "AutoWord Demo Document",
            "author": "Demo User",
            "creation_time": "2024-01-15T10:00:00Z",
            "modified_time": "2024-01-15T11:30:00Z",
            "word_version": "16.0",
            "page_count": 10,
            "paragraph_count": 150,
            "word_count": 2500
        },
        "styles": [
            {
                "name": "Normal",
                "type": "paragraph",
                "font": {
                    "east_asian": "宋体",
                    "latin": "Times New Roman",
                    "size_pt": 12,
                    "bold": False,
                    "color_hex": "#000000"
                },
                "paragraph": {
                    "line_spacing_mode": "MULTIPLE",
                    "line_spacing_value": 2.0
                }
            },
            {
                "name": "Heading 1",
                "type": "paragraph",
                "font": {
                    "east_asian": "楷体",
                    "latin": "Times New Roman",
                    "size_pt": 14,
                    "bold": True,
                    "color_hex": "#000000"
                }
            }
        ],
        "paragraphs": [
            {
                "index": 0,
                "style_name": "Normal",
                "preview_text": "This is the first paragraph of the demo document.",
                "is_heading": False
            },
            {
                "index": 1,
                "style_name": "Heading 1",
                "preview_text": "Chapter 1: Introduction",
                "is_heading": True,
                "heading_level": 1
            },
            {
                "index": 2,
                "style_name": "Normal",
                "preview_text": "This chapter introduces the main concepts...",
                "is_heading": False
            }
        ],
        "headings": [
            {
                "paragraph_index": 1,
                "level": 1,
                "text": "Chapter 1: Introduction",
                "style_name": "Heading 1"
            }
        ],
        "fields": [
            {
                "paragraph_index": 0,
                "field_type": "TOC",
                "field_code": "TOC \\o \"1-3\" \\h \\z \\u",
                "result_text": "Table of Contents"
            }
        ],
        "tables": []
    }
    
    print("1. Validating VALID structure data...")
    result = validate_structure(valid_structure)
    print(f"   Valid: {result.is_valid}")
    print(f"   Errors: {len(result.errors)}")
    print(f"   Warnings: {len(result.warnings)}")
    if result.warnings:
        for warning in result.warnings:
            print(f"   Warning: {warning}")
    print()
    
    # Test with Pydantic model
    print("2. Validating with Pydantic model...")
    try:
        structure_model = StructureV1(**valid_structure)
        result = validate_structure(structure_model)
        print(f"   Valid: {result.is_valid}")
        print(f"   Model created successfully: {structure_model.metadata.title}")
    except Exception as e:
        print(f"   Error creating model: {e}")
    print()
    
    # Invalid structure data
    print("3. Validating INVALID structure data...")
    invalid_structure = valid_structure.copy()
    invalid_structure["schema_version"] = "invalid.version"  # Wrong version
    invalid_structure["paragraphs"][0]["preview_text"] = "x" * 150  # Too long
    invalid_structure["headings"][0]["paragraph_index"] = 999  # Invalid reference
    
    result = validate_structure(invalid_structure)
    print(f"   Valid: {result.is_valid}")
    print(f"   Errors: {len(result.errors)}")
    for i, error in enumerate(result.errors[:3], 1):  # Show first 3 errors
        print(f"   Error {i}: {error}")
    if len(result.errors) > 3:
        print(f"   ... and {len(result.errors) - 3} more errors")
    print()


def demo_plan_validation():
    """Demonstrate plan.v1.json validation with whitelist enforcement."""
    print("=" * 60)
    print("PLAN VALIDATION DEMO")
    print("=" * 60)
    
    # Valid plan data
    valid_plan = {
        "schema_version": "plan.v1",
        "ops": [
            {
                "operation_type": "delete_section_by_heading",
                "heading_text": "摘要",
                "level": 1,
                "match": "EXACT",
                "case_sensitive": False
            },
            {
                "operation_type": "delete_section_by_heading",
                "heading_text": "参考文献",
                "level": 1,
                "match": "CONTAINS"
            },
            {
                "operation_type": "update_toc"
            },
            {
                "operation_type": "set_style_rule",
                "target_style_name": "Heading 1",
                "font": {
                    "east_asian": "楷体",
                    "size_pt": 12,
                    "bold": True
                },
                "paragraph": {
                    "line_spacing_mode": "MULTIPLE",
                    "line_spacing_value": 2.0
                }
            },
            {
                "operation_type": "reassign_paragraphs_to_style",
                "selector": {"style_name": "Normal"},
                "target_style_name": "Body Text",
                "clear_direct_formatting": True
            }
        ]
    }
    
    print("1. Validating VALID plan data...")
    result = validate_plan(valid_plan)
    print(f"   Valid: {result.is_valid}")
    print(f"   Operations: {len(valid_plan['ops'])}")
    print(f"   Errors: {len(result.errors)}")
    print()
    
    # Test whitelist enforcement
    print("2. Testing WHITELIST enforcement...")
    invalid_plan = {
        "schema_version": "plan.v1",
        "ops": [
            {
                "operation_type": "unauthorized_operation",  # Not in whitelist
                "malicious_param": "dangerous_value"
            },
            {
                "operation_type": "string_replacement",  # Also not allowed
                "find": "old_text",
                "replace": "new_text"
            }
        ]
    }
    
    result = validate_plan(invalid_plan)
    print(f"   Valid: {result.is_valid}")
    print(f"   Errors: {len(result.errors)}")
    for i, error in enumerate(result.errors, 1):
        print(f"   Error {i}: {error}")
    print()
    
    # Test operation-specific validation
    print("3. Testing OPERATION-SPECIFIC validation...")
    invalid_operations = {
        "schema_version": "plan.v1",
        "ops": [
            {
                "operation_type": "delete_section_by_heading",
                "heading_text": "",  # Empty heading text
                "level": 10  # Invalid level (must be 1-9)
            },
            {
                "operation_type": "set_style_rule",
                "target_style_name": "",  # Empty style name
                "font": {
                    "size_pt": 100  # Invalid size (must be 1-72)
                }
            },
            {
                "operation_type": "clear_direct_formatting",
                "scope": "document"
                # Missing authorization_required: true
            }
        ]
    }
    
    result = validate_plan(invalid_operations)
    print(f"   Valid: {result.is_valid}")
    print(f"   Errors: {len(result.errors)}")
    for i, error in enumerate(result.errors, 1):
        print(f"   Error {i}: {error}")
    print()


def demo_inventory_validation():
    """Demonstrate inventory.full.v1.json validation."""
    print("=" * 60)
    print("INVENTORY VALIDATION DEMO")
    print("=" * 60)
    
    # Valid inventory data
    valid_inventory = {
        "schema_version": "inventory.full.v1",
        "ooxml_fragments": {
            "chart_1": "<c:chart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:rPr lang=\"en-US\"/><a:t>Sales Chart</a:t></a:r></a:p></c:rich></c:tx></c:title></c:chart>",
            "formula_1": "<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:sSup><m:e><m:r><m:t>E</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup><m:r><m:t>=</m:t></m:r><m:sSup><m:e><m:r><m:t>mc</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath>",
            "content_control_1": "<w:sdt xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:sdtPr><w:alias w:val=\"Demo Control\"/><w:tag w:val=\"demo\"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Demo content</w:t></w:r></w:p></w:sdtContent></w:sdt>"
        },
        "media_indexes": {
            "word/media/image1.png": {
                "media_id": "rId1",
                "content_type": "image/png",
                "file_extension": ".png",
                "size_bytes": 45678
            },
            "word/media/chart1.xml": {
                "media_id": "rId2",
                "content_type": "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
                "file_extension": ".xml",
                "size_bytes": 12345
            }
        },
        "content_controls": [
            {
                "paragraph_index": 5,
                "control_id": "cc1",
                "control_type": "richText",
                "title": "Demo Rich Text Control",
                "tag": "demo_tag"
            },
            {
                "paragraph_index": 8,
                "control_id": "cc2",
                "control_type": "dropDownList",
                "title": "Demo Dropdown"
            }
        ],
        "formulas": [
            {
                "paragraph_index": 12,
                "formula_id": "eq1",
                "formula_type": "equation",
                "latex_code": "E = mc^2"
            },
            {
                "paragraph_index": 15,
                "formula_id": "eq2",
                "formula_type": "inline_equation",
                "latex_code": "\\sum_{i=1}^{n} x_i"
            }
        ],
        "charts": [
            {
                "paragraph_index": 20,
                "chart_id": "chart1",
                "chart_type": "column",
                "title": "Sales by Quarter"
            },
            {
                "paragraph_index": 25,
                "chart_id": "chart2",
                "chart_type": "pie",
                "title": "Market Share"
            }
        ]
    }
    
    print("1. Validating VALID inventory data...")
    result = validate_inventory(valid_inventory)
    print(f"   Valid: {result.is_valid}")
    print(f"   OOXML fragments: {len(valid_inventory['ooxml_fragments'])}")
    print(f"   Media indexes: {len(valid_inventory['media_indexes'])}")
    print(f"   Content controls: {len(valid_inventory['content_controls'])}")
    print(f"   Formulas: {len(valid_inventory['formulas'])}")
    print(f"   Charts: {len(valid_inventory['charts'])}")
    print(f"   Errors: {len(result.errors)}")
    print()
    
    # Invalid inventory data
    print("2. Validating INVALID inventory data...")
    invalid_inventory = valid_inventory.copy()
    
    # Invalid OOXML fragment (non-string)
    invalid_inventory["ooxml_fragments"]["invalid"] = 12345
    
    # Invalid media reference (missing required fields)
    invalid_inventory["media_indexes"]["invalid_media"] = {
        "file_extension": ".jpg"
        # Missing media_id and content_type
    }
    
    # Invalid content control (negative paragraph_index)
    invalid_inventory["content_controls"].append({
        "paragraph_index": -1,
        "control_id": "",  # Empty ID
        "control_type": "richText"
    })
    
    result = validate_inventory(invalid_inventory)
    print(f"   Valid: {result.is_valid}")
    print(f"   Errors: {len(result.errors)}")
    for i, error in enumerate(result.errors, 1):
        print(f"   Error {i}: {error}")
    print()


def demo_detailed_error_reporting():
    """Demonstrate detailed error reporting capabilities."""
    print("=" * 60)
    print("DETAILED ERROR REPORTING DEMO")
    print("=" * 60)
    
    # Create data with multiple types of errors
    complex_invalid_data = {
        "schema_version": "structure.v1",
        "metadata": {
            "title": None,  # Should be string or null
            "page_count": -5,  # Should be non-negative
            "word_count": "invalid"  # Should be integer
        },
        "styles": [
            {
                "name": "",  # Empty name not allowed
                "type": "invalid_type",  # Not in enum
                "font": {
                    "size_pt": 100,  # Outside valid range
                    "color_hex": "not_a_color"  # Invalid hex format
                }
            }
        ],
        "paragraphs": [
            {
                "index": 0,
                "preview_text": "x" * 200,  # Too long
                "is_heading": "yes"  # Should be boolean
            },
            {
                "index": 0,  # Duplicate index
                "preview_text": "Another paragraph",
                "is_heading": False
            }
        ],
        "headings": [
            {
                "paragraph_index": 999,  # Non-existent paragraph
                "level": 0,  # Invalid level
                "text": ""  # Empty text
            }
        ],
        "fields": [],
        "tables": []
    }
    
    print("1. Validating complex INVALID data...")
    try:
        result = validate_with_detailed_errors(
            complex_invalid_data, 
            SchemaVersion.STRUCTURE_V1,
            raise_on_error=False
        )
        
        print(f"   Valid: {result.is_valid}")
        print(f"   Total errors: {len(result.errors)}")
        print(f"   Total warnings: {len(result.warnings)}")
        print()
        
        print("   Detailed errors:")
        for i, error in enumerate(result.errors[:5], 1):  # Show first 5 errors
            print(f"   {i}. {error}")
        
        if len(result.errors) > 5:
            print(f"   ... and {len(result.errors) - 5} more errors")
        
        if result.warnings:
            print("\n   Warnings:")
            for i, warning in enumerate(result.warnings, 1):
                print(f"   {i}. {warning}")
        
    except Exception as e:
        print(f"   Unexpected error: {e}")
    print()
    
    print("2. Testing raise_on_error=True...")
    try:
        validate_with_detailed_errors(
            complex_invalid_data,
            SchemaVersion.STRUCTURE_V1,
            raise_on_error=True
        )
        print("   No exception raised (unexpected)")
    except SchemaValidationError as e:
        print(f"   SchemaValidationError raised as expected")
        print(f"   Schema: {e.details.get('schema_name', 'Unknown')}")
        print(f"   Error count: {len(e.details.get('validation_errors', []))}")
        validation_errors = e.details.get('validation_errors', [])
        print(f"   First error: {validation_errors[0] if validation_errors else 'None'}")
    print()


def demo_schema_utilities():
    """Demonstrate schema validator utility functions."""
    print("=" * 60)
    print("SCHEMA UTILITIES DEMO")
    print("=" * 60)
    
    validator = SchemaValidator()
    
    print("1. Available schemas:")
    schemas = validator.list_available_schemas()
    for schema in schemas:
        print(f"   - {schema}")
    print()
    
    print("2. Schema details:")
    for schema_version in schemas:
        try:
            schema = validator.get_schema(schema_version)
            print(f"   {schema_version}:")
            print(f"     Title: {schema.get('title', 'N/A')}")
            print(f"     Description: {schema.get('description', 'N/A')[:60]}...")
            print(f"     Required fields: {len(schema.get('required', []))}")
        except Exception as e:
            print(f"   Error loading {schema_version}: {e}")
    print()
    
    print("3. Pydantic model validation:")
    # Test direct Pydantic validation
    valid_metadata = {
        "title": "Test Document",
        "author": "Test Author",
        "page_count": 10
    }
    
    result = validator.validate_pydantic_model(valid_metadata, DocumentMetadata)
    print(f"   Valid metadata: {result.is_valid}")
    
    invalid_metadata = {
        "page_count": -1,  # Invalid
        "word_count": "not_a_number"  # Invalid type
    }
    
    result = validator.validate_pydantic_model(invalid_metadata, DocumentMetadata)
    print(f"   Invalid metadata: {result.is_valid}")
    print(f"   Errors: {len(result.errors)}")
    for error in result.errors:
        print(f"     - {error}")
    print()


def main():
    """Run all schema validation demos."""
    print("AutoWord vNext - Schema Validation System Demo")
    print("=" * 60)
    print("This demo showcases comprehensive JSON schema validation")
    print("for structure.v1, plan.v1, and inventory.full.v1 schemas")
    print("with detailed error reporting and whitelist enforcement.")
    print()
    
    try:
        demo_structure_validation()
        demo_plan_validation()
        demo_inventory_validation()
        demo_detailed_error_reporting()
        demo_schema_utilities()
        
        print("=" * 60)
        print("DEMO COMPLETED SUCCESSFULLY")
        print("=" * 60)
        print("Key features demonstrated:")
        print("✓ Structure.v1.json validation with field constraints")
        print("✓ Plan.v1.json validation with whitelist enforcement")
        print("✓ Inventory.full.v1.json validation with OOXML storage")
        print("✓ Detailed error reporting with field-level information")
        print("✓ Pydantic model integration and validation")
        print("✓ Schema utility functions and introspection")
        print("✓ Warning generation for non-critical issues")
        print("✓ Exception handling with SchemaValidationError")
        
    except Exception as e:
        print(f"Demo failed with error: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()