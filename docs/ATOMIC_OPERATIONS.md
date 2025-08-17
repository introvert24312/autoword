# AutoWord vNext - Atomic Operations Reference

## Overview

Atomic operations are the fundamental building blocks of document modifications in AutoWord vNext. Each operation is designed to be safe, predictable, and reversible through the rollback system.

## Operation Categories

### 1. Section Management Operations
- `delete_section_by_heading`: Remove document sections
- `move_section`: Relocate document sections (future)

### 2. Table of Contents Operations
- `update_toc`: Refresh TOC fields and page numbers
- `delete_toc`: Remove table of contents

### 3. Style Management Operations
- `set_style_rule`: Modify style definitions
- `reassign_paragraphs_to_style`: Change paragraph styles
- `clear_direct_formatting`: Remove direct formatting

### 4. Content Operations
- `insert_page_break`: Add page breaks (future)
- `update_fields`: Refresh document fields (future)

## Detailed Operation Reference

### delete_section_by_heading

**Purpose**: Delete a document section from a specified heading to the next heading of the same or higher level.

**Parameters**:
```json
{
  "op_type": "delete_section_by_heading",
  "heading_text": "string",
  "level": "number (1-9)",
  "match": "EXACT|CONTAINS|REGEX",
  "case_sensitive": "boolean",
  "occurrence_index": "number (optional)"
}
```

**Parameter Details**:
- `heading_text`: The text to match in headings
- `level`: Heading level (1-9, where 1 is highest)
- `match`: Matching strategy
  - `EXACT`: Exact text match
  - `CONTAINS`: Heading contains the text
  - `REGEX`: Regular expression matching
- `case_sensitive`: Whether matching is case-sensitive
- `occurrence_index`: Which occurrence to delete (1-based, optional)

**Examples**:

1. **Delete Abstract Section**:
```json
{
  "op_type": "delete_section_by_heading",
  "heading_text": "摘要",
  "level": 1,
  "match": "EXACT",
  "case_sensitive": false
}
```

2. **Delete Second Introduction**:
```json
{
  "op_type": "delete_section_by_heading",
  "heading_text": "Introduction",
  "level": 1,
  "match": "EXACT",
  "case_sensitive": false,
  "occurrence_index": 2
}
```

3. **Delete Any Section Containing "Appendix"**:
```json
{
  "op_type": "delete_section_by_heading",
  "heading_text": "Appendix",
  "level": 1,
  "match": "CONTAINS",
  "case_sensitive": false
}
```

**Behavior**:
- Deletes from the matched heading to the next heading of same or higher level
- If no next heading found, deletes to end of document
- Handles headings within tables by deleting table rows
- Logs NOOP if heading not found

**Edge Cases**:
- Heading in table: Deletes entire table row
- Last section: Deletes to document end
- Duplicate headings: Uses occurrence_index or deletes first
- Nested headings: Preserves section boundary logic

---

### update_toc

**Purpose**: Refresh all table of contents fields and update page numbers.

**Parameters**:
```json
{
  "op_type": "update_toc"
}
```

**Examples**:

1. **Standard TOC Update**:
```json
{
  "op_type": "update_toc"
}
```

**Behavior**:
- Calls `Fields.Update()` on all TOC fields
- Updates page numbers and heading text
- Maintains TOC formatting and styles
- Logs NOOP if no TOC found

**Edge Cases**:
- No TOC present: Logs NOOP, continues processing
- Multiple TOCs: Updates all found TOCs
- Corrupted TOC: Attempts repair, may fail gracefully

---

### delete_toc

**Purpose**: Remove table of contents from document.

**Parameters**:
```json
{
  "op_type": "delete_toc",
  "mode": "first|all|by_index"
}
```

**Parameter Details**:
- `mode`: Deletion strategy
  - `first`: Delete first TOC found
  - `all`: Delete all TOCs in document
  - `by_index`: Delete specific TOC by index (requires `toc_index`)

**Examples**:

1. **Delete First TOC**:
```json
{
  "op_type": "delete_toc",
  "mode": "first"
}
```

2. **Delete All TOCs**:
```json
{
  "op_type": "delete_toc",
  "mode": "all"
}
```

**Behavior**:
- Removes TOC field and associated text
- Preserves surrounding formatting
- Logs NOOP if no TOC found

---

### set_style_rule

**Purpose**: Modify style definitions including fonts, sizes, and formatting.

**Parameters**:
```json
{
  "op_type": "set_style_rule",
  "target_style_name": "string",
  "font": {
    "east_asian": "string",
    "latin": "string",
    "size_pt": "number",
    "bold": "boolean",
    "italic": "boolean",
    "color_hex": "string"
  },
  "paragraph": {
    "line_spacing_mode": "SINGLE|MULTIPLE|EXACTLY",
    "line_spacing_value": "number",
    "space_before_pt": "number",
    "space_after_pt": "number",
    "alignment": "LEFT|CENTER|RIGHT|JUSTIFY"
  }
}
```

**Parameter Details**:
- `target_style_name`: Name of style to modify
- `font`: Font specifications
  - `east_asian`: East Asian font name (for Chinese/Japanese/Korean)
  - `latin`: Latin font name (for English/European)
  - `size_pt`: Font size in points
  - `bold`: Bold formatting
  - `italic`: Italic formatting
  - `color_hex`: Font color in hex format
- `paragraph`: Paragraph formatting
  - `line_spacing_mode`: Line spacing type
  - `line_spacing_value`: Spacing value (depends on mode)
  - `space_before_pt`: Space before paragraph in points
  - `space_after_pt`: Space after paragraph in points
  - `alignment`: Text alignment

**Examples**:

1. **Set Heading 1 Style**:
```json
{
  "op_type": "set_style_rule",
  "target_style_name": "Heading 1",
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

2. **Set Normal Style**:
```json
{
  "op_type": "set_style_rule",
  "target_style_name": "Normal",
  "font": {
    "east_asian": "宋体",
    "latin": "Times New Roman",
    "size_pt": 12,
    "bold": false
  },
  "paragraph": {
    "line_spacing_mode": "MULTIPLE",
    "line_spacing_value": 2.0
  }
}
```

**Behavior**:
- Modifies style object in document's style collection
- Applies font fallback if specified font unavailable
- Uses localization manager for style name resolution
- All paragraphs using the style are automatically updated

**Font Fallback**:
- East Asian fonts: 楷体 → 楷体_GB2312 → STKaiti
- Latin fonts: Times New Roman → Times → serif
- Logs fallback usage to warnings.log

---

### reassign_paragraphs_to_style

**Purpose**: Change the style of specific paragraphs matching criteria.

**Parameters**:
```json
{
  "op_type": "reassign_paragraphs_to_style",
  "selector": {
    "current_style": "string",
    "text_contains": "string",
    "paragraph_indexes": ["array of numbers"],
    "heading_level": "number"
  },
  "target_style": "string",
  "clear_direct_formatting": "boolean"
}
```

**Parameter Details**:
- `selector`: Criteria for selecting paragraphs
  - `current_style`: Match paragraphs with this style
  - `text_contains`: Match paragraphs containing this text
  - `paragraph_indexes`: Specific paragraph indexes
  - `heading_level`: Match headings of this level
- `target_style`: Style to apply to selected paragraphs
- `clear_direct_formatting`: Whether to remove direct formatting

**Examples**:

1. **Convert All Heading 3 to Heading 2**:
```json
{
  "op_type": "reassign_paragraphs_to_style",
  "selector": {
    "current_style": "Heading 3"
  },
  "target_style": "Heading 2",
  "clear_direct_formatting": true
}
```

2. **Style Specific Paragraphs**:
```json
{
  "op_type": "reassign_paragraphs_to_style",
  "selector": {
    "paragraph_indexes": [5, 7, 12]
  },
  "target_style": "Normal",
  "clear_direct_formatting": false
}
```

**Behavior**:
- Selects paragraphs based on criteria
- Applies target style to selected paragraphs
- Optionally clears direct formatting
- Logs number of paragraphs modified

---

### clear_direct_formatting

**Purpose**: Remove direct formatting from specified document ranges.

**Parameters**:
```json
{
  "op_type": "clear_direct_formatting",
  "scope": "document|selection|style",
  "range_spec": {
    "start_paragraph": "number",
    "end_paragraph": "number",
    "style_name": "string"
  },
  "authorization": "EXPLICIT_USER_REQUEST"
}
```

**Parameter Details**:
- `scope`: Range of formatting to clear
  - `document`: Entire document
  - `selection`: Specific paragraph range
  - `style`: All paragraphs with specific style
- `range_spec`: Range specification (depends on scope)
- `authorization`: Required explicit authorization

**Examples**:

1. **Clear Document Formatting**:
```json
{
  "op_type": "clear_direct_formatting",
  "scope": "document",
  "authorization": "EXPLICIT_USER_REQUEST"
}
```

2. **Clear Range Formatting**:
```json
{
  "op_type": "clear_direct_formatting",
  "scope": "selection",
  "range_spec": {
    "start_paragraph": 10,
    "end_paragraph": 20
  },
  "authorization": "EXPLICIT_USER_REQUEST"
}
```

**Behavior**:
- Removes bold, italic, underline, color, font changes
- Preserves style-based formatting
- Requires explicit authorization for safety
- Logs range and number of paragraphs affected

**Security Note**: This operation requires explicit authorization due to its potentially destructive nature.

## Operation Validation

### Schema Validation

All operations must pass JSON schema validation:

```python
from autoword.vnext.planner import SchemaValidator

validator = SchemaValidator()
result = validator.validate_operation(operation_json)

if not result.is_valid:
    print(f"Validation errors: {result.errors}")
```

### Whitelist Enforcement

Only predefined operations are allowed:

```python
WHITELISTED_OPERATIONS = [
    "delete_section_by_heading",
    "update_toc",
    "delete_toc", 
    "set_style_rule",
    "reassign_paragraphs_to_style",
    "clear_direct_formatting"
]
```

### Parameter Validation

Each operation validates its parameters:

- String parameters: Non-empty, reasonable length
- Numeric parameters: Valid ranges (e.g., level 1-9)
- Enum parameters: Valid values only
- Optional parameters: Proper defaults

## Error Handling

### Operation Failures

```python
class OperationError(VNextError):
    def __init__(self, operation: str, reason: str, context: dict):
        self.operation = operation
        self.reason = reason
        self.context = context
```

### Common Error Scenarios

1. **Target Not Found**: Heading/style doesn't exist → NOOP logged
2. **Invalid Parameters**: Schema validation fails → Operation rejected
3. **COM Errors**: Word automation fails → Exception with rollback
4. **Authorization Missing**: Security check fails → Operation rejected

### Recovery Strategies

- **NOOP Logging**: Non-critical failures logged, processing continues
- **Operation Retry**: Transient COM errors retried with backoff
- **Graceful Degradation**: Font fallbacks, style approximations
- **Complete Rollback**: Critical failures trigger document restoration

## Performance Considerations

### Batch Operations

Similar operations are batched when possible:

```python
# Multiple style changes batched together
operations = [
    set_style_rule("Heading 1", ...),
    set_style_rule("Heading 2", ...),
    set_style_rule("Normal", ...)
]
# Executed as single Word COM session
```

### Memory Management

- Operations process incrementally
- Large documents handled in chunks
- COM objects properly disposed
- Memory usage monitored

### Optimization Tips

1. **Group Similar Operations**: Batch style changes together
2. **Minimize COM Calls**: Cache document structure
3. **Use Specific Selectors**: Avoid broad document scans
4. **Clear Formatting Sparingly**: Expensive operation

## Testing Operations

### Unit Testing

```python
def test_delete_section_operation():
    operation = {
        "op_type": "delete_section_by_heading",
        "heading_text": "Test Section",
        "level": 1,
        "match": "EXACT",
        "case_sensitive": False
    }
    
    result = executor.execute_operation(operation, mock_document)
    assert result.success
    assert "Test Section" not in result.final_structure.headings
```

### Integration Testing

```python
def test_complete_operation_sequence():
    operations = [
        delete_section_by_heading("Abstract", 1),
        update_toc(),
        set_style_rule("Heading 1", font_spec)
    ]
    
    result = pipeline.execute_operations(operations, test_document)
    assert result.status == "SUCCESS"
```

## Best Practices

### 1. Operation Design

- **Atomic**: Each operation does one thing well
- **Idempotent**: Safe to run multiple times
- **Reversible**: Can be undone through rollback
- **Validated**: All parameters checked

### 2. Error Handling

- **Graceful Degradation**: Continue when possible
- **Clear Logging**: Detailed error context
- **User Feedback**: Meaningful error messages
- **Recovery Options**: Rollback and retry strategies

### 3. Performance

- **Batch Similar Operations**: Reduce COM overhead
- **Cache Structure**: Minimize document parsing
- **Validate Early**: Catch errors before execution
- **Monitor Resources**: Track memory and time

### 4. Security

- **Whitelist Only**: No arbitrary operations
- **Parameter Validation**: Sanitize all inputs
- **Authorization Checks**: Explicit approval for destructive operations
- **Audit Trail**: Log all operations and outcomes

This reference provides the complete specification for all atomic operations available in AutoWord vNext. Each operation is designed to be safe, predictable, and fully auditable.