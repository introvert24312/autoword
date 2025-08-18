# AutoWord vNext Atomic Operations Reference

## Overview

AutoWord vNext executes document modifications through a set of whitelisted atomic operations. These operations are the only ways the system can modify documents, ensuring predictable and safe transformations. Each operation is designed to be:

- **Atomic**: Complete successfully or fail without partial changes
- **Reversible**: Can be undone through rollback mechanisms
- **Auditable**: All parameters and results are logged
- **Safe**: Cannot corrupt documents or cause data loss

## Operation Categories

### 1. Section Management Operations
- `delete_section_by_heading`: Remove document sections
- `move_section`: Relocate document sections (future)

### 2. Table of Contents Operations
- `update_toc`: Refresh table of contents
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

**Purpose**: Remove a section from the document based on heading text matching.

**Parameters**:
```json
{
  "operation": "delete_section_by_heading",
  "heading_text": "Abstract",
  "level": 1,
  "match": "EXACT",
  "case_sensitive": false,
  "occurrence_index": null
}
```

**Parameter Details**:
- `heading_text` (string, required): Text to match in headings
- `level` (integer, 1-9, required): Heading level to search
- `match` (string, required): Matching mode
  - `"EXACT"`: Exact text match
  - `"CONTAINS"`: Heading contains the text
  - `"REGEX"`: Regular expression matching
- `case_sensitive` (boolean, optional, default: false): Case sensitivity for matching
- `occurrence_index` (integer, optional): Which occurrence to delete (1-based, null for all)

**Behavior**:
- Finds headings at the specified level matching the criteria
- Deletes content from the matching heading to the next heading of the same or higher level
- If heading is in a table, deletes the entire table row
- Logs NOOP if no matching headings found

**Examples**:

Remove abstract section:
```json
{
  "operation": "delete_section_by_heading",
  "heading_text": "Abstract",
  "level": 1,
  "match": "EXACT",
  "case_sensitive": false
}
```

Remove all "References" sections:
```json
{
  "operation": "delete_section_by_heading",
  "heading_text": "References",
  "level": 1,
  "match": "CONTAINS",
  "case_sensitive": false
}
```

Remove second occurrence of "Introduction":
```json
{
  "operation": "delete_section_by_heading",
  "heading_text": "Introduction",
  "level": 2,
  "match": "EXACT",
  "case_sensitive": false,
  "occurrence_index": 2
}
```

**Error Conditions**:
- Invalid level (not 1-9): Operation rejected
- Invalid match mode: Operation rejected
- Negative occurrence_index: Operation rejected

### update_toc

**Purpose**: Refresh all table of contents fields in the document.

**Parameters**:
```json
{
  "operation": "update_toc"
}
```

**Parameter Details**: None required.

**Behavior**:
- Finds all TOC fields in the document
- Updates field content and page numbers
- Forces document repagination
- Updates document modification timestamp

**Examples**:

Basic TOC update:
```json
{
  "operation": "update_toc"
}
```

**Error Conditions**:
- No TOC found: Logs NOOP, continues processing
- TOC update fails: Operation fails, triggers rollback

### delete_toc

**Purpose**: Remove table of contents from the document.

**Parameters**:
```json
{
  "operation": "delete_toc",
  "mode": "ALL"
}
```

**Parameter Details**:
- `mode` (string, required): Deletion mode
  - `"ALL"`: Delete all TOC instances
  - `"FIRST"`: Delete only the first TOC
  - `"LAST"`: Delete only the last TOC

**Behavior**:
- Finds TOC fields based on mode
- Removes the entire TOC including surrounding paragraphs
- Preserves document structure around deleted TOCs

**Examples**:

Remove all TOCs:
```json
{
  "operation": "delete_toc",
  "mode": "ALL"
}
```

Remove only first TOC:
```json
{
  "operation": "delete_toc",
  "mode": "FIRST"
}
```

**Error Conditions**:
- No TOC found: Logs NOOP, continues processing
- Invalid mode: Operation rejected

### set_style_rule

**Purpose**: Modify style object definitions in the document.

**Parameters**:
```json
{
  "operation": "set_style_rule",
  "target_style": "Heading 1",
  "font_east_asian": "楷体",
  "font_latin": "Times New Roman",
  "font_size_pt": 12,
  "font_bold": true,
  "font_italic": false,
  "font_color_hex": "#000000",
  "line_spacing_mode": "MULTIPLE",
  "line_spacing_value": 2.0,
  "space_before_pt": 12,
  "space_after_pt": 6,
  "alignment": "LEFT"
}
```

**Parameter Details**:
- `target_style` (string, required): Name of style to modify
- `font_east_asian` (string, optional): East Asian font name
- `font_latin` (string, optional): Latin font name
- `font_size_pt` (number, optional): Font size in points
- `font_bold` (boolean, optional): Bold formatting
- `font_italic` (boolean, optional): Italic formatting
- `font_color_hex` (string, optional): Font color in hex format
- `line_spacing_mode` (string, optional): Line spacing mode
  - `"SINGLE"`: Single spacing
  - `"MULTIPLE"`: Multiple of single spacing
  - `"EXACTLY"`: Exact spacing in points
- `line_spacing_value` (number, optional): Spacing value (depends on mode)
- `space_before_pt` (number, optional): Space before paragraph in points
- `space_after_pt` (number, optional): Space after paragraph in points
- `alignment` (string, optional): Paragraph alignment
  - `"LEFT"`, `"CENTER"`, `"RIGHT"`, `"JUSTIFY"`

**Behavior**:
- Locates the named style in the document
- Applies font fallback chains for East Asian fonts
- Modifies only specified properties (others unchanged)
- Logs warnings for unavailable fonts

**Examples**:

Set Heading 1 style:
```json
{
  "operation": "set_style_rule",
  "target_style": "Heading 1",
  "font_east_asian": "楷体",
  "font_size_pt": 12,
  "font_bold": true,
  "line_spacing_mode": "MULTIPLE",
  "line_spacing_value": 2.0
}
```

Set Normal style with specific formatting:
```json
{
  "operation": "set_style_rule",
  "target_style": "Normal",
  "font_east_asian": "宋体",
  "font_latin": "Times New Roman",
  "font_size_pt": 12,
  "line_spacing_mode": "MULTIPLE",
  "line_spacing_value": 2.0,
  "alignment": "JUSTIFY"
}
```

**Error Conditions**:
- Style not found: Logs NOOP, continues processing
- Invalid font size: Operation rejected
- Invalid spacing values: Operation rejected
- Invalid color format: Operation rejected

### reassign_paragraphs_to_style

**Purpose**: Change the style assignment of paragraphs matching specific criteria.

**Parameters**:
```json
{
  "operation": "reassign_paragraphs_to_style",
  "selector": {
    "current_style": "Normal",
    "text_contains": "Chapter",
    "level": null,
    "position": "starts_with"
  },
  "target_style": "Heading 1",
  "clear_direct_formatting": true
}
```

**Parameter Details**:
- `selector` (object, required): Criteria for selecting paragraphs
  - `current_style` (string, optional): Current style name to match
  - `text_contains` (string, optional): Text content to match
  - `level` (integer, optional): Heading level to match
  - `position` (string, optional): Text position matching
    - `"starts_with"`: Text at beginning of paragraph
    - `"ends_with"`: Text at end of paragraph
    - `"contains"`: Text anywhere in paragraph
- `target_style` (string, required): Style to assign to matching paragraphs
- `clear_direct_formatting` (boolean, optional, default: false): Remove direct formatting

**Behavior**:
- Searches document for paragraphs matching all selector criteria
- Changes style assignment for matching paragraphs
- Optionally removes direct formatting overrides
- Preserves paragraph content and structure

**Examples**:

Convert Normal paragraphs starting with "Chapter" to Heading 1:
```json
{
  "operation": "reassign_paragraphs_to_style",
  "selector": {
    "current_style": "Normal",
    "text_contains": "Chapter",
    "position": "starts_with"
  },
  "target_style": "Heading 1",
  "clear_direct_formatting": true
}
```

Convert all Heading 2 paragraphs to Heading 3:
```json
{
  "operation": "reassign_paragraphs_to_style",
  "selector": {
    "current_style": "Heading 2"
  },
  "target_style": "Heading 3",
  "clear_direct_formatting": false
}
```

**Error Conditions**:
- Target style not found: Operation rejected
- Invalid selector criteria: Operation rejected
- No matching paragraphs: Logs NOOP, continues processing

### clear_direct_formatting

**Purpose**: Remove direct formatting overrides from specified document ranges.

**Parameters**:
```json
{
  "operation": "clear_direct_formatting",
  "scope": "DOCUMENT",
  "range_spec": {
    "start_paragraph": 1,
    "end_paragraph": 10
  },
  "authorization": "EXPLICIT_USER_REQUEST"
}
```

**Parameter Details**:
- `scope` (string, required): Scope of formatting removal
  - `"DOCUMENT"`: Entire document
  - `"SELECTION"`: Specified range
  - `"STYLE"`: All paragraphs with specific style
- `range_spec` (object, optional): Range specification for SELECTION scope
  - `start_paragraph` (integer): Starting paragraph index
  - `end_paragraph` (integer): Ending paragraph index
- `authorization` (string, required): Explicit authorization token
  - Must be `"EXPLICIT_USER_REQUEST"` for security

**Behavior**:
- Removes direct formatting (bold, italic, font changes, etc.)
- Preserves style-based formatting
- Requires explicit authorization due to potential data loss
- Cannot be undone except through document rollback

**Examples**:

Clear all direct formatting in document:
```json
{
  "operation": "clear_direct_formatting",
  "scope": "DOCUMENT",
  "authorization": "EXPLICIT_USER_REQUEST"
}
```

Clear formatting in specific range:
```json
{
  "operation": "clear_direct_formatting",
  "scope": "SELECTION",
  "range_spec": {
    "start_paragraph": 5,
    "end_paragraph": 15
  },
  "authorization": "EXPLICIT_USER_REQUEST"
}
```

**Error Conditions**:
- Missing authorization: Operation rejected
- Invalid scope: Operation rejected
- Invalid range specification: Operation rejected

## Operation Execution Model

### Execution Order

Operations are executed in the order specified in the plan:
1. Each operation is validated before execution
2. Operations are executed atomically (all-or-nothing)
3. Failed operations trigger immediate rollback
4. Successful operations are logged in audit trail

### Error Handling

**Operation Validation**:
- Schema validation against operation definitions
- Parameter range and type checking
- Security authorization verification

**Execution Errors**:
- COM automation failures
- Document corruption detection
- Resource exhaustion handling

**Recovery Mechanisms**:
- Automatic rollback on any failure
- Original document restoration
- Complete audit trail preservation

### Localization Support

**Style Name Resolution**:
- English names tried first: "Heading 1"
- Chinese fallbacks: "标题 1"
- Custom aliases from configuration

**Font Fallback Chains**:
- Primary font: "楷体"
- Fallback 1: "楷体_GB2312"
- Fallback 2: "STKaiti"
- System default if all fail

**Example Localization**:
```json
{
  "operation": "set_style_rule",
  "target_style": "Heading 1",  // Will try "标题 1" if not found
  "font_east_asian": "楷体"      // Will use fallback chain
}
```

## Security and Constraints

### Whitelist Enforcement

Only operations listed in this document are allowed:
- Unknown operations are rejected during plan validation
- Operation parameters are strictly validated
- No arbitrary code execution possible

### Safety Mechanisms

**Content Protection**:
- No string replacement on document content
- No direct text manipulation
- All changes through Word object model

**Authorization Requirements**:
- Destructive operations require explicit authorization
- Clear audit trail for all modifications
- Automatic rollback on validation failure

### Runtime Constraints

**Memory Limits**:
- Large document handling with streaming
- COM object lifecycle management
- Garbage collection optimization

**Time Limits**:
- Configurable operation timeouts
- Progress monitoring and cancellation
- Resource cleanup on timeout

## Best Practices

### Operation Design

**Atomic Operations**:
- Each operation should be independently executable
- Operations should not depend on execution order
- Failed operations should not leave partial changes

**Parameter Validation**:
- Validate all parameters before execution
- Provide clear error messages for invalid parameters
- Use sensible defaults where appropriate

### Error Recovery

**Graceful Degradation**:
- Log NOOP for operations that find no targets
- Continue processing after non-critical failures
- Provide detailed error context in audit trails

**Rollback Strategy**:
- Preserve original document before any modifications
- Implement complete rollback on any validation failure
- Maintain audit trail even for failed operations

### Performance Optimization

**Efficient Execution**:
- Batch similar operations when possible
- Minimize COM round-trips
- Cache document structure during processing

**Resource Management**:
- Proper COM object disposal
- Memory usage monitoring
- Timeout handling for long operations

This atomic operations reference provides the complete specification for all document modification capabilities in AutoWord vNext, ensuring safe, predictable, and auditable document processing.