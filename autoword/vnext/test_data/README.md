# AutoWord vNext Test Data Package

This directory contains sample input/output packages for all 7 Definition of Done (DoD) validation scenarios. Each scenario demonstrates specific capabilities and edge cases of the vNext pipeline.

## Directory Structure

```
test_data/
├── README.md                           # This file
├── scenario_1_normal_paper/            # Normal paper processing
├── scenario_2_no_toc_document/         # Document without TOC
├── scenario_3_duplicate_headings/      # Duplicate heading handling
├── scenario_4_headings_in_tables/      # Headings within table cells
├── scenario_5_missing_fonts/           # Font fallback handling
├── scenario_6_complex_objects/         # Complex document elements
└── scenario_7_revision_tracking/       # Track changes handling
```

## Scenario Descriptions

### Scenario 1: Normal Paper Processing
**Purpose**: Standard document processing workflow
**Operations**: Delete 摘要/参考文献 sections, update TOC, apply style rules
**Expected**: Clean document with proper formatting and updated TOC

### Scenario 2: No-TOC Document Handling
**Purpose**: Handle documents without Table of Contents
**Operations**: NOOP update_toc operation, section deletion, style application
**Expected**: Successful processing with NOOP warning logged

### Scenario 3: Duplicate Heading Handling
**Purpose**: Handle multiple headings with same text
**Operations**: Delete specific occurrence using occurrence_index
**Expected**: Correct heading deleted, others preserved

### Scenario 4: Headings in Tables
**Purpose**: Delete sections where headings are in table cells
**Operations**: Section deletion with table-embedded headings
**Expected**: Proper section boundaries identified and deleted

### Scenario 5: Missing Font Handling
**Purpose**: Font fallback chain and warnings
**Operations**: Style application with unavailable fonts
**Expected**: Fallback fonts applied, warnings logged

### Scenario 6: Complex Objects Preservation
**Purpose**: Handle formulas, content controls, charts
**Operations**: Document modification preserving complex objects
**Expected**: Objects preserved in inventory, structure maintained

### Scenario 7: Revision Tracking Handling
**Purpose**: Handle documents with track changes
**Operations**: Process with revision handling strategy
**Expected**: Revisions handled according to configuration

## File Naming Convention

Each scenario directory contains:
- `input.docx` - Original document
- `expected_output.docx` - Expected result after processing
- `input_structure.v1.json` - Expected structure extraction
- `expected_plan.v1.json` - Expected LLM-generated plan
- `expected_inventory.full.v1.json` - Expected inventory extraction
- `user_intent.txt` - User intent for planning
- `validation_assertions.json` - Expected validation results
- `expected_warnings.log` - Expected warning messages
- `scenario_notes.md` - Detailed scenario documentation

## Usage

### For Testing
```python
from autoword.vnext.pipeline import VNextPipeline

# Load test scenario
scenario_dir = "test_data/scenario_1_normal_paper"
pipeline = VNextPipeline()

# Run pipeline
result = pipeline.process_document(
    f"{scenario_dir}/input.docx",
    open(f"{scenario_dir}/user_intent.txt").read()
)

# Validate against expected outputs
assert result.status == "SUCCESS"
# Compare with expected_output.docx
```

### For Development
Use these scenarios to:
1. Validate new features against known good outputs
2. Test edge cases and error handling
3. Benchmark performance improvements
4. Verify cross-platform compatibility

## Maintenance

When updating the pipeline:
1. Run all scenarios to ensure no regressions
2. Update expected outputs if behavior changes intentionally
3. Add new scenarios for new features or discovered edge cases
4. Keep documentation synchronized with actual test data