# Scenario 1: Normal Paper Processing

## Overview
This scenario tests the standard document processing workflow for a typical academic paper with abstract and references sections that need to be removed.

## Input Document Characteristics
- Contains "摘要" section at level 1
- Contains "参考文献" section at level 1
- Has existing Table of Contents
- Uses mixed heading styles that need standardization
- Standard paragraph formatting throughout

## User Intent
Delete abstract and references sections, update TOC, and standardize heading styles according to Chinese academic paper format.

## Expected Operations
1. `delete_section_by_heading` for "摘要"
2. `delete_section_by_heading` for "参考文献"
3. `set_style_rule` for Heading 1 (楷体, 12pt, bold)
4. `set_style_rule` for Heading 2 (宋体, 12pt, bold)
5. `set_style_rule` for Normal (宋体, 12pt, regular)
6. `update_toc` to refresh page numbers

## Validation Assertions
- Chapter assertion: No "摘要" or "参考文献" at level 1
- Style assertions: H1, H2, Normal styles match specifications
- TOC assertion: TOC items match heading tree with correct page numbers
- Pagination assertion: Fields updated, modified time changed

## Expected Warnings
None - this is a clean processing scenario.

## Success Criteria
- Status: SUCCESS
- All sections properly deleted
- Styles applied correctly
- TOC updated with accurate page numbers
- No validation failures
- Clean audit trail generated