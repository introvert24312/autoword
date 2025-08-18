# Scenario 4: Headings in Tables

## Test Objective
Verify that the pipeline can properly handle and delete sections where headings are located within table cells.

## Document Characteristics
- Contains headings formatted within table cells
- Tests complex document structure with embedded headings
- Verifies section boundary detection across table structures
- May include tables with merged cells containing headings

## Expected Behavior
1. Extract structure correctly identifying headings within tables
2. Map table cell references to paragraph indexes
3. Generate standard deletion plan (no special table handling needed)
4. Execute deletion including entire sections with table-embedded headings
5. Preserve table structure integrity where possible
6. Update TOC to reflect structural changes

## Validation Points
- Headings within tables are correctly identified
- Section boundaries properly calculated across table structures
- Tables containing target headings are included in deletion
- Remaining table structures maintain integrity
- TOC reflects accurate document structure

## Edge Cases Covered
- Headings in table cells
- Section boundaries spanning table structures
- Table cell paragraph index mapping
- Complex table layouts with merged cells
- Preservation of non-target table content