# Scenario 3: Duplicate Headings

## Test Objective
Verify that the pipeline can handle documents with duplicate heading text at different levels and correctly target specific occurrences using occurrence_index.

## Document Characteristics
- Contains multiple headings with same text ("摘要", "参考文献") at different levels
- Tests occurrence_index parameter functionality
- Verifies precise heading targeting in complex document structures

## Expected Behavior
1. Extract structure correctly identifying all heading occurrences
2. Generate plan with occurrence_index specification for duplicate headings
3. Execute deletion targeting only the specified occurrence
4. Validate that correct sections were deleted while preserving others
5. Update TOC to reflect changes

## Validation Points
- Only first occurrence of each target heading is deleted
- Other occurrences with same text remain intact
- Document structure integrity maintained
- TOC updated correctly

## Edge Cases Covered
- Duplicate heading text at same level
- Duplicate heading text at different levels  
- Occurrence indexing (0-based)
- Selective deletion with occurrence specification