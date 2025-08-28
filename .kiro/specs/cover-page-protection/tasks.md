# Implementation Plan

- [x] 1. Enhance existing _set_style_rule method with style separation





  - Modify `_set_style_rule` in SimplePipeline to detect Normal/正文 style modifications
  - Add logic to create "BodyText (AutoWord)" style when Normal/正文 is targeted
  - Implement style cloning using existing Word COM operations
  - Update existing style_mappings dictionary to include body text style
  - Test style separation with existing test documents
  - _Requirements: 1.1, 1.2, 3.1, 3.5_

- [x] 2. Improve existing cover detection in _is_cover_or_toc_content method





  - Enhance keyword detection with more comprehensive cover page indicators
  - Add better style-based detection for cover elements
  - Improve text pattern recognition for academic paper covers
  - Add detection for text boxes and shapes on cover pages
  - Test enhanced detection with various document types
  - _Requirements: 4.1, 4.2, 4.3_

- [x] 3. Extend _apply_styles_to_content with paragraph reassignment





  - Add logic to reassign main content paragraphs from Normal to "BodyText (AutoWord)"
  - Integrate with existing `_is_cover_or_toc_content` filtering
  - Enhance existing paragraph iteration to handle style reassignment
  - Add logging for reassignment operations using existing logger
  - Test paragraph reassignment with existing cover protection tests
  - _Requirements: 1.3, 1.4, 3.3_



- [x] 4. Add shape and text frame processing to existing pipeline



  - Extend existing shape processing in `_apply_styles_to_content`
  - Add anchor page detection using Word COM Information property
  - Implement text frame paragraph style reassignment
  - Add filtering to skip shapes anchored to cover pages
  - Test shape processing with documents containing cover text boxes
  - _Requirements: 4.4, 4.5_

- [x] 5. Enhance existing validation with cover format checking






  - Extend `_validate_result` method to check cover page formatting
  - Add before/after comparison for cover page line spacing and fonts
  - Implement rollback trigger when cover changes are detected
  - Add cover validation logging to existing warning system
  - Test validation with documents where cover formatting should be preserved
  - _Requirements: 4.5_




- [-] 6. Create comprehensive test suite for cover protection

  - Develop test documents with various cover page scenarios
  - Create automated tests that verify cover preservation
  - Add regression tests for existing functionality
  - Test integration with existing SimplePipeline workflow
  - Document test results and known limitations
  - _Requirements: 1.5, 3.4, 4.5_