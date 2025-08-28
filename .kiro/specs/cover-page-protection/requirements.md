# Requirements Document

## Introduction

The Cover Page Protection feature enhances the existing AutoWord vNext SimplePipeline to prevent cover page formatting corruption during line spacing modifications. The current system already has basic cover detection and section break insertion, but needs enhancement to implement style separation and targeted formatting. The solution builds on the existing `_is_cover_or_toc_content` and `_apply_styles_to_content` methods to create a more robust protection mechanism.

## Requirements

### Requirement 1

**User Story:** As a document processor, I want to enhance the existing SimplePipeline with style separation, so that Normal/正文 style changes don't affect cover pages.

#### Acceptance Criteria

1. WHEN executing set_style_rule operations on Normal/正文 THEN the system SHALL create "BodyText (AutoWord)" style instead
2. WHEN creating body text styles THEN the system SHALL clone from Normal/正文 and apply modifications to the clone only
3. WHEN reassigning paragraphs THEN the system SHALL use existing `_is_cover_or_toc_content` logic to skip cover paragraphs
4. WHEN preserving cover pages THEN the system SHALL leverage existing `first_content_index` detection
5. WHEN cloning styles THEN the system SHALL integrate with existing Word COM operations in `_set_style_rule`

### Requirement 2

**User Story:** As a document processor, I want to enhance the existing `_apply_styles_to_content` method with better filtering, so that cover protection is more reliable.

#### Acceptance Criteria

1. WHEN applying styles to content THEN the system SHALL use enhanced `_is_cover_or_toc_content` logic
2. WHEN filtering paragraphs THEN the system SHALL extend existing page number and content analysis
3. WHEN detecting cover content THEN the system SHALL improve existing keyword detection and style analysis
4. WHEN processing shapes THEN the system SHALL add anchor page checking to existing shape processing
5. WHEN handling text frames THEN the system SHALL integrate with existing COM operations

### Requirement 3

**User Story:** As a document processor, I want to enhance the existing `_set_style_rule` method to avoid direct Normal/正文 modifications, so that cover pages using these styles remain unchanged.

#### Acceptance Criteria

1. WHEN executing set_style_rule on Normal/正文 THEN the system SHALL create and use "BodyText (AutoWord)" instead
2. WHEN creating body text styles THEN the system SHALL integrate with existing Word COM style operations
3. WHEN reassigning paragraphs THEN the system SHALL extend existing paragraph processing in `_apply_styles_to_content`
4. WHEN logging operations THEN the system SHALL use existing logger to record style separation activities
5. WHEN handling style mappings THEN the system SHALL extend existing `style_mappings` dictionary

### Requirement 4

**User Story:** As a document processor, I want to improve the existing cover detection logic, so that more cover page patterns are correctly identified and protected.

#### Acceptance Criteria

1. WHEN detecting covers THEN the system SHALL enhance existing `_is_cover_or_toc_content` method
2. WHEN analyzing content THEN the system SHALL improve existing keyword detection and style analysis
3. WHEN processing shapes THEN the system SHALL add anchor page checking to existing shape iteration
4. WHEN inserting section breaks THEN the system SHALL enhance existing `_insert_page_break` logic
5. WHEN validating results THEN the system SHALL extend existing validation in `_validate_result`