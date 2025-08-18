# Implementation Plan

- [x] 1. Set up vNext project structure and core data models






  - Create autoword/vnext/ directory structure with extractor/, planner/, executor/, validator/, auditor/ modules
  - Define StructureV1, PlanV1, InventoryFullV1 data models with Pydantic validation
  - Create JSON schema files for structure.v1.json and plan.v1.json validation
  - Implement base exception classes for vNext pipeline errors
  - _Requirements: 1.1, 1.2, 7.1, 7.2_

- [x] 2. Implement Extractor module for DOCX to JSON conversion






  - Create DocumentExtractor class with Word COM integration for structure extraction
  - Implement extract_structure() method to generate structure.v1.json with skeleton data
  - Implement extract_inventory() method to capture OOXML fragments and media in inventory.full.v1.json
  - Add metadata extraction (title, author, creation/modification times, Word version)
  - Create comprehensive unit tests with mock Word COM objects
  - _Requirements: 1.1, 1.2, 1.3, 1.4, 7.3, 7.4_

- [x] 3. Build Planner module with LLM integration and schema validation






  - Create DocumentPlanner class with LLM client integration for plan generation
  - Implement generate_plan() method that sends structure.v1.json + user intent to LLM
  - Add strict JSON schema validation for plan.v1.json with whitelist operation checking
  - Implement plan rejection logic that returns INVALID_PLAN for non-compliant plans
  - Create unit tests with mock LLM responses and schema validation scenarios
  - _Requirements: 2.1, 2.2, 2.3, 2.4, 2.5, 7.2, 7.5_



- [x] 4. Create Executor module with atomic operations implementation




  - Create DocumentExecutor class with Word COM automation for plan execution
  - Implement delete_section_by_heading atomic operation with heading text matching
  - Implement update_toc and delete_toc atomic operations for TOC management
  - Implement set_style_rule atomic operation for style object modification
  - Implement reassign_paragraphs_to_style atomic operation for paragraph style assignment
  - Implement clear_direct_formatting atomic operation with explicit authorization
  - _Requirements: 3.1, 3.2, 3.3, 3.4, 3.5, 3.6_

- [x] 5. Add localization support with style aliases and font fallbacks
  - Create LocalizationManager class with style alias mapping (Heading 1 ↔ 标题 1)
  - Implement font fallback system for East Asian fonts (楷体→楷体_GB2312→STKaiti)
  - Add dynamic style mapping detection after document opening
  - Implement font availability detection and fallback selection
  - Create warnings.log integration for font fallback notifications
  - _Requirements: 6.1, 6.2, 6.3, 6.4, 6.5_

- [x] 6. Implement Validator module with assertion checking and rollback





  - Create DocumentValidator class for post-execution validation
  - Implement chapter assertions to verify no "摘要/参考文献" at level 1
  - Implement style assertions for H1/H2/Normal style specifications (fonts, sizes, spacing)
  - Implement TOC assertions to verify consistency between TOC items and heading tree
  - Implement pagination assertions with Fields.Update() and Repaginate() verification
  - Add rollback functionality that restores original DOCX on validation failure
  - _Requirements: 4.1, 4.2, 4.3, 4.4, 4.5, 4.6_

- [x] 7. Build Auditor module for complete audit trail generation





  - Create DocumentAuditor class for timestamped audit directory creation
  - Implement snapshot saving for before/after DOCX files and structure JSON files
  - Implement diff report generation comparing before/after document structures
  - Add warnings.log creation for NOOP operations and font fallbacks
  - Implement result.status.txt writing with SUCCESS/ROLLBACK/FAILED_VALIDATION status
  - _Requirements: 5.1, 5.2, 5.3, 5.4, 5.5_



- [x] 8. Create comprehensive JSON schema validation system





  - Implement schema validation for structure.v1.json with all required fields
  - Implement schema validation for plan.v1.json with whitelist operation enforcement
  - Add schema validation for inventory.full.v1.json with OOXML fragment storage
  - Create schema validation utilities with detailed error reporting
  - Write comprehensive tests for schema validation with valid/invalid inputs
  - _Requirements: 7.1, 7.2, 7.3, 7.4, 7.5_

- [x] 9. Implement error handling and recovery system





  - Create comprehensive exception handling for all pipeline stages
  - Implement automatic rollback on any exception with FAILED_VALIDATION status
  - Add NOOP operation logging to warnings.log for unmatched operations
  - Implement security violation detection and logging for unauthorized operations
  - Create revision handling strategy configuration for documents with track changes
  - _Requirements: 8.1, 8.2, 8.3, 8.4, 8.5_

- [x] 10. Build complex document element handling





  - Extend Extractor to handle tables with cell references to paragraph indexes
  - Add formula and content control detection with inventory.full storage
  - Implement chart/SmartArt/OLE object handling as OOXML/binary in inventory
  - Add footnote/endnote reference preservation in structure extraction
  - Implement cross-reference identification and relationship mapping
  - _Requirements: 9.1, 9.2, 9.3, 9.4, 9.5_

- [x] 11. Create runtime constraint enforcement system




  - Implement whitelist operation validation that rejects non-approved operations
  - Add string replacement prevention for content and TOC display text modifications
  - Enforce Word object layer modifications only (no direct text manipulation)
  - Implement LLM output validation with JSON schema gateway
  - Add user input sanitization and parameter validation for all operations
  - _Requirements: 10.1, 10.2, 10.3, 10.4, 10.5_

- [x] 12. Build main pipeline orchestrator








  - Create VNextPipeline class that orchestrates all five modules in sequence
  - Implement complete error handling with rollback at pipeline level
  - Add progress reporting and status updates throughout pipeline execution
  - Create timestamped run directory management with fixed file naming
  - Write integration tests for complete pipeline execution with sample documents
  - _Requirements: 1.1, 2.1, 3.1, 4.1, 5.1_

- [x] 13. Implement DoD validation test suite




  - Create test for normal paper processing (delete 摘要/参考文献, update TOC, style application)
  - Create test for no-TOC document handling with NOOP update_toc operation
  - Create test for duplicate heading handling with occurrence_index specification
  - Create test for headings in tables with proper section deletion
  - Create test for missing font handling with fallback chain and warnings.log
  - Create test for complex objects (formulas/content controls/charts) preservation
  - Create test for revision tracking handling with both accept/reject and bypass modes
  - _Requirements: All requirements validation through DoD scenarios_

- [x] 14. Create configuration and deployment system





  - Create configuration files for style aliases and font fallback tables
  - Implement JSON schema documentation (structure.v1.md and plan.v1.md)
  - Create sample input/output packages for all 7 DoD validation scenarios
  - Write KNOWN-ISSUES.md documenting uncovered edge cases and handling recommendations
  - Set up packaging system for complete vNext pipeline deployment
  - _Requirements: 6.1, 6.3, 7.1, 7.2_

- [x] 15. Add comprehensive logging and monitoring





  - Implement detailed operation logging throughout all pipeline stages
  - Add performance monitoring for large document processing
  - Create debug logging for troubleshooting complex document scenarios
  - Implement memory usage monitoring and optimization alerts
  - Add execution time tracking for each pipeline stage and atomic operation
  - _Requirements: 8.1, 8.2, 8.5_

- [x] 16. Build CLI interface for vNext pipeline





  - Create command-line interface for vNext pipeline execution
  - Add options for user intent specification and configuration overrides
  - Implement batch processing capabilities for multiple documents
  - Add dry-run mode for plan generation without execution
  - Create verbose output options for debugging and monitoring
  - _Requirements: 2.1, 3.1, 5.1_

- [x] 17. Create performance optimization system




  - Implement efficient Word COM object management with proper disposal
  - Add document structure caching to minimize repeated COM calls
  - Optimize JSON serialization/deserialization for large structures
  - Implement memory-efficient processing for very large documents
  - Add parallel processing capabilities where COM threading allows
  - _Requirements: 1.4, 3.1, 7.3_


- [x] 18. Implement advanced validation and quality assurance




  - Create document integrity verification beyond basic assertions
  - Add style consistency checking across entire document
  - Implement cross-reference validation and repair capabilities
  - Add document accessibility compliance checking
  - Create formatting quality metrics and reporting
  - _Requirements: 4.1, 4.2, 4.3, 4.4_
-

- [-] 19. Build comprehensive test automation


  - Create automated test suite for all atomic operations with real Word documents
  - Implement regression testing for all DoD scenarios
  - Add performance benchmarking tests for large document processing
  - Create stress testing for error handling and rollback scenarios
  - Implement continuous integration testing with multiple Word versions
  - _Requirements: All requirements through comprehensive testing_

- [x] 20. Create documentation and user guides





  - Write technical documentation for vNext architecture and implementation
  - Create user guide for vNext pipeline usage and configuration
  - Document all atomic operations with examples and parameters
  - Create troubleshooting guide for common issues and solutions
  - Write API documentation for all public interfaces and data models
  - _Requirements: All requirements through complete documentation_