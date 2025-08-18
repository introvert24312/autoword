# Implementation Plan

- [ ] 1. Create comment extraction infrastructure
  - Implement CommentExtractor class with Word COM integration for comment extraction
  - Create CommentsV1 data model with schema validation
  - Write comment anchor positioning logic (paragraph_start, paragraph_end, char_start, char_end)
  - Add comment filtering by resolved status and optional EXECUTE tags
  - _Requirements: 1.1, 1.2, 1.3, 1.5_

- [ ] 2. Implement scope detection and resolution system
  - Create ScopeResolver class with keyword-based scope detection
  - Implement explicit scope marker parsing (SCOPE=GLOBAL|SECTION|ANCHOR)
  - Add configurable keyword lists for GLOBAL/SECTION scope detection
  - Write scope hint detection logic with priority rules (explicit > keywords > default)
  - Create unit tests for scope detection with Chinese keywords
  - _Requirements: 2.1, 2.2, 2.3, 2.4, 2.5_

- [ ] 3. Build DSL parser for direct comment translation
  - Implement DSLParser class for structured command parsing
  - Add support for JSON fragment parsing in comments
  - Create pattern matching for common operations (DEL_SECTION, SET_STYLE, TOC)
  - Write validation logic for DSL syntax compliance
  - Add error handling for malformed DSL with fallback to LLM
  - _Requirements: 3.1, 3.5_

- [ ] 4. Implement LLM fallback parser with constraints
  - Create LLMParser class with scope-constrained prompt generation
  - Build prompt templates that include document structure context and scope boundaries
  - Implement whitelist enforcement for LLM-generated operations
  - Add timeout handling and error recovery for LLM requests
  - Write validation logic for LLM output against scope boundaries
  - _Requirements: 3.2, 3.3, 3.4_

- [ ] 5. Create priority-based plan merger
  - Implement PlanMerger class with L1a → L1b → L2 → L3 priority logic
  - Build conflict resolution for same-scope operations (document order execution)
  - Add cross-scope conflict handling (preserve local effects, GLOBAL applies to untouched areas)
  - Create operation source tracking (comment_id, template_name)
  - Write comprehensive conflict resolution tests
  - _Requirements: 4.1, 4.2, 4.3, 4.4, 4.5_

- [ ] 6. Extend executor with scope-aware execution
  - Extend DocumentExecutor with scope-ordered execution (ANCHOR → SECTION → GLOBAL)
  - Implement ScopeBoundaryGuard for operation boundary enforcement
  - Add anchor boundary validation for ANCHOR operations
  - Create section boundary validation for SECTION operations
  - Write scope violation detection and SCOPE_BLOCKED status handling
  - _Requirements: 5.1, 5.2, 5.3, 5.4, 5.5_

- [ ] 7. Build per-comment result tracking system
  - Create CommentProcessingResult data model with status tracking
  - Implement comment_id to operation mapping in execution results
  - Add detailed status reporting (APPLIED/SKIPPED:NO_TARGET/INVALID_PLAN/SCOPE_BLOCKED)
  - Create comment-specific warning and error logging
  - Write diff report generation with comment_id associations
  - _Requirements: 6.1, 6.2, 6.3, 6.4, 6.5_

- [ ] 8. Implement configuration and policy management
  - Create CommentConfig data model with all configuration options
  - Implement comment_keywords.yml loading for GLOBAL/SECTION keywords
  - Add comment_tags_policy.yml for tag filtering configuration
  - Create configuration validation and default value handling
  - Write configuration file parsing with error handling
  - _Requirements: 9.1, 9.2_

- [ ] 9. Add CLI support for comment processing
  - Extend CLI with --use-comments flag for enabling comment processing
  - Add --comments-tags=EXECUTE flag for tag-based filtering
  - Implement --dry-run flag for preview mode (generate comments.v1.json + plan.v1.json only)
  - Create CLI help documentation for comment-related options
  - Write CLI integration tests for comment processing flags
  - _Requirements: 9.5_

- [ ] 10. Implement performance optimizations
  - Add comment parsing result caching with hash-based keys
  - Implement parallel LLM parsing with rate limiting for multiple comments
  - Create batch processing for similar comment types
  - Add memory-efficient comment extraction for large documents
  - Write performance benchmarks for comment processing
  - _Requirements: 9.3, 9.4_

- [ ] 11. Build comprehensive error handling
  - Implement CommentProcessingError exception hierarchy
  - Add graceful handling for comment parsing failures (continue with other comments)
  - Create scope violation error handling with detailed logging
  - Implement LLM timeout handling with fallback to SKIPPED status
  - Add comment thread processing (latest comment only)
  - _Requirements: 10.1, 10.2, 10.3, 10.4, 10.5_

- [ ] 12. Integrate comment processing into main orchestrator
  - Extend main pipeline to include comment extraction stage
  - Add comment parsing and operation generation to planner phase
  - Implement progress reporting for comment processing stages
  - Create audit trail integration for comment-to-operation mapping
  - Write end-to-end integration tests for complete comment-driven workflow
  - _Requirements: 8.1, 8.2, 8.3_

- [ ] 13. Create comprehensive test suite for comment processing
  - Write unit tests for scope detection with all keyword combinations
  - Create integration tests for DSL parsing and LLM fallback
  - Implement end-to-end tests for all six validation scenarios (ANCHOR/SECTION/GLOBAL/conflict/keywords/tags)
  - Add performance tests for large comment volumes
  - Create test fixtures with sample documents containing various comment types
  - _Requirements: All requirements validation_

- [ ] 14. Add GUI integration for comment processing controls
  - Implement GUI toggles for comment processing configuration
  - Add dry-run preview functionality with comment→operation mapping display
  - Create comment processing status display in GUI
  - Implement configuration management UI for keywords and policies
  - Write GUI integration tests for comment processing features
  - _Requirements: 7.1, 7.2, 7.3, 7.4, 7.5_

- [ ] 15. Implement comprehensive logging and audit trails
  - Add comment_id tracking throughout the processing pipeline
  - Implement detailed logging for scope detection and operation generation
  - Create warnings.log integration for comment-specific issues
  - Add audit trail generation for comment processing results
  - Write log analysis tools for troubleshooting comment processing
  - _Requirements: 8.4, 8.5_

- [ ] 16. Create documentation and examples
  - Write comprehensive documentation for comment-driven planning feature
  - Create DSL syntax reference with examples
  - Add scope detection keyword reference
  - Implement example documents with various comment types
  - Write troubleshooting guide for common comment processing issues
  - _Requirements: User documentation and examples_