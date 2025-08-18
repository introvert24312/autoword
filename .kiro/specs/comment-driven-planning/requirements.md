# Requirements Document

## Introduction

Comment-Driven Planning extends AutoWord vNext with the ability to use Word document comments as instruction sources, supporting both local scope (ANCHOR) and global scope (GLOBAL/SECTION) operations. This feature enables users to provide instructions directly within documents through comments, with automatic scope detection and priority-based execution ordering.

## Requirements

### Requirement 1

**User Story:** As a document processor, I want to extract and parse Word comments as instruction sources, so that I can execute document modifications based on in-document annotations.

#### Acceptance Criteria

1. WHEN processing a DOCX file THEN the system SHALL extract all unresolved comments into comments.v1.json
2. WHEN extracting comments THEN the system SHALL capture comment_id, author, created_time, text, and anchor position
3. WHEN processing comments THEN the system SHALL optionally filter by EXECUTE tags based on configuration
4. WHEN parsing comment text THEN the system SHALL detect scope hints using keywords (全文/全局/本节/以下段落)
5. WHEN extracting comments THEN the system SHALL preserve anchor information (paragraph_start, paragraph_end, char_start, char_end)

### Requirement 2

**User Story:** As a document processor, I want automatic scope detection for comments, so that I can apply operations at the correct document level (ANCHOR/SECTION/GLOBAL).

#### Acceptance Criteria

1. WHEN parsing comment text THEN the system SHALL detect explicit scope markers (SCOPE=GLOBAL|SECTION|ANCHOR)
2. WHEN detecting keywords THEN the system SHALL classify GLOBAL scope for "全文/全局/全篇/整体/全文统一"
3. WHEN detecting keywords THEN the system SHALL classify SECTION scope for "本节/以该标题为范围/以下段落"
4. WHEN no scope indicators exist THEN the system SHALL default to ANCHOR scope
5. WHEN scope is determined THEN the system SHALL add scope metadata to each parsed operation

### Requirement 3

**User Story:** As a document processor, I want comment-to-operation parsing with DSL and LLM fallback, so that I can convert natural language comments into executable operations.

#### Acceptance Criteria

1. WHEN parsing comments THEN the system SHALL attempt direct DSL/JSON translation first
2. WHEN DSL parsing fails THEN the system SHALL use LLM parsing with scope and whitelist constraints
3. WHEN using LLM parsing THEN the system SHALL provide document structure context and comment scope
4. WHEN generating operations THEN the system SHALL enforce whitelist compliance and scope boundaries
5. WHEN parsing fails THEN the system SHALL mark comment as INVALID_DSL and continue processing

### Requirement 4

**User Story:** As a document processor, I want priority-based plan merging, so that I can ensure comment instructions take precedence over GUI templates with proper conflict resolution.

#### Acceptance Criteria

1. WHEN merging plans THEN the system SHALL prioritize L1a (ANCHOR/SECTION comments) over L1b (GLOBAL comments) over L2 (GUI) over L3 (defaults)
2. WHEN merging operations THEN the system SHALL execute ANCHOR operations first, then SECTION, then GLOBAL
3. WHEN conflicts occur within same scope THEN the system SHALL execute in document order
4. WHEN operations target deleted ranges THEN the system SHALL mark subsequent operations as NO_TARGET
5. WHEN GLOBAL operations conflict with local operations THEN the system SHALL preserve local effects and apply GLOBAL only to untouched areas

### Requirement 5

**User Story:** As a document processor, I want scope-aware execution ordering, so that I can ensure operations are applied in the correct sequence to prevent conflicts.

#### Acceptance Criteria

1. WHEN executing plans THEN the system SHALL process operations in strict order: ANCHOR → SECTION → GLOBAL
2. WHEN executing ANCHOR operations THEN the system SHALL only modify content within comment anchor boundaries
3. WHEN executing SECTION operations THEN the system SHALL operate on entire heading sections (heading to next same-level heading)
4. WHEN executing GLOBAL operations THEN the system SHALL apply to entire document while respecting local modifications
5. WHEN execution completes THEN the system SHALL perform unified cleanup: Fields.Update() → TOC.Update() → Repaginate()

### Requirement 6

**User Story:** As a document processor, I want per-comment result tracking and validation, so that I can provide detailed feedback on which comments were successfully applied.

#### Acceptance Criteria

1. WHEN executing comment-based operations THEN the system SHALL track results per comment_id
2. WHEN operations complete THEN the system SHALL generate status for each comment: APPLIED/SKIPPED:NO_TARGET/INVALID_PLAN/SCOPE_BLOCKED
3. WHEN generating diff reports THEN the system SHALL include comment_id mapping to applied changes
4. WHEN validation fails THEN the system SHALL rollback entire batch and report failed comment operations
5. WHEN warnings occur THEN the system SHALL log comment-specific issues to warnings.log

### Requirement 7

**User Story:** As a document processor, I want GUI integration with comment processing controls, so that I can configure comment-driven behavior and preview operations before execution.

#### Acceptance Criteria

1. WHEN configuring the system THEN the system SHALL provide toggle for "读取批注作为指令"
2. WHEN filtering comments THEN the system SHALL provide option to "只执行带EXECUTE标签的批注"
3. WHEN processing comments THEN the system SHALL provide option for "批注直译失败时降级LLM解析"
4. WHEN previewing operations THEN the system SHALL provide "预跑(Dry-Run)" mode to show comment→operation mapping
5. WHEN dry-run completes THEN the system SHALL generate comments.v1.json + plan.v1.json for user review

### Requirement 8

**User Story:** As a document processor, I want comprehensive comment operation logging and audit trails, so that I can track comment processing and troubleshoot issues.

#### Acceptance Criteria

1. WHEN processing comments THEN the system SHALL log comment_id mapping to generated operations
2. WHEN operations fail THEN the system SHALL log SCOPE_BLOCKED/NO_TARGET/INVALID_DSL with comment context
3. WHEN generating audit trails THEN the system SHALL include comment processing stages in progress reporting
4. WHEN creating warnings THEN the system SHALL associate warnings with originating comment_id
5. WHEN processing completes THEN the system SHALL provide summary of comment processing results

### Requirement 9

**User Story:** As a document processor, I want configurable comment processing policies, so that I can customize comment parsing behavior for different use cases.

#### Acceptance Criteria

1. WHEN configuring keywords THEN the system SHALL load GLOBAL/SECTION keywords from comment_keywords.yml
2. WHEN filtering comments THEN the system SHALL apply tag policies from comment_tags_policy.yml
3. WHEN processing multiple comments THEN the system SHALL support parallel LLM parsing with rate limiting
4. WHEN caching is enabled THEN the system SHALL cache comment parsing results for performance
5. WHEN CLI is used THEN the system SHALL support --use-comments/--comments-tags=EXECUTE/--dry-run flags

### Requirement 10

**User Story:** As a document processor, I want robust error handling for comment processing, so that I can maintain system stability when comment parsing fails.

#### Acceptance Criteria

1. WHEN comment parsing fails THEN the system SHALL continue processing other comments and log failures
2. WHEN scope detection fails THEN the system SHALL default to ANCHOR scope and log warning
3. WHEN LLM parsing times out THEN the system SHALL mark comment as SKIPPED and continue
4. WHEN operations exceed scope boundaries THEN the system SHALL block execution and log SCOPE_BLOCKED
5. WHEN comment threads exist THEN the system SHALL process only the latest comment in each thread