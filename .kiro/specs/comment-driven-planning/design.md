# Comment-Driven Planning - Design Document

## Overview

Comment-Driven Planning extends the AutoWord vNext pipeline by adding comment extraction and parsing capabilities that convert Word document comments into executable operations. The system supports three scope levels (ANCHOR, SECTION, GLOBAL) with priority-based merging and execution ordering to ensure predictable document modifications.

The core principle is "批注优先" (comments first) - comment-derived operations take precedence over GUI templates, with sophisticated conflict resolution and scope-aware execution to handle both local modifications ("这里删一段") and global changes ("全文改格式").

## Architecture

### Extended Pipeline Architecture

```
┌─────────────────┐    ┌─────────────────┐    ┌─────────────────┐
│   Extractor     │───►│    Planner      │───►│   Executor      │
│                 │    │                 │    │                 │
│ DOCX →          │    │ structure.json  │    │ plan.json →     │
│ structure.json  │    │ + comments.json │    │ modified DOCX   │
│ inventory.json  │    │ + user_intent → │    │                 │
│ comments.json   │◄───┤ plan.json       │    │                 │
└─────────────────┘    └─────────────────┘    └─────────────────┘
         │                       │                        │
         │              ┌─────────────────┐    ┌─────────────────┐
         │              │ Comment Parser  │    │   Validator     │
         │              │                 │    │                 │
         │              │ • Scope Detect  │    │ structure.json' │
         │              │ • DSL/LLM Parse │    │ → assertions    │
         │              │ • Priority Merge│    │ → comment status│
         │              └─────────────────┘    └─────────────────┘
         │                                              │
         │                                    ┌─────────────────┐
         └────────────────────────────────────┤    Auditor      │
                                              │                 │
                                              │ • Comment Audit │
                                              │ • Scope Reports │
                                              │ • Status Mapping│
                                              └─────────────────┘
```

### Comment Processing Data Flow

```
Input: DOCX File
    ↓
[Extractor] → structure.v1.json + inventory.full.v1.json + comments.v1.json
    ↓
[Comment Parser] → ops(L1a: ANCHOR/SECTION) + ops(L1b: GLOBAL)
    ↓
[Planner] + GUI/Template → Plan = Merge(L1a → L1b → L2 → L3)
    ↓
[Executor] → Execute by scope order: ANCHOR → SECTION → GLOBAL
    ↓
[Validator] → Per-comment status + overall validation
    ↓
[Auditor] → Comment-to-operation mapping + results
```

## Components and Interfaces

### 1. Comment Extractor (Extension to DocumentExtractor)

**Purpose**: Extract Word comments with anchor positioning and scope hints.

```python
class CommentExtractor:
    def extract_comments(self, docx_path: str, config: CommentConfig) -> CommentsV1:
        """Extract comments with filtering and scope detection"""
        
    def parse_comment_anchors(self, comment: Comment) -> CommentAnchor:
        """Extract precise anchor positioning"""
        
    def detect_scope_hints(self, comment_text: str) -> ScopeHint:
        """Detect scope keywords and explicit markers"""
        
    def filter_by_tags(self, comments: List[Comment], config: CommentConfig) -> List[Comment]:
        """Filter comments by EXECUTE tags if configured"""
```

**Comment Data Structure**:
```python
@dataclass
class CommentsV1:
    schema_version: str = "comments.v1"
    comments: List[CommentData]

@dataclass
class CommentData:
    comment_id: str
    author: str
    created_time: datetime
    resolved: bool
    text: str
    anchor: CommentAnchor
    tags: List[str]  # ["EXECUTE"] if present
    scope_hint: Optional[str]  # "GLOBAL|SECTION|ANCHOR|null"

@dataclass
class CommentAnchor:
    paragraph_start: int
    paragraph_end: int
    char_start: int
    char_end: int
    heading_hint: Optional[str]  # Nearest heading for context
```

### 2. Comment Parser

**Purpose**: Convert comments to operations with scope detection and DSL/LLM parsing.

```python
class CommentParser:
    def __init__(self, keywords_config: CommentKeywordsConfig, llm_client: LLMClient):
        self.scope_resolver = ScopeResolver(keywords_config)
        self.dsl_parser = DSLParser()
        self.llm_parser = LLMParser(llm_client)
    
    def parse_comments_to_ops(self, comments: CommentsV1, structure: StructureV1) -> Tuple[List[Operation], List[Operation]]:
        """Parse comments to L1a (ANCHOR/SECTION) and L1b (GLOBAL) operations"""
        
    def parse_single_comment(self, comment: CommentData, structure: StructureV1) -> List[Operation]:
        """Parse individual comment with DSL → LLM fallback"""
        
    def resolve_scope(self, comment: CommentData) -> ScopeType:
        """Determine final scope based on hints and keywords"""
```

**Scope Resolution Logic**:
```python
class ScopeResolver:
    GLOBAL_KEYWORDS = ["全文", "全局", "全篇", "整体", "全文统一", "全文应用"]
    SECTION_KEYWORDS = ["本节", "以该标题为范围", "以下段落"]
    
    def resolve_scope(self, comment_text: str, scope_hint: Optional[str]) -> ScopeType:
        """Priority: explicit marker > keywords > default ANCHOR"""
        
        # 1. Check explicit markers
        if "SCOPE=GLOBAL" in comment_text:
            return ScopeType.GLOBAL
        elif "SCOPE=SECTION" in comment_text:
            return ScopeType.SECTION
        elif "SCOPE=ANCHOR" in comment_text:
            return ScopeType.ANCHOR
            
        # 2. Check keywords
        if any(keyword in comment_text for keyword in self.GLOBAL_KEYWORDS):
            return ScopeType.GLOBAL
        elif any(keyword in comment_text for keyword in self.SECTION_KEYWORDS):
            return ScopeType.SECTION
            
        # 3. Default to ANCHOR
        return ScopeType.ANCHOR
```

### 3. DSL Parser

**Purpose**: Direct translation of structured comment text to operations.

```python
class DSLParser:
    def parse_dsl(self, comment_text: str, scope: ScopeType, anchor: CommentAnchor) -> Optional[List[Operation]]:
        """Attempt direct DSL/JSON translation"""
        
    def parse_json_fragment(self, text: str) -> Optional[dict]:
        """Parse minimal JSON fragments in comments"""
        
    def parse_structured_commands(self, text: str, scope: ScopeType) -> Optional[List[Operation]]:
        """Parse structured commands like DEL_SECTION, SET_STYLE"""

# Example DSL patterns:
# "DEL_SECTION: 讨论 level=1"
# "SET_STYLE: Heading 1 east_asian=楷体 size=12 bold=true"
# "SCOPE=GLOBAL; SET_STYLE: Normal east_asian=宋体 size=12; TOC: UPDATE"
```

### 4. LLM Parser (Fallback)

**Purpose**: Natural language comment parsing with scope and whitelist constraints.

```python
class LLMParser:
    def parse_with_llm(self, comment: CommentData, structure: StructureV1, scope: ScopeType) -> List[Operation]:
        """Parse natural language comment with constraints"""
        
    def build_constrained_prompt(self, comment: CommentData, structure: StructureV1, scope: ScopeType) -> str:
        """Build prompt with scope boundaries and whitelist"""
        
    def validate_llm_output(self, operations: List[Operation], scope: ScopeType, anchor: CommentAnchor) -> List[Operation]:
        """Validate LLM output against scope boundaries"""
```

**LLM Prompt Template**:
```
You are parsing a Word document comment into atomic operations.

Comment: "{comment_text}"
Scope: {scope_type}
Anchor: {anchor_info}
Available Operations: {whitelist}

Constraints:
- ANCHOR scope: Only modify paragraphs {paragraph_start}-{paragraph_end}
- SECTION scope: Only modify section under heading "{heading_hint}"
- GLOBAL scope: Document-wide operations only
- Use only whitelisted operations
- Output valid JSON matching plan.v1 schema

Generate operations:
```

### 5. Priority-Based Plan Merger

**Purpose**: Merge comment operations with GUI templates using priority rules.

```python
class PlanMerger:
    def merge_plans(self, 
                   comment_ops_local: List[Operation],  # L1a: ANCHOR/SECTION
                   comment_ops_global: List[Operation], # L1b: GLOBAL
                   gui_ops: List[Operation],            # L2: GUI/Template
                   default_ops: List[Operation]         # L3: System defaults
                   ) -> PlanV1:
        """Merge with priority: L1a → L1b → L2 → L3"""
        
    def resolve_conflicts(self, operations: List[Operation]) -> List[Operation]:
        """Handle conflicts within same scope"""
        
    def validate_scope_boundaries(self, operations: List[Operation]) -> List[Operation]:
        """Ensure operations respect scope boundaries"""
```

**Conflict Resolution Rules**:
```python
class ConflictResolver:
    def resolve_same_scope_conflicts(self, ops: List[Operation]) -> List[Operation]:
        """Within same scope: document order execution, later ops may get NO_TARGET"""
        
    def resolve_cross_scope_conflicts(self, anchor_ops: List[Operation], 
                                    section_ops: List[Operation], 
                                    global_ops: List[Operation]) -> List[Operation]:
        """Cross-scope: preserve local effects, GLOBAL applies to untouched areas"""
```

### 6. Scope-Aware Executor (Extension to DocumentExecutor)

**Purpose**: Execute operations in strict scope order with boundary enforcement.

```python
class ScopeAwareExecutor(DocumentExecutor):
    def execute_plan_with_scopes(self, plan: PlanV1, docx_path: str) -> ExecutionResult:
        """Execute in order: ANCHOR → SECTION → GLOBAL"""
        
    def execute_anchor_operations(self, ops: List[Operation], doc: WordDocument) -> List[OperationResult]:
        """Execute operations within anchor boundaries only"""
        
    def execute_section_operations(self, ops: List[Operation], doc: WordDocument) -> List[OperationResult]:
        """Execute operations within section boundaries"""
        
    def execute_global_operations(self, ops: List[Operation], doc: WordDocument) -> List[OperationResult]:
        """Execute document-wide operations"""
        
    def enforce_scope_boundaries(self, op: Operation, scope: ScopeType, doc: WordDocument) -> bool:
        """Validate operation doesn't exceed scope boundaries"""
```

**Scope Boundary Enforcement**:
```python
class ScopeBoundaryGuard:
    def check_anchor_boundaries(self, op: Operation, anchor: CommentAnchor, doc: WordDocument) -> bool:
        """Ensure operation only affects anchor range"""
        
    def check_section_boundaries(self, op: Operation, section_info: SectionInfo, doc: WordDocument) -> bool:
        """Ensure operation only affects section range"""
        
    def check_global_scope_allowed(self, op: Operation) -> bool:
        """Ensure operation is valid for global scope"""
```

## Data Models

### Extended Operation Schema

```python
@dataclass
class Operation:
    # Existing operation fields...
    source: OperationSource
    scope: OperationScope

@dataclass
class OperationSource:
    type: str  # "COMMENT", "GUI", "DEFAULT"
    comment_id: Optional[str] = None
    template_name: Optional[str] = None

@dataclass
class OperationScope:
    type: ScopeType  # ANCHOR, SECTION, GLOBAL
    anchor: Optional[CommentAnchor] = None
    section: Optional[SectionInfo] = None

@dataclass
class SectionInfo:
    heading_text: str
    level: int
    occurrence: int  # For duplicate headings
```

### Comment Processing Results

```python
@dataclass
class CommentProcessingResult:
    comment_id: str
    status: CommentStatus  # APPLIED, SKIPPED, FAILED
    operations_generated: List[str]  # Operation IDs
    error_message: Optional[str] = None
    warnings: List[str] = field(default_factory=list)

class CommentStatus(Enum):
    APPLIED = "APPLIED"
    SKIPPED_NO_TARGET = "SKIPPED:NO_TARGET"
    SKIPPED_INVALID_PLAN = "SKIPPED:INVALID_PLAN"
    SKIPPED_SCOPE_BLOCKED = "SKIPPED:SCOPE_BLOCKED"
    SKIPPED_TAG_FILTER = "SKIPPED:TAG_FILTER"
    FAILED_PARSING = "FAILED_PARSING"
```

### Configuration Schema

```python
@dataclass
class CommentConfig:
    enabled: bool = True
    execute_tags_only: bool = False
    llm_fallback_enabled: bool = True
    parallel_parsing: bool = True
    cache_enabled: bool = True
    keywords_file: str = "comment_keywords.yml"
    tags_policy_file: str = "comment_tags_policy.yml"
```

## Error Handling

### Comment Processing Error Strategy

```python
class CommentProcessingError(Exception):
    def __init__(self, comment_id: str, error_type: str, message: str):
        self.comment_id = comment_id
        self.error_type = error_type
        self.message = message

class CommentErrorHandler:
    def handle_parsing_error(self, comment: CommentData, error: Exception) -> CommentProcessingResult:
        """Log error and continue with other comments"""
        
    def handle_scope_violation(self, comment: CommentData, operation: Operation) -> CommentProcessingResult:
        """Block operation and log SCOPE_BLOCKED"""
        
    def handle_llm_timeout(self, comment: CommentData) -> CommentProcessingResult:
        """Mark as SKIPPED and continue"""
        
    def handle_conflict_resolution(self, conflicting_ops: List[Operation]) -> List[Operation]:
        """Resolve conflicts and mark affected operations"""
```

### Rollback Strategy for Comment Operations

```python
class CommentAwareRollback:
    def rollback_with_comment_context(self, original_docx: str, comment_results: List[CommentProcessingResult]):
        """Rollback and provide comment-specific error context"""
        
    def generate_comment_error_report(self, results: List[CommentProcessingResult]) -> CommentErrorReport:
        """Generate detailed report of comment processing failures"""
```

## Testing Strategy

### Comment Processing Test Cases

```python
class CommentProcessingTests:
    def test_scope_detection(self):
        """Test GLOBAL/SECTION/ANCHOR keyword detection"""
        
    def test_dsl_parsing(self):
        """Test direct DSL translation"""
        
    def test_llm_fallback(self):
        """Test LLM parsing with scope constraints"""
        
    def test_priority_merging(self):
        """Test L1a → L1b → L2 → L3 priority"""
        
    def test_conflict_resolution(self):
        """Test same-scope and cross-scope conflicts"""
        
    def test_scope_boundary_enforcement(self):
        """Test operations don't exceed scope boundaries"""
```

### Integration Test Scenarios

1. **ANCHOR Scope**: Comment on paragraph → modify only that paragraph
2. **SECTION Scope**: Comment "本节删除" → delete entire section + update TOC
3. **GLOBAL Scope**: Comment "全文统一样式" → apply styles globally while preserving local changes
4. **Conflict Resolution**: Multiple comments on same range → document order execution
5. **Priority Override**: Comment operations override GUI templates
6. **Tag Filtering**: Only execute comments with EXECUTE tags

### Performance Testing

```python
class CommentPerformanceTests:
    def test_large_comment_volume(self):
        """Test processing documents with 100+ comments"""
        
    def test_parallel_llm_parsing(self):
        """Test concurrent LLM parsing with rate limiting"""
        
    def test_comment_parsing_cache(self):
        """Test caching effectiveness for repeated processing"""
```

## Security and Constraints

### Comment Processing Security

```python
class CommentSecurityGuard:
    def validate_comment_operations(self, operations: List[Operation]) -> List[Operation]:
        """Ensure comment-derived operations are whitelisted"""
        
    def sanitize_comment_text(self, comment_text: str) -> str:
        """Sanitize comment text before LLM processing"""
        
    def enforce_scope_isolation(self, operation: Operation, scope: ScopeType) -> bool:
        """Prevent scope boundary violations"""
```

### LLM Integration Constraints

- Comment text sanitization before LLM processing
- Scope-constrained prompts with whitelist enforcement
- Rate limiting for parallel comment parsing
- Timeout handling for LLM requests
- Validation of LLM output against scope boundaries

## Performance Considerations

### Comment Processing Optimization

```python
class CommentOptimizer:
    def batch_similar_comments(self, comments: List[CommentData]) -> List[List[CommentData]]:
        """Group similar comments for batch LLM processing"""
        
    def cache_comment_parsing(self, comment_hash: str, operations: List[Operation]):
        """Cache parsed operations for identical comments"""
        
    def parallel_llm_processing(self, comments: List[CommentData]) -> List[List[Operation]]:
        """Process multiple comments concurrently with rate limiting"""
```

### Memory and Performance Limits

- Comment extraction streaming for large documents
- LLM request batching and rate limiting
- Comment parsing result caching
- Efficient scope boundary checking
- Minimal Word COM round-trips for scope validation

This design provides a comprehensive foundation for implementing comment-driven planning that integrates seamlessly with the existing AutoWord vNext pipeline while maintaining all security, validation, and auditability requirements.