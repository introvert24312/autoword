# AutoWord Document Automation - Design Document

## Overview

AutoWord is a Windows desktop application that automates Word document editing by converting comments into executable tasks through LLM processing. The system uses a pipeline architecture with strict formatting protection, Word COM automation, and a modern PySide6 GUI.

The core design principle is "safe automation" - the system can only modify content and formatting when explicitly authorized by document comments, with multiple layers of protection against unauthorized changes.

## Architecture

### High-Level Architecture

```
┌─────────────────┐    ┌──────────────────┐    ┌─────────────────┐
│   PySide6 GUI   │◄──►│  Core Pipeline   │◄──►│   Word COM      │
│                 │    │                  │    │   Automation    │
│ • Files         │    │ • DocLoader      │    │                 │
│ • Inspect       │    │ • DocInspector   │    │ • Range API     │
│ • Plan          │    │ • PromptBuilder  │    │ • Bookmark API  │
│ • Run           │    │ • Planner        │    │ • Styles API    │
│ • Review        │    │ • WordExecutor   │    │ • TOC API       │
└─────────────────┘    │ • Validator      │    └─────────────────┘
                       │ • Exporter       │
                       └──────────────────┘
                                │
                       ┌──────────────────┐
                       │   LLM Services   │
                       │                  │
                       │ • GPT-5          │
                       │ • Claude 3.7     │
                       │ • JSON Schema    │
                       └──────────────────┘
```

### Threading Model

- **Main Thread**: GUI operations and user interaction
- **Worker Thread**: All Word COM operations (single-threaded COM requirement)
- **Communication**: Qt signals/slots for thread-safe progress updates

### Data Flow Pipeline

1. **Load Phase**: Document → Structure + Comments extraction
2. **Plan Phase**: Structure + Comments → LLM → JSON Tasks → Validated Plan
3. **Execute Phase**: Plan → Word COM operations → Modified Document
4. **Finalize Phase**: TOC/Link updates → Validation → Export

## Components and Interfaces

### Core Pipeline Components

#### DocLoader
```python
class DocLoader:
    def load_document(self, file_path: str) -> WordDocument:
        """Load document, disable tracking, create backup"""
    
    def create_backup(self, doc: WordDocument) -> str:
        """Create timestamped backup copy"""
    
    def close_document(self, doc: WordDocument, save: bool = True):
        """Safely close document with COM cleanup"""
```

#### DocInspector  
```python
class DocInspector:
    def extract_comments(self, doc: WordDocument) -> List[Comment]:
        """Extract all comments with metadata and anchor text"""
    
    def extract_structure(self, doc: WordDocument) -> DocumentStructure:
        """Extract headings, styles, TOC, hyperlinks, references"""
    
    def create_snapshot(self, doc: WordDocument) -> DocumentSnapshot:
        """Create complete document state snapshot for comparison"""
```

#### PromptBuilder
```python
class PromptBuilder:
    def build_system_prompt(self) -> str:
        """Build LLM system prompt with formatting constraints"""
    
    def build_user_prompt(self, structure: DocumentStructure, 
                         comments: List[Comment]) -> str:
        """Build user prompt with document context and schema"""
    
    def handle_context_overflow(self, content: str) -> List[str]:
        """Split content for chunked processing if needed"""
```

#### Planner
```python
class Planner:
    def generate_plan(self, structure: DocumentStructure, 
                     comments: List[Comment]) -> TaskPlan:
        """Generate and validate task plan from LLM"""
    
    def validate_tasks(self, tasks: List[Task]) -> ValidationResult:
        """Apply formatting protection rules"""
    
    def sort_by_dependencies(self, tasks: List[Task]) -> List[Task]:
        """Topological sort with risk-based ordering"""
```

#### WordExecutor
```python
class WordExecutor:
    def execute_plan(self, doc: WordDocument, plan: TaskPlan) -> ExecutionResult:
        """Execute all tasks with progress reporting"""
    
    def execute_task(self, doc: WordDocument, task: Task) -> TaskResult:
        """Execute single task with pre/post validation"""
    
    def create_locator(self, doc: WordDocument, locator_spec: dict) -> Locator:
        """Create bookmark/range locator for precise positioning"""
```

### GUI Components

#### MainWindow (FluentWindow + FramelessWindow)
```python
class MainWindow(FluentWindow, FramelessWindow):
    def __init__(self):
        """Initialize modern Windows UI with navigation"""
    
    def setup_navigation(self):
        """Setup Files/Inspect/Plan/Run/Review sections"""
    
    def connect_signals(self):
        """Connect worker thread signals to UI updates"""
```

#### Worker Thread
```python
class DocumentWorker(QThread):
    progress_updated = Signal(int, str)
    task_completed = Signal(str, bool, str)
    error_occurred = Signal(str, str)
    
    def run_pipeline(self, file_path: str, options: dict):
        """Execute complete pipeline in worker thread"""
```

### Data Models

#### Core Data Structures
```python
@dataclass
class Comment:
    id: str
    author: str
    page: int
    anchor_text: str
    comment_text: str
    range_start: int
    range_end: int

@dataclass
class Task:
    id: str
    type: TaskType
    source_comment_id: Optional[str]
    locator: Locator
    instruction: str
    dependencies: List[str]
    risk: RiskLevel
    requires_user_review: bool

@dataclass
class DocumentStructure:
    headings: List[Heading]
    styles: List[Style]
    toc_entries: List[TocEntry]
    hyperlinks: List[Hyperlink]
    references: List[Reference]
```

## Error Handling

### Formatting Protection Layers

1. **LLM Prompt Layer**: Hard constraints in system prompt
2. **Planning Layer**: Filter tasks without comment authorization
3. **Execution Layer**: Pre-execution validation blocks unauthorized tasks
4. **Validation Layer**: Post-execution comparison and rollback

### COM Error Handling

```python
class COMErrorHandler:
    def handle_com_exception(self, exc: Exception, context: str) -> ErrorAction:
        """Categorize COM errors and determine recovery action"""
    
    def retry_with_backoff(self, operation: Callable, max_retries: int = 3):
        """Retry COM operations with exponential backoff"""
    
    def cleanup_com_resources(self):
        """Ensure proper COM cleanup on errors"""
```

### LLM Error Handling

```python
class LLMErrorHandler:
    def handle_json_parse_error(self, response: str) -> dict:
        """Retry with explicit JSON format requirements"""
    
    def handle_context_overflow(self, content: str) -> List[str]:
        """Switch to chunked processing automatically"""
    
    def handle_api_error(self, error: Exception) -> str:
        """Provide user-friendly error messages"""
```

## Testing Strategy

### Unit Testing
- **Core Components**: Mock Word COM interfaces for isolated testing
- **Data Models**: Validate serialization/deserialization
- **Utilities**: Test helper functions and validators

### Integration Testing
- **LLM Integration**: Test with sample documents and known responses
- **COM Integration**: Test with real Word documents in controlled environment
- **GUI Integration**: Test signal/slot communication and state management

### End-to-End Testing
- **Pipeline Testing**: Complete workflow with sample documents
- **Error Scenarios**: Test error handling and recovery paths
- **Performance Testing**: Large documents and complex task plans

### Test Data Strategy
```
tests/
├── fixtures/
│   ├── sample_documents/
│   │   ├── simple.docx
│   │   ├── complex_with_toc.docx
│   │   └── comments_only.docx
│   ├── expected_outputs/
│   └── mock_llm_responses/
├── unit/
├── integration/
└── e2e/
```

## Security and Privacy

### API Key Management
- Environment variables only (GPT5_KEY, CLAUDE37_KEY)
- No keys in code, logs, or configuration files
- Secure key validation on startup

### Data Privacy
- Only comments and document structure sent to LLM by default
- Full content only when explicitly required
- No sensitive data in logs or exports
- Local processing with external LLM calls only

### COM Security
- Single-threaded COM operations
- Proper CoInitialize/CoUninitialize pairing
- Resource cleanup on exceptions
- Backup creation before any modifications

## Performance Considerations

### Memory Management
- Stream processing for large documents
- Lazy loading of document sections
- Proper COM object disposal
- GUI responsiveness through threading

### Optimization Strategies
- Batch similar operations
- Cache document structure during session
- Minimize Word COM round-trips
- Efficient JSON parsing and validation

### Scalability Limits
- Single document processing (no batch mode in MVP)
- Memory constraints for very large documents (>100MB)
- LLM context limits handled through chunking
- COM single-threading limits parallelization

## Deployment and Packaging

### Dependencies
```
Core:
- Python 3.10+
- PySide6 (GUI framework)
- pywin32 (Word COM automation)
- requests (LLM API calls)
- pydantic (data validation)

UI:
- qfluentwidgets (modern UI components)
- qframelesswindow (frameless window support)

Optional:
- orjson (faster JSON processing)
- python-dotenv (development environment)
```

### Packaging Strategy
```bash
pyinstaller --onefile --windowed gui/main.py \
    --name AutoWord \
    --add-data "schemas/*.json;schemas/" \
    --add-data "qfluentwidgets/resources;qfluentwidgets/resources" \
    --hidden-import win32timezone
```

### Installation Requirements
- Windows 10 or later
- Microsoft Word (any recent version with COM support)
- .NET Framework (for some COM operations)
- 4GB RAM minimum, 8GB recommended
- 500MB disk space for installation

This design provides a robust foundation for the AutoWord system while maintaining the strict formatting protection and reliable Word automation that are core to the project's success.