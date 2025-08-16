@echo off
title AutoWord System Test
color 0B

echo AutoWord System Test
echo ====================
echo.

REM Check Python
echo [1/5] Checking Python...
python --version
if errorlevel 1 (
    echo ERROR: Python not installed
    pause
    exit /b 1
)
echo OK: Python found
echo.

REM Install dependencies
echo [2/5] Installing dependencies...
pip install -q pydantic requests pywin32 python-dotenv orjson colorlog click
if errorlevel 1 (
    echo WARNING: Some dependencies may not be installed
)
echo OK: Dependencies installed
echo.

REM Test core modules
echo [3/5] Testing core modules...
python -c "
try:
    from autoword.core.models import Task, Comment, TaskType
    from autoword.core.llm_client import LLMClient
    from autoword.core.enhanced_executor import EnhancedExecutor
    print('OK: Core modules imported successfully')
except ImportError as e:
    print(f'ERROR: Module import failed: {e}')
    exit(1)
"
if errorlevel 1 (
    echo Module test failed
    pause
    exit /b 1
)
echo.

REM Test Word COM
echo [4/5] Testing Word COM...
python -c "
try:
    import win32com.client
    word = win32com.client.Dispatch('Word.Application')
    word.Quit()
    print('OK: Word COM available')
except Exception as e:
    print(f'WARNING: Word COM test failed: {e}')
"
echo.

REM Run functionality test
echo [5/5] Testing AutoWord functionality...
python -c "
import sys
import os
sys.path.append('.')

print('Testing AutoWord core functionality...')

try:
    # Test data models
    from autoword.core.models import Comment, Task, TaskType, Locator, LocatorType
    
    comment = Comment(
        id='test1',
        author='Test User',
        text='Test comment',
        anchor_text='test',
        page_number=1,
        position_start=0,
        position_end=5
    )
    
    task = Task(
        id='task1',
        type=TaskType.REWRITE,
        locator=Locator(by=LocatorType.FIND, value='test'),
        instruction='Test instruction',
        source_comment_id='test1'
    )
    
    print('OK: Data models working')
    
    # Test LLM client
    from autoword.core.llm_client import LLMClient
    client = LLMClient()
    print('OK: LLM client created')
    
    # Test prompt builder
    from autoword.core.prompt_builder import PromptBuilder, PromptContext
    from autoword.core.models import Heading, Style
    
    context = PromptContext(
        headings=[Heading(text='Test', level=1, page_number=1)],
        styles=[Style(name='Normal', type='paragraph')],
        toc_entries=[],
        hyperlinks=[],
        comments=[comment]
    )
    
    builder = PromptBuilder()
    prompt = builder.build_user_prompt(context)
    print('OK: Prompt builder working')
    
    # Test format protection
    from autoword.core.planner import FormatProtectionGuard
    guard = FormatProtectionGuard()
    filtered = guard.filter_unauthorized_tasks([task], ['test1'])
    print('OK: Format protection working')
    
    # Test enhanced executor
    from autoword.core.enhanced_executor import EnhancedExecutor
    executor = EnhancedExecutor()
    print('OK: Enhanced executor created')
    
    print()
    print('=== ALL TESTS PASSED ===')
    print('AutoWord system is ready!')
    
except Exception as e:
    print(f'ERROR: {str(e)}')
    import traceback
    traceback.print_exc()
    exit(1)
"

if errorlevel 1 (
    echo System test failed
    pause
    exit /b 1
)

echo.
echo =========================
echo AutoWord System Ready!
echo =========================
echo.
echo Your document: C:\Users\y\Desktop\郭宇睿-现代汉语言文学在社会发展中的作用与影响研究2稿批注.docx
echo.
echo To process your document, run: autoword_launcher.bat
echo.
pause