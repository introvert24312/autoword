@echo off
title AutoWord Document Processor
color 0A

echo AutoWord Document Automation System
echo ====================================
echo.

REM Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found. Please install Python 3.8+
    pause
    exit /b 1
)

REM Install dependencies if needed
echo Installing dependencies...
pip install -q pydantic requests pywin32 python-dotenv orjson colorlog click

REM Set your document path
set DOCUMENT_PATH=C:\Users\y\Desktop\郭宇睿-现代汉语言文学在社会发展中的作用与影响研究2稿批注.docx

echo.
echo Processing document: %DOCUMENT_PATH%
echo.

REM Run AutoWord
python -c "
import sys
import os
sys.path.append('.')

# Check if document exists
document_path = r'%DOCUMENT_PATH%'
if not os.path.exists(document_path):
    print('ERROR: Document not found:', document_path)
    input('Press Enter to exit...')
    sys.exit(1)

print('Starting AutoWord processing...')
print('Document:', document_path)
print()

try:
    from autoword.core.enhanced_executor import EnhancedExecutor, WorkflowMode
    
    # Create executor
    executor = EnhancedExecutor()
    
    # Process document in dry run mode first
    print('Running in DRY RUN mode (safe preview)...')
    result = executor.execute_workflow(
        document_path=document_path,
        mode=WorkflowMode.DRY_RUN
    )
    
    print()
    print('=== PROCESSING RESULTS ===')
    print(f'Success: {result.success}')
    print(f'Total tasks: {result.total_tasks}')
    print(f'Completed: {result.completed_tasks}')
    print(f'Failed: {result.failed_tasks}')
    print(f'Execution time: {result.execution_time:.2f}s')
    
    if result.task_results:
        print()
        print('Task Details:')
        for i, task_result in enumerate(result.task_results, 1):
            status = 'SUCCESS' if task_result.success else 'FAILED'
            print(f'  {i}. [{status}] {task_result.task_id}: {task_result.message}')
    
    if not result.success and result.error_summary:
        print()
        print('Error Summary:', result.error_summary)
    
    print()
    if result.success and result.total_tasks > 0:
        print('DRY RUN completed successfully!')
        print('To apply changes, set mode to NORMAL or SAFE in the script.')
    elif result.total_tasks == 0:
        print('No tasks generated. Check if document has comments.')
    else:
        print('Processing failed. Check error details above.')
        
except Exception as e:
    print(f'ERROR: {str(e)}')
    import traceback
    traceback.print_exc()

print()
input('Press Enter to exit...')
"

pause