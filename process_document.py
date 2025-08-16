#!/usr/bin/env python3
"""
AutoWord 文档处理器
专门处理指定的Word文档
"""

import sys
import os
from pathlib import Path

# 添加项目路径
sys.path.append(str(Path(__file__).parent))

def main():
    # 你的文档路径
    document_path = r"C:\Users\y\Desktop\郭宇睿-现代汉语言文学在社会发展中的作用与影响研究2稿批注.docx"
    
    print("AutoWord Document Processor")
    print("=" * 50)
    print(f"Document: {document_path}")
    print()
    
    # 检查文档是否存在
    if not os.path.exists(document_path):
        print(f"ERROR: Document not found: {document_path}")
        print("Please check the file path and try again.")
        input("Press Enter to exit...")
        return False
    
    try:
        # 导入AutoWord模块
        from autoword.core.enhanced_executor import EnhancedExecutor, WorkflowMode
        from autoword.core.llm_client import ModelType
        
        print("✅ AutoWord modules loaded successfully")
        
        # 检查API密钥
        api_key_available = bool(os.getenv('GPT5_KEY') or os.getenv('CLAUDE37_KEY'))
        if api_key_available:
            print("✅ API key configured")
        else:
            print("⚠️  No API key configured - using demo mode")
            print("   Set GPT5_KEY or CLAUDE37_KEY environment variable for full functionality")
        
        print()
        print("Processing options:")
        print("1. DRY RUN - Preview changes (safe, no modifications)")
        print("2. SAFE MODE - Apply changes with backup")
        print("3. NORMAL MODE - Apply changes directly")
        print()
        
        while True:
            choice = input("Select mode (1-3): ").strip()
            if choice in ['1', '2', '3']:
                break
            print("Invalid choice. Please enter 1, 2, or 3.")
        
        # 设置执行模式
        if choice == '1':
            mode = WorkflowMode.DRY_RUN
            print("\n🔍 Running in DRY RUN mode (preview only)...")
        elif choice == '2':
            mode = WorkflowMode.SAFE
            print("\n🛡️ Running in SAFE mode (with backup)...")
        else:
            mode = WorkflowMode.NORMAL
            print("\n⚡ Running in NORMAL mode...")
        
        # 创建执行器
        executor = EnhancedExecutor(
            llm_model=ModelType.GPT5 if os.getenv('GPT5_KEY') else ModelType.CLAUDE37,
            visible=False  # 隐藏Word窗口
        )
        
        print("✅ Executor created")
        print()
        print("Starting document processing...")
        print("-" * 30)
        
        # 执行文档处理
        result = executor.execute_workflow(
            document_path=document_path,
            mode=mode
        )
        
        # 显示结果
        print()
        print("=" * 50)
        print("PROCESSING RESULTS")
        print("=" * 50)
        
        if result.success:
            print("✅ Processing completed successfully!")
        else:
            print("❌ Processing failed")
        
        print(f"📊 Total tasks: {result.total_tasks}")
        print(f"✅ Completed: {result.completed_tasks}")
        print(f"❌ Failed: {result.failed_tasks}")
        print(f"⏱️ Execution time: {result.execution_time:.2f} seconds")
        
        if result.task_results:
            print()
            print("Task Details:")
            print("-" * 20)
            for i, task_result in enumerate(result.task_results, 1):
                status = "✅" if task_result.success else "❌"
                print(f"{i:2d}. {status} {task_result.task_id}")
                print(f"     {task_result.message}")
                if task_result.error_details:
                    print(f"     Error: {task_result.error_details}")
                print()
        
        if result.error_summary:
            print("Error Summary:")
            print(result.error_summary)
        
        print()
        if mode == WorkflowMode.DRY_RUN:
            print("💡 This was a preview run. No changes were made to your document.")
            print("   To apply changes, run again and select SAFE or NORMAL mode.")
        elif mode == WorkflowMode.SAFE:
            print("🛡️ Changes applied with backup protection.")
            print("   Original document backed up automatically.")
        else:
            print("⚡ Changes applied directly to your document.")
        
        return result.success
        
    except ImportError as e:
        print(f"❌ Module import error: {e}")
        print("Please ensure AutoWord is properly installed.")
        return False
    except Exception as e:
        print(f"❌ Processing error: {e}")
        import traceback
        traceback.print_exc()
        return False
    
    finally:
        print()
        input("Press Enter to exit...")

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)