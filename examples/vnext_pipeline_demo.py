"""
VNext Pipeline Demo

This example demonstrates how to use the VNext pipeline orchestrator
for complete document processing through all five stages.
"""

import os
import sys
import logging
from pathlib import Path

# Add the project root to the path
sys.path.insert(0, str(Path(__file__).parent.parent))

from autoword.vnext import VNextPipeline, ProcessingResult
from autoword.core.llm_client import LLMClient, ModelType


def setup_logging():
    """Setup logging for the demo."""
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.StreamHandler(),
            logging.FileHandler('vnext_pipeline_demo.log')
        ]
    )


def progress_callback(stage_name: str, progress_percent: int):
    """Progress callback function for pipeline updates."""
    print(f"[{progress_percent:3d}%] {stage_name}")


def main():
    """Main demo function."""
    setup_logging()
    logger = logging.getLogger(__name__)
    
    print("=== AutoWord vNext Pipeline Demo ===")
    print()
    
    # Configuration
    docx_path = "sample_document.docx"
    user_intent = "删除摘要和参考文献部分，更新目录，应用标准样式格式"
    audit_dir = "./demo_audit_trails"
    
    # Check if sample document exists
    if not os.path.exists(docx_path):
        print(f"Error: Sample document '{docx_path}' not found.")
        print("Please create a sample DOCX file or update the docx_path variable.")
        return
    
    try:
        # Initialize LLM client
        print("1. Initializing LLM client...")
        llm_client = LLMClient(
            model=ModelType.GPT4,
            api_key=os.getenv("OPENAI_API_KEY"),
            temperature=0.1
        )
        
        # Initialize pipeline
        print("2. Initializing vNext pipeline...")
        pipeline = VNextPipeline(
            llm_client=llm_client,
            base_audit_dir=audit_dir,
            visible=False,  # Set to True to see Word application
            progress_callback=progress_callback
        )
        
        print(f"3. Processing document: {docx_path}")
        print(f"   User intent: {user_intent}")
        print()
        
        # Process document
        result = pipeline.process_document(docx_path, user_intent)
        
        print()
        print("=== Processing Results ===")
        print(f"Status: {result.status}")
        print(f"Message: {result.message}")
        
        if result.audit_directory:
            print(f"Audit directory: {result.audit_directory}")
        
        if result.errors:
            print("Errors:")
            for error in result.errors:
                print(f"  - {error}")
        
        if result.warnings:
            print("Warnings:")
            for warning in result.warnings:
                print(f"  - {warning}")
        
        # Show audit trail contents
        if result.audit_directory and os.path.exists(result.audit_directory):
            print()
            print("=== Audit Trail Contents ===")
            audit_files = os.listdir(result.audit_directory)
            for file in sorted(audit_files):
                file_path = os.path.join(result.audit_directory, file)
                if os.path.isfile(file_path):
                    size = os.path.getsize(file_path)
                    print(f"  {file} ({size} bytes)")
        
        print()
        if result.status == "SUCCESS":
            print("✅ Document processing completed successfully!")
        elif result.status == "FAILED_VALIDATION":
            print("⚠️  Document processing failed validation - changes rolled back")
        elif result.status == "ROLLBACK":
            print("❌ Document processing failed - rollback performed")
        elif result.status == "INVALID_PLAN":
            print("❌ LLM generated invalid plan - processing aborted")
        
    except Exception as e:
        logger.error(f"Demo failed: {str(e)}", exc_info=True)
        print(f"❌ Demo failed: {str(e)}")
        return 1
    
    return 0


def demo_with_custom_progress():
    """Demo with custom progress tracking."""
    print("\n=== Custom Progress Demo ===")
    
    progress_stages = []
    
    def custom_progress_callback(stage_name: str, progress_percent: int):
        """Custom progress callback that tracks all stages."""
        progress_stages.append((stage_name, progress_percent))
        print(f"📊 Progress: {stage_name} - {progress_percent}%")
    
    # This would be a real implementation with mocked components for demo
    print("This demo shows how progress tracking works throughout the pipeline:")
    print()
    
    # Simulate progress updates
    stages = ["Extract", "Plan", "Execute", "Validate", "Audit"]
    for i, stage in enumerate(stages):
        progress = int((i / len(stages)) * 100)
        custom_progress_callback(stage, progress)
    
    # Final completion
    custom_progress_callback("Audit", 100)
    
    print()
    print(f"Total progress updates: {len(progress_stages)}")


def demo_error_handling():
    """Demo error handling scenarios."""
    print("\n=== Error Handling Demo ===")
    print("This demo shows how the pipeline handles various error scenarios:")
    print()
    
    error_scenarios = [
        ("ExtractionError", "Document file corrupted or inaccessible"),
        ("PlanningError", "LLM generated invalid plan or API failure"),
        ("ExecutionError", "Word COM operation failed during execution"),
        ("ValidationError", "Document validation assertions failed"),
        ("AuditError", "Failed to create audit trail")
    ]
    
    for error_type, description in error_scenarios:
        print(f"• {error_type}: {description}")
        print(f"  → Pipeline performs rollback and creates audit trail")
    
    print()
    print("All errors result in:")
    print("  1. Automatic rollback to original document")
    print("  2. Complete audit trail with error details")
    print("  3. Structured ProcessingResult with error information")


if __name__ == "__main__":
    exit_code = main()
    
    # Additional demos
    demo_with_custom_progress()
    demo_error_handling()
    
    sys.exit(exit_code)