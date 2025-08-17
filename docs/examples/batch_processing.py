#!/usr/bin/env python3
"""
AutoWord vNext - Batch Processing Examples

This script demonstrates batch processing capabilities for AutoWord vNext,
including processing multiple documents, handling errors, and generating reports.
"""

import os
import sys
import glob
import time
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import List, Dict, Any

# Add autoword to path for examples
sys.path.insert(0, str(Path(__file__).parent.parent.parent))

from autoword.vnext import VNextPipeline
from autoword.vnext.core import VNextConfig, LLMConfig, AuditConfig


class BatchProcessor:
    """Batch processing manager for AutoWord vNext"""
    
    def __init__(self, config: VNextConfig = None):
        self.pipeline = VNextPipeline(config)
        self.results = []
        
    def process_single_document(self, doc_path: str, user_intent: str) -> Dict[str, Any]:
        """Process a single document and return result summary"""
        start_time = time.time()
        
        try:
            result = self.pipeline.process_document(doc_path, user_intent)
            
            return {
                'document': doc_path,
                'status': result.status,
                'success': result.status == "SUCCESS",
                'output_path': result.output_path,
                'audit_directory': result.audit_directory,
                'error': result.error,
                'validation_errors': result.validation_errors,
                'processing_time': time.time() - start_time
            }
            
        except Exception as e:
            return {
                'document': doc_path,
                'status': 'EXCEPTION',
                'success': False,
                'error': str(e),
                'processing_time': time.time() - start_time
            }
    
    def process_batch(self, document_paths: List[str], user_intent: str) -> List[Dict[str, Any]]:
        """Process multiple documents sequentially"""
        results = []
        
        print(f"Processing {len(document_paths)} documents...")
        
        for i, doc_path in enumerate(document_paths, 1):
            print(f"[{i}/{len(document_paths)}] Processing: {doc_path}")
            
            result = self.process_single_document(doc_path, user_intent)
            results.append(result)
            
            # Print immediate result
            if result['success']:
                print(f"  ✅ Success ({result['processing_time']:.1f}s)")
            else:
                print(f"  ❌ Failed: {result['status']}")
                if result.get('error'):
                    print(f"     Error: {result['error']}")
        
        self.results.extend(results)
        return results
    
    def generate_batch_report(self, results: List[Dict[str, Any]]) -> Dict[str, Any]:
        """Generate summary report for batch processing"""
        total_docs = len(results)
        successful = sum(1 for r in results if r['success'])
        failed = total_docs - successful
        
        total_time = sum(r['processing_time'] for r in results)
        avg_time = total_time / total_docs if total_docs > 0 else 0
        
        # Group failures by status
        failure_types = {}
        for result in results:
            if not result['success']:
                status = result['status']
                failure_types[status] = failure_types.get(status, 0) + 1
        
        return {
            'total_documents': total_docs,
            'successful': successful,
            'failed': failed,
            'success_rate': (successful / total_docs * 100) if total_docs > 0 else 0,
            'total_time': total_time,
            'average_time': avg_time,
            'failure_types': failure_types,
            'results': results
        }


def simple_batch_example():
    """Simple batch processing example"""
    print("=== Simple Batch Processing ===")
    
    # Find all DOCX files in current directory
    document_paths = glob.glob("*.docx")
    
    if not document_paths:
        print("No DOCX files found in current directory")
        # Create sample file list for demonstration
        document_paths = ["doc1.docx", "doc2.docx", "doc3.docx"]
        print(f"Using sample paths: {document_paths}")
    
    # Initialize batch processor
    processor = BatchProcessor()
    
    # Define common user intent
    user_intent = "Delete abstract and references sections, update table of contents"
    
    # Process batch
    results = processor.process_batch(document_paths, user_intent)
    
    # Generate report
    report = processor.generate_batch_report(results)
    
    print(f"\nBatch Processing Report:")
    print(f"  Total documents: {report['total_documents']}")
    print(f"  Successful: {report['successful']}")
    print(f"  Failed: {report['failed']}")
    print(f"  Success rate: {report['success_rate']:.1f}%")
    print(f"  Total time: {report['total_time']:.1f}s")
    print(f"  Average time: {report['average_time']:.1f}s per document")
    
    if report['failure_types']:
        print(f"  Failure types:")
        for status, count in report['failure_types'].items():
            print(f"    {status}: {count}")


def academic_papers_batch():
    """Batch process academic papers with specific formatting"""
    print("\n=== Academic Papers Batch Processing ===")
    
    # Configuration for academic papers
    config = VNextConfig(
        llm=LLMConfig(
            provider="openai",
            model="gpt-4",
            temperature=0.1
        ),
        audit=AuditConfig(
            save_snapshots=True,
            generate_diff_reports=True,
            output_directory="./academic_batch_audit"
        )
    )
    
    processor = BatchProcessor(config)
    
    # Academic paper specific intent
    user_intent = """
    Process academic paper:
    1. Delete abstract section (摘要 or Abstract)
    2. Delete references section (参考文献 or References)
    3. Update table of contents
    4. Set Heading 1 to 楷体, 12pt, bold, 2.0 line spacing
    5. Set Heading 2 to 宋体, 12pt, bold, 2.0 line spacing
    6. Set Normal text to 宋体, 12pt, 2.0 line spacing
    """
    
    # Sample academic papers
    academic_papers = [
        "research_paper_1.docx",
        "thesis_chapter_2.docx", 
        "conference_paper_3.docx"
    ]
    
    print(f"Processing {len(academic_papers)} academic papers...")
    
    results = processor.process_batch(academic_papers, user_intent)
    report = processor.generate_batch_report(results)
    
    print(f"\nAcademic Papers Report:")
    print(f"  Success rate: {report['success_rate']:.1f}%")
    print(f"  Average processing time: {report['average_time']:.1f}s")
    
    # Detailed results
    for result in results:
        doc_name = Path(result['document']).name
        if result['success']:
            print(f"  ✅ {doc_name}: Success")
        else:
            print(f"  ❌ {doc_name}: {result['status']}")


def error_handling_batch():
    """Demonstrate error handling in batch processing"""
    print("\n=== Error Handling in Batch Processing ===")
    
    processor = BatchProcessor()
    
    # Mix of valid and problematic documents
    test_documents = [
        "valid_document.docx",
        "locked_document.docx",      # Simulated locked file
        "corrupted_document.docx",   # Simulated corrupted file
        "missing_document.docx",     # Non-existent file
        "another_valid_doc.docx"
    ]
    
    user_intent = "Update document formatting and TOC"
    
    results = []
    
    for doc_path in test_documents:
        print(f"Processing: {doc_path}")
        
        # Simulate different error conditions
        if "locked" in doc_path:
            result = {
                'document': doc_path,
                'status': 'EXCEPTION',
                'success': False,
                'error': 'Document is locked for editing',
                'processing_time': 0.5
            }
        elif "corrupted" in doc_path:
            result = {
                'document': doc_path,
                'status': 'FAILED_VALIDATION',
                'success': False,
                'validation_errors': ['Document structure corrupted'],
                'processing_time': 2.1
            }
        elif "missing" in doc_path:
            result = {
                'document': doc_path,
                'status': 'EXCEPTION',
                'success': False,
                'error': 'File not found',
                'processing_time': 0.1
            }
        else:
            # Simulate successful processing
            result = {
                'document': doc_path,
                'status': 'SUCCESS',
                'success': True,
                'output_path': f"output/{doc_path}",
                'processing_time': 3.2
            }
        
        results.append(result)
        
        # Show immediate result
        if result['success']:
            print(f"  ✅ Success")
        else:
            print(f"  ❌ {result['status']}: {result.get('error', 'Unknown error')}")
    
    # Generate error analysis
    report = processor.generate_batch_report(results)
    
    print(f"\nError Handling Report:")
    print(f"  Total: {report['total_documents']}")
    print(f"  Success: {report['successful']}")
    print(f"  Failed: {report['failed']}")
    
    print(f"\nError breakdown:")
    for status, count in report['failure_types'].items():
        print(f"  {status}: {count} documents")


def custom_batch_with_different_intents():
    """Process documents with different user intents"""
    print("\n=== Custom Batch with Different Intents ===")
    
    processor = BatchProcessor()
    
    # Documents with specific processing requirements
    document_configs = [
        {
            'path': 'report_template.docx',
            'intent': 'Update TOC and standardize heading fonts to 楷体'
        },
        {
            'path': 'meeting_minutes.docx', 
            'intent': 'Remove draft watermark and update page numbers'
        },
        {
            'path': 'technical_spec.docx',
            'intent': 'Delete appendix sections and update all cross-references'
        },
        {
            'path': 'user_manual.docx',
            'intent': 'Apply consistent formatting: Heading 1 bold, Normal 12pt'
        }
    ]
    
    results = []
    
    for config in document_configs:
        doc_path = config['path']
        user_intent = config['intent']
        
        print(f"Processing: {doc_path}")
        print(f"  Intent: {user_intent}")
        
        result = processor.process_single_document(doc_path, user_intent)
        results.append(result)
        
        if result['success']:
            print(f"  ✅ Success ({result['processing_time']:.1f}s)")
        else:
            print(f"  ❌ Failed: {result['status']}")
    
    # Summary
    report = processor.generate_batch_report(results)
    print(f"\nCustom Batch Summary:")
    print(f"  Success rate: {report['success_rate']:.1f}%")
    print(f"  Total time: {report['total_time']:.1f}s")


def save_batch_report(report: Dict[str, Any], output_file: str):
    """Save batch processing report to file"""
    import json
    from datetime import datetime
    
    # Add timestamp to report
    report['timestamp'] = datetime.now().isoformat()
    report['report_version'] = '1.0'
    
    # Save as JSON
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(report, f, indent=2, ensure_ascii=False)
    
    print(f"Report saved to: {output_file}")


def main():
    """Run all batch processing examples"""
    print("AutoWord vNext - Batch Processing Examples")
    print("=" * 60)
    
    try:
        # Run examples
        simple_batch_example()
        academic_papers_batch()
        error_handling_batch()
        custom_batch_with_different_intents()
        
        print("\n" + "=" * 60)
        print("Batch processing examples completed.")
        print("Check the User Guide for more batch processing options.")
        
    except ImportError as e:
        print(f"❌ Import error: {e}")
        print("Please ensure AutoWord vNext is properly installed.")
        
    except Exception as e:
        print(f"❌ Unexpected error: {e}")


if __name__ == "__main__":
    main()