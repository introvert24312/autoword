#!/usr/bin/env python3
"""
AutoWord vNext - Error Handling Examples

This script demonstrates comprehensive error handling patterns for AutoWord vNext,
including exception handling, validation failures, rollback scenarios, and recovery strategies.
"""

import os
import sys
import time
import logging
from pathlib import Path
from typing import Optional, Dict, Any

# Add autoword to path for examples
sys.path.insert(0, str(Path(__file__).parent.parent.parent))

from autoword.vnext import VNextPipeline
from autoword.vnext.core import VNextConfig, LLMConfig, ValidationConfig
from autoword.vnext.exceptions import (
    VNextError, ExtractionError, PlanningError, 
    ExecutionError, ValidationError, AuditError
)


def setup_logging():
    """Setup logging for error handling examples"""
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.StreamHandler(),
            logging.FileHandler('error_handling_examples.log')
        ]
    )
    return logging.getLogger(__name__)


def basic_exception_handling():
    """Basic exception handling patterns"""
    print("=== Basic Exception Handling ===")
    
    pipeline = VNextPipeline()
    logger = logging.getLogger(__name__)
    
    # Test with non-existent file
    try:
        result = pipeline.process_document(
            "non_existent_file.docx",
            "Delete abstract section"
        )
        
    except FileNotFoundError as e:
        print(f"❌ File not found: {e}")
        logger.error(f"File not found: {e}")
        
    except PermissionError as e:
        print(f"❌ Permission denied: {e}")
        logger.error(f"Permission denied: {e}")
        
    except VNextError as e:
        print(f"❌ VNext pipeline error: {e}")
        logger.error(f"VNext error: {e}")
        
    except Exception as e:
        print(f"❌ Unexpected error: {e}")
        logger.error(f"Unexpected error: {e}")
        
    print("✅ Basic exception handling demonstrated")


def specific_exception_handling():
    """Handle specific VNext exception types"""
    print("\n=== Specific Exception Handling ===")
    
    pipeline = VNextPipeline()
    logger = logging.getLogger(__name__)
    
    test_cases = [
        {
            'file': 'corrupted_document.docx',
            'intent': 'Delete abstract section',
            'expected_error': ExtractionError
        },
        {
            'file': 'valid_document.docx', 
            'intent': 'Invalid operation request',
            'expected_error': PlanningError
        },
        {
            'file': 'locked_document.docx',
            'intent': 'Update TOC',
            'expected_error': ExecutionError
        }
    ]
    
    for test_case in test_cases:
        print(f"\nTesting: {test_case['file']}")
        
        try:
            result = pipeline.process_document(
                test_case['file'],
                test_case['intent']
            )
            
            print(f"✅ Unexpected success: {result.status}")
            
        except ExtractionError as e:
            print(f"📤 Extraction error: {e}")
            logger.warning(f"Extraction failed for {test_case['file']}: {e}")
            
        except PlanningError as e:
            print(f"🤖 Planning error: {e}")
            logger.warning(f"Planning failed for {test_case['file']}: {e}")
            
        except ExecutionError as e:
            print(f"⚡ Execution error: {e}")
            logger.warning(f"Execution failed for {test_case['file']}: {e}")
            
        except ValidationError as e:
            print(f"✅ Validation error: {e}")
            print(f"   Assertion failures: {e.assertion_failures}")
            print(f"   Rollback path: {e.rollback_path}")
            logger.warning(f"Validation failed for {test_case['file']}: {e}")
            
        except VNextError as e:
            print(f"❌ General VNext error: {e}")
            logger.error(f"VNext error for {test_case['file']}: {e}")


def validation_failure_handling():
    """Handle validation failures and rollback scenarios"""
    print("\n=== Validation Failure Handling ===")
    
    # Configure with strict validation
    config = VNextConfig(
        validation=ValidationConfig(
            strict_mode=True,
            rollback_on_failure=True,
            forbidden_level1_headings=["摘要", "参考文献", "Abstract", "References"]
        )
    )
    
    pipeline = VNextPipeline(config)
    logger = logging.getLogger(__name__)
    
    # Simulate document that would fail validation
    test_document = "document_with_forbidden_headings.docx"
    user_intent = "Update document formatting"
    
    try:
        result = pipeline.process_document(test_document, user_intent)
        
        if result.status == "FAILED_VALIDATION":
            print("❌ Validation failed as expected")
            print(f"   Status: {result.status}")
            
            if result.validation_errors:
                print("   Validation errors:")
                for error in result.validation_errors:
                    print(f"     - {error}")
            
            # Check if rollback occurred
            if hasattr(result, 'rollback_path') and result.rollback_path:
                print(f"   Document rolled back to: {result.rollback_path}")
                
        elif result.status == "SUCCESS":
            print("✅ Unexpected success - validation should have failed")
            
    except ValidationError as e:
        print("❌ Validation exception caught")
        print(f"   Assertion failures: {e.assertion_failures}")
        print(f"   Rollback path: {e.rollback_path}")
        
        # Log detailed validation failure
        logger.error(f"Validation failed: {e.assertion_failures}")
        
        # Check if rollback file exists
        if e.rollback_path and Path(e.rollback_path).exists():
            print(f"✅ Rollback file confirmed: {e.rollback_path}")
        else:
            print(f"❌ Rollback file missing: {e.rollback_path}")


def retry_mechanism_example():
    """Implement retry mechanism for transient errors"""
    print("\n=== Retry Mechanism Example ===")
    
    pipeline = VNextPipeline()
    logger = logging.getLogger(__name__)
    
    def process_with_retry(doc_path: str, user_intent: str, max_retries: int = 3) -> Optional[Dict[str, Any]]:
        """Process document with retry logic"""
        
        for attempt in range(max_retries):
            try:
                print(f"Attempt {attempt + 1}/{max_retries}: {doc_path}")
                
                result = pipeline.process_document(doc_path, user_intent)
                
                if result.status == "SUCCESS":
                    print(f"✅ Success on attempt {attempt + 1}")
                    return {
                        'success': True,
                        'result': result,
                        'attempts': attempt + 1
                    }
                elif result.status in ["FAILED_VALIDATION", "INVALID_PLAN"]:
                    # Don't retry validation or plan failures
                    print(f"❌ Non-retryable failure: {result.status}")
                    return {
                        'success': False,
                        'result': result,
                        'attempts': attempt + 1,
                        'reason': 'non_retryable'
                    }
                else:
                    print(f"⚠️  Retryable failure: {result.status}")
                    if attempt < max_retries - 1:
                        time.sleep(2 ** attempt)  # Exponential backoff
                        
            except (ExecutionError, ExtractionError) as e:
                print(f"⚠️  Retryable exception: {e}")
                logger.warning(f"Attempt {attempt + 1} failed: {e}")
                
                if attempt < max_retries - 1:
                    time.sleep(2 ** attempt)
                else:
                    return {
                        'success': False,
                        'error': str(e),
                        'attempts': attempt + 1,
                        'reason': 'max_retries_exceeded'
                    }
                    
            except (PlanningError, ValidationError) as e:
                # Don't retry these errors
                print(f"❌ Non-retryable exception: {e}")
                return {
                    'success': False,
                    'error': str(e),
                    'attempts': attempt + 1,
                    'reason': 'non_retryable'
                }
        
        return {
            'success': False,
            'attempts': max_retries,
            'reason': 'max_retries_exceeded'
        }
    
    # Test retry mechanism
    test_documents = [
        "transient_error_doc.docx",
        "permanent_error_doc.docx",
        "valid_document.docx"
    ]
    
    for doc in test_documents:
        print(f"\nTesting retry with: {doc}")
        result = process_with_retry(doc, "Update TOC and formatting", max_retries=3)
        
        if result['success']:
            print(f"✅ Success after {result['attempts']} attempts")
        else:
            print(f"❌ Failed after {result['attempts']} attempts")
            print(f"   Reason: {result.get('reason', 'unknown')}")


def graceful_degradation_example():
    """Demonstrate graceful degradation strategies"""
    print("\n=== Graceful Degradation Example ===")
    
    pipeline = VNextPipeline()
    logger = logging.getLogger(__name__)
    
    def process_with_fallback(doc_path: str, primary_intent: str, fallback_intent: str) -> Dict[str, Any]:
        """Process with fallback intent if primary fails"""
        
        # Try primary intent
        try:
            print(f"Trying primary intent: {primary_intent}")
            result = pipeline.process_document(doc_path, primary_intent)
            
            if result.status == "SUCCESS":
                return {
                    'success': True,
                    'result': result,
                    'intent_used': 'primary'
                }
            else:
                print(f"Primary intent failed: {result.status}")
                
        except Exception as e:
            print(f"Primary intent exception: {e}")
            logger.warning(f"Primary intent failed: {e}")
        
        # Try fallback intent
        try:
            print(f"Trying fallback intent: {fallback_intent}")
            result = pipeline.process_document(doc_path, fallback_intent)
            
            return {
                'success': result.status == "SUCCESS",
                'result': result,
                'intent_used': 'fallback'
            }
            
        except Exception as e:
            print(f"Fallback intent exception: {e}")
            logger.error(f"Both intents failed: {e}")
            
            return {
                'success': False,
                'error': str(e),
                'intent_used': 'none'
            }
    
    # Test graceful degradation
    test_case = {
        'document': 'complex_document.docx',
        'primary_intent': 'Delete abstract, references, and appendices; update TOC; apply complex formatting',
        'fallback_intent': 'Update table of contents only'
    }
    
    print(f"Document: {test_case['document']}")
    result = process_with_fallback(
        test_case['document'],
        test_case['primary_intent'],
        test_case['fallback_intent']
    )
    
    if result['success']:
        print(f"✅ Success using {result['intent_used']} intent")
    else:
        print(f"❌ Both intents failed")


def error_reporting_and_logging():
    """Comprehensive error reporting and logging"""
    print("\n=== Error Reporting and Logging ===")
    
    # Setup detailed logging
    error_logger = logging.getLogger('vnext_errors')
    error_handler = logging.FileHandler('vnext_detailed_errors.log')
    error_handler.setFormatter(
        logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    )
    error_logger.addHandler(error_handler)
    error_logger.setLevel(logging.DEBUG)
    
    pipeline = VNextPipeline()
    
    def detailed_error_report(doc_path: str, user_intent: str) -> Dict[str, Any]:
        """Generate detailed error report"""
        
        error_report = {
            'document': doc_path,
            'user_intent': user_intent,
            'timestamp': time.time(),
            'errors': [],
            'warnings': [],
            'system_info': {
                'python_version': sys.version,
                'platform': sys.platform,
                'working_directory': os.getcwd()
            }
        }
        
        try:
            result = pipeline.process_document(doc_path, user_intent)
            
            error_report['status'] = result.status
            error_report['success'] = result.status == "SUCCESS"
            
            if result.status != "SUCCESS":
                if result.error:
                    error_report['errors'].append(result.error)
                if result.validation_errors:
                    error_report['errors'].extend(result.validation_errors)
                if result.plan_errors:
                    error_report['errors'].extend(result.plan_errors)
            
            # Log detailed information
            error_logger.info(f"Processing attempt: {doc_path}")
            error_logger.info(f"Status: {result.status}")
            
            if error_report['errors']:
                for error in error_report['errors']:
                    error_logger.error(f"Error: {error}")
            
        except Exception as e:
            error_report['status'] = 'EXCEPTION'
            error_report['success'] = False
            error_report['errors'].append(str(e))
            error_report['exception_type'] = type(e).__name__
            
            error_logger.error(f"Exception in {doc_path}: {e}", exc_info=True)
        
        return error_report
    
    # Generate error reports for test cases
    test_documents = [
        "valid_document.docx",
        "problematic_document.docx",
        "missing_document.docx"
    ]
    
    reports = []
    for doc in test_documents:
        print(f"Generating error report for: {doc}")
        report = detailed_error_report(doc, "Update document formatting")
        reports.append(report)
        
        if report['success']:
            print(f"✅ Success")
        else:
            print(f"❌ Failed: {report['status']}")
            for error in report['errors']:
                print(f"   - {error}")
    
    # Save consolidated error report
    import json
    with open('error_reports.json', 'w') as f:
        json.dump(reports, f, indent=2, default=str)
    
    print(f"✅ Error reports saved to: error_reports.json")


def recovery_strategies_example():
    """Demonstrate various recovery strategies"""
    print("\n=== Recovery Strategies Example ===")
    
    pipeline = VNextPipeline()
    logger = logging.getLogger(__name__)
    
    def smart_recovery_process(doc_path: str, user_intent: str) -> Dict[str, Any]:
        """Implement smart recovery strategies"""
        
        recovery_strategies = [
            {
                'name': 'original_intent',
                'intent': user_intent,
                'description': 'Original user intent'
            },
            {
                'name': 'simplified_intent',
                'intent': 'Update table of contents only',
                'description': 'Simplified operation'
            },
            {
                'name': 'minimal_intent',
                'intent': 'Validate document structure',
                'description': 'Minimal validation only'
            }
        ]
        
        for strategy in recovery_strategies:
            try:
                print(f"Trying strategy: {strategy['name']}")
                print(f"  Description: {strategy['description']}")
                
                result = pipeline.process_document(doc_path, strategy['intent'])
                
                if result.status == "SUCCESS":
                    return {
                        'success': True,
                        'strategy_used': strategy['name'],
                        'result': result,
                        'description': strategy['description']
                    }
                else:
                    print(f"  Strategy failed: {result.status}")
                    logger.warning(f"Strategy {strategy['name']} failed: {result.status}")
                    
            except Exception as e:
                print(f"  Strategy exception: {e}")
                logger.warning(f"Strategy {strategy['name']} exception: {e}")
        
        return {
            'success': False,
            'strategy_used': 'none',
            'error': 'All recovery strategies failed'
        }
    
    # Test recovery strategies
    test_document = "challenging_document.docx"
    complex_intent = "Delete multiple sections, update TOC, apply complex formatting, and validate everything"
    
    print(f"Document: {test_document}")
    print(f"Original intent: {complex_intent}")
    
    result = smart_recovery_process(test_document, complex_intent)
    
    if result['success']:
        print(f"✅ Recovery successful using: {result['strategy_used']}")
        print(f"   Description: {result['description']}")
    else:
        print(f"❌ All recovery strategies failed")


def main():
    """Run all error handling examples"""
    print("AutoWord vNext - Error Handling Examples")
    print("=" * 60)
    
    # Setup logging
    logger = setup_logging()
    logger.info("Starting error handling examples")
    
    try:
        # Run error handling examples
        basic_exception_handling()
        specific_exception_handling()
        validation_failure_handling()
        retry_mechanism_example()
        graceful_degradation_example()
        error_reporting_and_logging()
        recovery_strategies_example()
        
        print("\n" + "=" * 60)
        print("Error handling examples completed.")
        print("Check the log files for detailed error information.")
        
        logger.info("Error handling examples completed successfully")
        
    except ImportError as e:
        print(f"❌ Import error: {e}")
        print("Please ensure AutoWord vNext is properly installed.")
        logger.error(f"Import error: {e}")
        
    except Exception as e:
        print(f"❌ Unexpected error in examples: {e}")
        logger.error(f"Unexpected error: {e}", exc_info=True)


if __name__ == "__main__":
    main()