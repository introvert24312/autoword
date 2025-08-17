"""
Command-line interface for VNext Pipeline.

This module provides a CLI for running the VNext pipeline with various options
for document processing, batch operations, and debugging.
"""

import os
import sys
import argparse
import logging
import json
from datetime import datetime
from pathlib import Path
from typing import Optional, List, Dict, Any

from .pipeline import VNextPipeline
from .models import ProcessingResult
from .monitoring import MonitoringLevel
from ..core.llm_client import LLMClient, ModelType


def setup_logging(verbose: bool = False, log_file: Optional[str] = None):
    """Setup logging configuration."""
    level = logging.DEBUG if verbose else logging.INFO
    
    handlers = [logging.StreamHandler()]
    if log_file:
        handlers.append(logging.FileHandler(log_file))
    
    logging.basicConfig(
        level=level,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=handlers
    )


def load_config(config_path: str) -> Dict[str, Any]:
    """Load configuration from JSON file."""
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)
        print(f"Loaded configuration from: {config_path}")
        return config
    except FileNotFoundError:
        print(f"[ERROR] Configuration file not found: {config_path}")
        sys.exit(1)
    except json.JSONDecodeError as e:
        print(f"[ERROR] Invalid JSON in configuration file: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"[ERROR] Error loading configuration: {e}")
        sys.exit(1)


def save_config(args, config_path: str):
    """Save current configuration to JSON file."""
    config = {
        "model": args.model,
        "temperature": args.temperature,
        "audit_dir": args.audit_dir,
        "visible": args.visible,
        "verbose": args.verbose,
        "monitoring_level": args.monitoring_level,
        "enable_memory_monitoring": args.enable_memory_monitoring,
        "memory_warning_threshold": args.memory_warning_threshold,
        "memory_critical_threshold": args.memory_critical_threshold,
        "log_file": args.log_file
    }
    
    try:
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
        print(f"[OK] Configuration saved to: {config_path}")
    except Exception as e:
        print(f"[ERROR] Error saving configuration: {e}")
        sys.exit(1)


def apply_config(args, config: Dict[str, Any]):
    """Apply configuration values to arguments, with CLI args taking precedence."""
    # Only apply config values if CLI argument is at default value
    defaults = {
        'model': None,
        'temperature': 0.1,
        'audit_dir': './audit_trails',
        'visible': False,
        'verbose': False,
        'monitoring_level': 'detailed',
        'enable_memory_monitoring': True,
        'memory_warning_threshold': 1024,
        'memory_critical_threshold': 2048,
        'log_file': None
    }
    
    for key, default_value in defaults.items():
        if hasattr(args, key) and getattr(args, key) == default_value and key in config:
            setattr(args, key, config[key])
            print(f"Applied config: {key} = {config[key]}")


def show_performance_summary(result: ProcessingResult):
    """Show performance summary if available."""
    if hasattr(result, 'performance_metrics') and result.performance_metrics:
        print("\n=== Performance Summary ===")
        metrics = result.performance_metrics
        if hasattr(metrics, 'total_duration_ms'):
            print(f"Total processing time: {metrics.total_duration_ms:.1f}ms")
        if hasattr(metrics, 'peak_memory_mb'):
            print(f"Peak memory usage: {metrics.peak_memory_mb:.1f}MB")
        if hasattr(metrics, 'stage_durations'):
            print("Stage durations:")
            for stage, duration in metrics.stage_durations.items():
                print(f"  {stage}: {duration:.1f}ms")
    
    if hasattr(result, 'memory_alerts') and result.memory_alerts:
        print(f"\n[WARNING] Memory alerts: {len(result.memory_alerts)}")
        for alert in result.memory_alerts[-3:]:  # Show last 3 alerts
            print(f"  {alert.alert_level}: {alert.message}")


def progress_callback(stage_name: str, progress_percent: int):
    """Progress callback for CLI output."""
    print(f"[{progress_percent:3d}%] {stage_name}")


def process_single_document(args) -> int:
    """Process a single document."""
    print(f"Processing document: {args.input}")
    print(f"User intent: {args.intent}")
    
    try:
        # Initialize LLM client
        llm_client = None
        if args.model:
            model_type = ModelType[args.model.upper()]
            llm_client = LLMClient(
                model=model_type,
                api_key=os.getenv("OPENAI_API_KEY") or os.getenv("ANTHROPIC_API_KEY"),
                temperature=args.temperature
            )
        
        # Initialize pipeline with monitoring configuration
        monitoring_level = MonitoringLevel[args.monitoring_level.upper()]
        pipeline = VNextPipeline(
            llm_client=llm_client,
            base_audit_dir=args.audit_dir,
            visible=args.visible,
            progress_callback=progress_callback if args.verbose else None,
            monitoring_level=monitoring_level,
            enable_memory_monitoring=args.enable_memory_monitoring,
            memory_warning_threshold_mb=args.memory_warning_threshold,
            memory_critical_threshold_mb=args.memory_critical_threshold
        )
        
        # Process document
        result = pipeline.process_document(args.input, args.intent)
        
        # Output results
        print(f"\nStatus: {result.status}")
        if result.message:
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
        
        # Show performance summary in verbose mode
        if args.verbose:
            show_performance_summary(result)
        
        # Return appropriate exit code
        if result.status == "SUCCESS":
            print("[SUCCESS] Document processing completed successfully!")
            return 0
        elif result.status == "FAILED_VALIDATION":
            print("[WARNING] Document processing failed validation - changes rolled back")
            return 2
        elif result.status == "ROLLBACK":
            print("[ERROR] Document processing failed - rollback performed")
            return 3
        elif result.status == "INVALID_PLAN":
            print("[ERROR] LLM generated invalid plan - processing aborted")
            return 4
        else:
            print(f"[UNKNOWN] Unknown status: {result.status}")
            return 5
            
    except Exception as e:
        print(f"[ERROR] Processing failed: {str(e)}")
        if args.verbose:
            import traceback
            traceback.print_exc()
        return 1


def process_batch_documents(args) -> int:
    """Process multiple documents in batch."""
    print(f"Batch processing documents from: {args.batch_dir}")
    
    # Find all DOCX files
    batch_dir = Path(args.batch_dir)
    docx_files = list(batch_dir.glob("*.docx"))
    
    if not docx_files:
        print("No DOCX files found in batch directory")
        return 1
    
    print(f"Found {len(docx_files)} DOCX files")
    
    results = []
    failed_count = 0
    
    try:
        # Initialize LLM client
        llm_client = None
        if args.model:
            model_type = ModelType[args.model.upper()]
            llm_client = LLMClient(
                model=model_type,
                api_key=os.getenv("OPENAI_API_KEY") or os.getenv("ANTHROPIC_API_KEY"),
                temperature=args.temperature
            )
        
        # Process each document
        for i, docx_file in enumerate(docx_files, 1):
            print(f"\n[{i}/{len(docx_files)}] Processing: {docx_file.name}")
            
            # Initialize pipeline for each document with monitoring configuration
            monitoring_level = MonitoringLevel[args.monitoring_level.upper()]
            pipeline = VNextPipeline(
                llm_client=llm_client,
                base_audit_dir=args.audit_dir,
                visible=args.visible,
                progress_callback=progress_callback if args.verbose else None,
                monitoring_level=monitoring_level,
                enable_memory_monitoring=args.enable_memory_monitoring,
                memory_warning_threshold_mb=args.memory_warning_threshold,
                memory_critical_threshold_mb=args.memory_critical_threshold
            )
            
            # Process document
            result = pipeline.process_document(str(docx_file), args.intent)
            results.append((docx_file.name, result))
            
            if result.status != "SUCCESS":
                failed_count += 1
                print(f"  [FAILED] {result.status}")
                if result.errors:
                    for error in result.errors[:2]:  # Show first 2 errors
                        print(f"    - {error}")
            else:
                print(f"  [SUCCESS]")
        
        # Summary
        print(f"\n=== Batch Processing Summary ===")
        print(f"Total documents: {len(docx_files)}")
        print(f"Successful: {len(docx_files) - failed_count}")
        print(f"Failed: {failed_count}")
        
        # Show detailed results
        if args.verbose or failed_count > 0:
            print("\nDetailed Results:")
            for filename, result in results:
                status_icon = "[OK]" if result.status == "SUCCESS" else "[FAIL]"
                print(f"  {status_icon} {filename}: {result.status}")
                if result.status != "SUCCESS" and result.errors:
                    for error in result.errors[:1]:  # Show first error
                        print(f"    Error: {error}")
                if args.verbose and result.audit_directory:
                    print(f"    Audit: {result.audit_directory}")
        
        # Create batch summary report
        if args.audit_dir:
            try:
                summary_file = os.path.join(args.audit_dir, f"batch_summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json")
                summary_data = {
                    "timestamp": datetime.now().isoformat(),
                    "total_documents": len(docx_files),
                    "successful": len(docx_files) - failed_count,
                    "failed": failed_count,
                    "user_intent": args.intent,
                    "results": [
                        {
                            "filename": filename,
                            "status": result.status,
                            "audit_directory": result.audit_directory,
                            "error_count": len(result.errors) if result.errors else 0
                        }
                        for filename, result in results
                    ]
                }
                
                os.makedirs(args.audit_dir, exist_ok=True)
                with open(summary_file, 'w', encoding='utf-8') as f:
                    json.dump(summary_data, f, indent=2, ensure_ascii=False)
                print(f"\nBatch summary saved to: {summary_file}")
                
            except Exception as e:
                print(f"[WARNING] Could not save batch summary: {e}")
        
        return 0 if failed_count == 0 else 1
        
    except Exception as e:
        print(f"[ERROR] Batch processing failed: {str(e)}")
        if args.verbose:
            import traceback
            traceback.print_exc()
        return 1


def check_system_status(args) -> int:
    """Check system status and requirements."""
    print("=== AutoWord vNext System Status ===\n")
    
    status_ok = True
    
    # Check Python version
    python_version = sys.version_info
    print(f"Python Version: {python_version.major}.{python_version.minor}.{python_version.micro}")
    if python_version < (3, 8):
        print("  [ERROR] Python 3.8+ required")
        status_ok = False
    else:
        print("  [OK] Python version OK")
    
    # Check required modules
    required_modules = [
        ('win32com.client', 'pywin32'),
        ('pydantic', 'pydantic'),
        ('psutil', 'psutil'),
        ('pathlib', 'built-in'),
    ]
    
    print("\nRequired Modules:")
    for module_name, package_name in required_modules:
        try:
            __import__(module_name)
            print(f"  [OK] {package_name}")
        except ImportError:
            print(f"  [ERROR] {package_name} - install with: pip install {package_name}")
            status_ok = False
    
    # Check Word COM availability
    print("\nMicrosoft Word COM:")
    try:
        import win32com.client
        word_app = win32com.client.Dispatch("Word.Application")
        word_app.Visible = False
        version = word_app.Version
        word_app.Quit()
        print(f"  [OK] Microsoft Word {version} available")
    except Exception as e:
        print(f"  [ERROR] Microsoft Word COM not available: {e}")
        status_ok = False
    
    # Check environment variables
    print("\nEnvironment Variables:")
    api_keys = [
        ("OPENAI_API_KEY", "OpenAI API access"),
        ("ANTHROPIC_API_KEY", "Anthropic API access")
    ]
    
    has_api_key = False
    for env_var, description in api_keys:
        if os.getenv(env_var):
            print(f"  [OK] {env_var} configured")
            has_api_key = True
        else:
            print(f"  [WARNING] {env_var} not set ({description})")
    
    if not has_api_key:
        print("  [WARNING] No LLM API keys found - plan generation will fail")
    
    # Check disk space
    print("\nSystem Resources:")
    try:
        import shutil
        total, used, free = shutil.disk_usage(".")
        free_gb = free / (1024**3)
        print(f"  Disk space: {free_gb:.1f}GB free")
        if free_gb < 1.0:
            print("  [WARNING] Low disk space - may affect audit trail storage")
    except Exception:
        print("  [WARNING] Could not check disk space")
    
    # Memory check
    try:
        import psutil
        memory = psutil.virtual_memory()
        available_gb = memory.available / (1024**3)
        print(f"  Available memory: {available_gb:.1f}GB")
        if available_gb < 2.0:
            print("  [WARNING] Low memory - consider reducing memory thresholds")
    except Exception:
        print("  [WARNING] Could not check memory")
    
    # Overall status
    print(f"\n=== Overall Status ===")
    if status_ok:
        print("[SUCCESS] System ready for AutoWord vNext processing")
        return 0
    else:
        print("[ERROR] System has issues that need to be resolved")
        return 1


def handle_config_command(args) -> int:
    """Handle configuration management commands."""
    if args.config_action == "show":
        print("Current Configuration Defaults:")
        print(f"  Model: {args.model or 'auto-detect'}")
        print(f"  Temperature: {args.temperature}")
        print(f"  Audit Directory: {args.audit_dir}")
        print(f"  Visible: {args.visible}")
        print(f"  Verbose: {args.verbose}")
        print(f"  Monitoring Level: {args.monitoring_level}")
        print(f"  Memory Monitoring: {args.enable_memory_monitoring}")
        print(f"  Memory Warning Threshold: {args.memory_warning_threshold}MB")
        print(f"  Memory Critical Threshold: {args.memory_critical_threshold}MB")
        print(f"  Log File: {args.log_file or 'console only'}")
        return 0
        
    elif args.config_action == "create":
        template_config = {
            "model": "gpt4",
            "temperature": 0.1,
            "audit_dir": "./audit_trails",
            "visible": False,
            "verbose": False,
            "monitoring_level": "detailed",
            "enable_memory_monitoring": True,
            "memory_warning_threshold": 1024,
            "memory_critical_threshold": 2048,
            "log_file": None
        }
        
        try:
            with open(args.output, 'w', encoding='utf-8') as f:
                json.dump(template_config, f, indent=2, ensure_ascii=False)
            print(f"[OK] Configuration template created: {args.output}")
            print("Edit the file and use with --config option")
            return 0
        except Exception as e:
            print(f"[ERROR] Error creating configuration template: {e}")
            return 1
    
    else:
        print("Unknown config action")
        return 1


def dry_run_document(args) -> int:
    """Perform dry run (plan generation only)."""
    print(f"Dry run for document: {args.input}")
    print(f"User intent: {args.intent}")
    
    try:
        # Initialize LLM client
        llm_client = None
        if args.model:
            model_type = ModelType[args.model.upper()]
            llm_client = LLMClient(
                model=model_type,
                api_key=os.getenv("OPENAI_API_KEY") or os.getenv("ANTHROPIC_API_KEY"),
                temperature=args.temperature
            )
        
        # Initialize pipeline with monitoring configuration
        monitoring_level = MonitoringLevel[args.monitoring_level.upper()]
        pipeline = VNextPipeline(
            llm_client=llm_client,
            base_audit_dir=args.audit_dir,
            visible=False,  # Never show Word for dry run
            progress_callback=progress_callback if args.verbose else None,
            monitoring_level=monitoring_level,
            enable_memory_monitoring=args.enable_memory_monitoring,
            memory_warning_threshold_mb=args.memory_warning_threshold,
            memory_critical_threshold_mb=args.memory_critical_threshold
        )
        
        # Setup and extract
        pipeline._setup_run_environment(args.input)
        try:
            structure, inventory = pipeline._extract_document()
            plan = pipeline._generate_plan(structure, args.intent)
            
            print(f"\n[SUCCESS] Dry run completed successfully!")
            print(f"Generated plan with {len(plan.ops)} operations:")
            
            for i, op in enumerate(plan.ops, 1):
                op_type = type(op).__name__
                print(f"  {i}. {op_type}")
                if hasattr(op, 'heading_text'):
                    print(f"     - Heading: {op.heading_text}")
                elif hasattr(op, 'target_style_name'):
                    print(f"     - Style: {op.target_style_name}")
            
            return 0
            
        finally:
            pipeline._cleanup_run_environment()
            
    except Exception as e:
        print(f"[ERROR] Dry run failed: {str(e)}")
        if args.verbose:
            import traceback
            traceback.print_exc()
        return 1


def main():
    """Main CLI entry point."""
    parser = argparse.ArgumentParser(
        description="AutoWord vNext Pipeline CLI",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Process single document
  python -m autoword.vnext.cli process document.docx "Remove abstract and references"
  
  # Batch process all DOCX files in directory
  python -m autoword.vnext.cli batch ./documents "Apply standard formatting"
  
  # Dry run to see generated plan
  python -m autoword.vnext.cli dry-run document.docx "Update TOC and styles"
  
  # Process with specific model and verbose output
  python -m autoword.vnext.cli process document.docx "Remove sections" --model gpt4 --verbose
  
  # Process with debug monitoring and custom memory thresholds
  python -m autoword.vnext.cli process document.docx "Format document" --monitoring-level debug --memory-warning-threshold 512
  
  # Batch process with performance monitoring
  python -m autoword.vnext.cli batch ./docs "Standardize" --monitoring-level performance --verbose
  
  # Use configuration file
  python -m autoword.vnext.cli process document.docx "Process" --config config.json
  
  # Save current configuration
  python -m autoword.vnext.cli process document.docx "Test" --model gpt4 --save-config my-config.json
  
  # Check system status
  python -m autoword.vnext.cli status
  
  # Create configuration template
  python -m autoword.vnext.cli config create my-config.json
  
  # Show current configuration
  python -m autoword.vnext.cli config show
        """
    )
    
    # Global options
    parser.add_argument("--verbose", "-v", action="store_true",
                       help="Enable verbose output and progress reporting")
    parser.add_argument("--audit-dir", default="./audit_trails",
                       help="Base directory for audit trails (default: ./audit_trails)")
    parser.add_argument("--model", choices=["gpt4", "gpt35", "claude37", "claude3"],
                       help="LLM model to use for plan generation")
    parser.add_argument("--temperature", type=float, default=0.1,
                       help="LLM temperature (default: 0.1)")
    parser.add_argument("--visible", action="store_true",
                       help="Show Word application windows during processing")
    parser.add_argument("--log-file", help="Log file path")
    
    # Monitoring and performance options
    parser.add_argument("--monitoring-level", choices=["basic", "detailed", "debug", "performance"],
                       default="detailed", help="Monitoring detail level (default: detailed)")
    parser.add_argument("--disable-memory-monitoring", dest="enable_memory_monitoring", 
                       action="store_false", default=True,
                       help="Disable memory usage monitoring")
    parser.add_argument("--memory-warning-threshold", type=float, default=1024,
                       help="Memory warning threshold in MB (default: 1024)")
    parser.add_argument("--memory-critical-threshold", type=float, default=2048,
                       help="Memory critical threshold in MB (default: 2048)")
    
    # Configuration file support
    parser.add_argument("--config", help="Configuration file path (JSON format)")
    parser.add_argument("--save-config", help="Save current configuration to file")
    
    # Subcommands
    subparsers = parser.add_subparsers(dest="command", help="Available commands")
    
    # Process command
    process_parser = subparsers.add_parser("process", help="Process a single document")
    process_parser.add_argument("input", help="Input DOCX file path")
    process_parser.add_argument("intent", help="User intent description")
    
    # Batch command
    batch_parser = subparsers.add_parser("batch", help="Batch process documents")
    batch_parser.add_argument("batch_dir", help="Directory containing DOCX files")
    batch_parser.add_argument("intent", help="User intent description for all documents")
    
    # Dry run command
    dry_run_parser = subparsers.add_parser("dry-run", help="Generate plan without execution")
    dry_run_parser.add_argument("input", help="Input DOCX file path")
    dry_run_parser.add_argument("intent", help="User intent description")
    
    # Config command
    config_parser = subparsers.add_parser("config", help="Configuration management")
    config_subparsers = config_parser.add_subparsers(dest="config_action", help="Configuration actions")
    
    # Config show
    config_show_parser = config_subparsers.add_parser("show", help="Show current configuration")
    
    # Config create
    config_create_parser = config_subparsers.add_parser("create", help="Create configuration template")
    config_create_parser.add_argument("output", help="Output configuration file path")
    
    # Status command
    status_parser = subparsers.add_parser("status", help="Check system status and requirements")
    
    # Parse arguments
    args = parser.parse_args()
    
    if not args.command:
        parser.print_help()
        return 1
    
    # Load configuration file if specified
    if hasattr(args, 'config') and args.config:
        config = load_config(args.config)
        apply_config(args, config)
    
    # Save configuration if requested
    if hasattr(args, 'save_config') and args.save_config:
        save_config(args, args.save_config)
    
    # Setup logging
    setup_logging(args.verbose, args.log_file)
    
    # Execute command
    if args.command == "process":
        return process_single_document(args)
    elif args.command == "batch":
        return process_batch_documents(args)
    elif args.command == "dry-run":
        return dry_run_document(args)
    elif args.command == "config":
        return handle_config_command(args)
    elif args.command == "status":
        return check_system_status(args)
    else:
        print(f"Unknown command: {args.command}")
        return 1


if __name__ == "__main__":
    sys.exit(main())