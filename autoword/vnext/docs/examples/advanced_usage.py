#!/usr/bin/env python3
"""
AutoWord vNext Advanced Usage Examples

This module demonstrates advanced usage patterns including custom configurations,
module-level APIs, performance optimization, and integration scenarios.
"""

import json
import time
from pathlib import Path
from typing import Dict, List, Optional

from autoword.vnext.pipeline import VNextPipeline
from autoword.vnext.extractor import DocumentExtractor
from autoword.vnext.planner import DocumentPlanner
from autoword.vnext.executor import DocumentExecutor
from autoword.vnext.validator import DocumentValidator
from autoword.vnext.auditor import DocumentAuditor
from autoword.vnext.config import ConfigManager
from autoword.vnext.exceptions import VNextException


def module_level_api_usage():
    """
    Example 1: Using individual modules directly
    
    Shows how to use each pipeline module independently for
    custom workflows and advanced control.
    """
    print("=== Module-Level API Usage ===")
    
    try:
        # Initialize modules with custom configuration
        config = {
            "model": "gpt4",
            "temperature": 0.1,
            "visible": False
        }
        
        extractor = DocumentExtractor(config)
        planner = DocumentPlanner(config)
        executor = DocumentExecutor(config)
        validator = DocumentValidator(config)
        auditor = DocumentAuditor(config)
        
        document_path = "sample_document.docx"
        
        # Step 1: Extract document structure
        print("1. Extracting document structure...")
        structure = extractor.extract_structure(document_path)
        inventory = extractor.extract_inventory(document_path)
        
        print(f"   ✓ Found {len(structure.headings)} headings")
        print(f"   ✓ Found {len(structure.styles)} styles")
        print(f"   ✓ Inventory has {len(inventory.ooxml_fragments)} fragments")
        
        # Step 2: Generate execution plan
        print("2. Generating execution plan...")
        user_intent = "Remove abstract section and apply standard formatting"
        plan = planner.generate_plan(structure, user_intent)
        
        print(f"   ✓ Generated {len(plan.ops)} operations")
        for i, op in enumerate(plan.ops, 1):
            print(f"   {i}. {op.operation}")
        
        # Step 3: Execute plan (in this example, we'll skip actual execution)
        print("3. Plan execution (skipped in example)")
        print("   Would execute operations on document...")
        
        # Step 4: Validation (simulated)
        print("4. Validation (simulated)")
        print("   Would validate changes against assertions...")
        
        # Step 5: Create audit trail
        print("5. Creating audit trail...")
        audit_dir = auditor.create_audit_directory("./advanced_audits")
        print(f"   ✓ Created audit directory: {audit_dir}")
        
    except Exception as e:
        print(f"✗ Module-level API error: {e}")


def custom_configuration_management():
    """
    Example 2: Advanced configuration management
    
    Shows how to create, validate, and use custom configurations
    for different scenarios.
    """
    print("\n=== Custom Configuration Management ===")
    
    try:
        # Create different configurations for different scenarios
        
        # Development configuration
        dev_config = {
            "model": "gpt35",  # Faster for development
            "temperature": 0.2,
            "verbose": True,
            "visible": True,  # Show Word windows for debugging
            "monitoring_level": "debug",
            "audit_dir": "./dev_audits"
        }
        
        # Production configuration
        prod_config = {
            "model": "gpt4",  # More reliable for production
            "temperature": 0.05,
            "verbose": False,
            "visible": False,
            "monitoring_level": "detailed",
            "audit_dir": "./prod_audits",
            "max_execution_time_seconds": 600
        }
        
        # Batch processing configuration
        batch_config = {
            "model": "gpt35",  # Faster for batch operations
            "temperature": 0.1,
            "monitoring_level": "performance",
            "enable_memory_monitoring": True,
            "memory_warning_threshold": 1024,
            "memory_critical_threshold": 2048
        }
        
        # Save configurations
        config_dir = Path("./configs")
        config_dir.mkdir(exist_ok=True)
        
        for name, config in [("dev", dev_config), ("prod", prod_config), ("batch", batch_config)]:
            config_path = config_dir / f"{name}_config.json"
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=2, ensure_ascii=False)
            print(f"✓ Saved {name} configuration: {config_path}")
        
        # Validate configurations
        for name in ["dev", "prod", "batch"]:
            config_path = config_dir / f"{name}_config.json"
            try:
                config = ConfigManager.load_config(str(config_path))
                validation_result = ConfigManager.validate_config(config)
                
                if validation_result.is_valid:
                    print(f"✓ {name} configuration is valid")
                else:
                    print(f"✗ {name} configuration errors: {validation_result.errors}")
                    
            except Exception as e:
                print(f"✗ Error validating {name} config: {e}")
        
        # Use configuration in pipeline
        pipeline = VNextPipeline(dev_config)
        print("✓ Pipeline initialized with development configuration")
        
    except Exception as e:
        print(f"✗ Configuration management error: {e}")


def performance_optimization_example():
    """
    Example 3: Performance optimization techniques
    
    Shows how to optimize performance for large documents
    and batch processing scenarios.
    """
    print("\n=== Performance Optimization Example ===")
    
    try:
        # Performance-optimized configuration
        perf_config = {
            "model": "gpt35",  # Faster model
            "temperature": 0.1,
            "monitoring_level": "performance",
            "enable_memory_monitoring": True,
            "memory_warning_threshold": 2048,
            "memory_critical_threshold": 4096,
            "max_execution_time_seconds": 300
        }
        
        pipeline = VNextPipeline(perf_config)
        
        # Simulate processing multiple documents with timing
        documents = ["doc1.docx", "doc2.docx", "doc3.docx"]  # Example documents
        intent = "Apply standard formatting"
        
        total_start_time = time.time()
        results = []
        
        for i, doc_path in enumerate(documents, 1):
            print(f"Processing document {i}/{len(documents)}: {doc_path}")
            
            start_time = time.time()
            
            try:
                # Use dry run for this example to avoid needing actual files
                result = pipeline.dry_run(doc_path, intent)
                
                end_time = time.time()
                processing_time = end_time - start_time
                
                if result.plan_generated:
                    print(f"   ✓ Plan generated in {processing_time:.2f}s")
                    print(f"   Operations: {len(result.plan.ops)}")
                else:
                    print(f"   ✗ Plan generation failed: {result.error}")
                
                results.append({
                    "document": doc_path,
                    "success": result.plan_generated,
                    "time": processing_time,
                    "operations": len(result.plan.ops) if result.plan_generated else 0
                })
                
            except Exception as e:
                print(f"   ✗ Error processing {doc_path}: {e}")
                results.append({
                    "document": doc_path,
                    "success": False,
                    "time": time.time() - start_time,
                    "error": str(e)
                })
        
        total_time = time.time() - total_start_time
        
        # Performance summary
        print(f"\nPerformance Summary:")
        print(f"Total time: {total_time:.2f}s")
        print(f"Average time per document: {total_time/len(documents):.2f}s")
        
        successful = sum(1 for r in results if r["success"])
        print(f"Success rate: {successful}/{len(documents)} ({successful/len(documents)*100:.1f}%)")
        
        if successful > 0:
            avg_ops = sum(r["operations"] for r in results if r["success"]) / successful
            print(f"Average operations per document: {avg_ops:.1f}")
        
    except Exception as e:
        print(f"✗ Performance optimization error: {e}")


def custom_validation_rules():
    """
    Example 4: Custom validation rules and assertions
    
    Shows how to implement custom validation logic for
    specific document requirements.
    """
    print("\n=== Custom Validation Rules ===")
    
    try:
        # Custom validation configuration
        validation_config = {
            "validation_rules": {
                "chapter_assertions": {
                    "forbidden_level_1_headings": ["摘要", "参考文献", "Abstract", "References"],
                    "required_level_1_headings": ["引言", "结论"]
                },
                "style_assertions": {
                    "Heading 1": {
                        "font_east_asian": "楷体",
                        "font_size_pt": 12,
                        "font_bold": True,
                        "line_spacing_mode": "MULTIPLE",
                        "line_spacing_value": 2.0
                    },
                    "Heading 2": {
                        "font_east_asian": "宋体",
                        "font_size_pt": 12,
                        "font_bold": True,
                        "line_spacing_mode": "MULTIPLE",
                        "line_spacing_value": 2.0
                    },
                    "Normal": {
                        "font_east_asian": "宋体",
                        "font_size_pt": 12,
                        "font_bold": False,
                        "line_spacing_mode": "MULTIPLE",
                        "line_spacing_value": 2.0
                    }
                },
                "toc_assertions": {
                    "require_toc": True,
                    "max_toc_levels": 3,
                    "toc_style_consistency": True
                }
            }
        }
        
        # Initialize validator with custom rules
        validator = DocumentValidator(validation_config)
        
        # Example: Validate a document structure (simulated)
        print("Custom validation rules configured:")
        
        rules = validation_config["validation_rules"]
        
        print("Chapter assertions:")
        forbidden = rules["chapter_assertions"]["forbidden_level_1_headings"]
        print(f"   Forbidden level 1 headings: {', '.join(forbidden)}")
        
        required = rules["chapter_assertions"]["required_level_1_headings"]
        print(f"   Required level 1 headings: {', '.join(required)}")
        
        print("Style assertions:")
        for style_name, style_rules in rules["style_assertions"].items():
            print(f"   {style_name}:")
            for rule, value in style_rules.items():
                print(f"     {rule}: {value}")
        
        print("TOC assertions:")
        toc_rules = rules["toc_assertions"]
        for rule, value in toc_rules.items():
            print(f"   {rule}: {value}")
        
        print("✓ Custom validation rules configured successfully")
        
    except Exception as e:
        print(f"✗ Custom validation error: {e}")


def localization_and_fonts():
    """
    Example 5: Advanced localization and font handling
    
    Shows how to configure localization settings for different
    languages and font environments.
    """
    print("\n=== Localization and Font Handling ===")
    
    try:
        # Advanced localization configuration
        localization_config = {
            "localization": {
                "style_aliases": {
                    # English to Chinese mappings
                    "Heading 1": "标题 1",
                    "Heading 2": "标题 2",
                    "Heading 3": "标题 3",
                    "Normal": "正文",
                    "Title": "标题",
                    "Subtitle": "副标题",
                    
                    # Custom style mappings
                    "Chapter Title": "章节标题",
                    "Section Title": "节标题"
                },
                "font_fallbacks": {
                    # Chinese fonts with comprehensive fallback chains
                    "楷体": ["楷体", "楷体_GB2312", "STKaiti", "KaiTi", "SimKai"],
                    "宋体": ["宋体", "SimSun", "NSimSun", "宋体-简"],
                    "黑体": ["黑体", "SimHei", "Microsoft YaHei", "微软雅黑"],
                    "仿宋": ["仿宋", "FangSong", "仿宋_GB2312", "STFangsong"],
                    
                    # English fonts
                    "Times New Roman": ["Times New Roman", "Times", "serif"],
                    "Arial": ["Arial", "Helvetica", "sans-serif"],
                    "Calibri": ["Calibri", "Arial", "sans-serif"],
                    
                    # Custom fonts with fallbacks
                    "CustomFont": ["CustomFont", "Arial", "sans-serif"]
                },
                "regional_settings": {
                    "default_language": "zh-CN",
                    "date_format": "YYYY-MM-DD",
                    "number_format": "chinese"
                }
            }
        }
        
        # Initialize pipeline with localization
        pipeline = VNextPipeline(localization_config)
        
        print("Localization configuration:")
        
        # Display style aliases
        aliases = localization_config["localization"]["style_aliases"]
        print(f"Style aliases ({len(aliases)} mappings):")
        for english, chinese in aliases.items():
            print(f"   {english} → {chinese}")
        
        # Display font fallbacks
        fallbacks = localization_config["localization"]["font_fallbacks"]
        print(f"\nFont fallback chains ({len(fallbacks)} fonts):")
        for primary, chain in fallbacks.items():
            print(f"   {primary}: {' → '.join(chain)}")
        
        # Example: Test font resolution (simulated)
        print(f"\nFont resolution examples:")
        test_fonts = ["楷体", "宋体", "NonexistentFont"]
        
        for font in test_fonts:
            if font in fallbacks:
                chain = fallbacks[font]
                print(f"   {font}: Will try {' → '.join(chain[:3])}...")
            else:
                print(f"   {font}: No fallback chain configured")
        
        print("✓ Localization configured successfully")
        
    except Exception as e:
        print(f"✗ Localization error: {e}")


def integration_with_external_systems():
    """
    Example 6: Integration with external systems
    
    Shows how to integrate AutoWord vNext with external systems,
    databases, and workflows.
    """
    print("\n=== Integration with External Systems ===")
    
    try:
        # Example: Document processing workflow integration
        class DocumentWorkflow:
            def __init__(self):
                self.pipeline = VNextPipeline({
                    "model": "gpt4",
                    "audit_dir": "./workflow_audits"
                })
                self.processed_documents = []
            
            def process_document_with_metadata(self, doc_path: str, metadata: Dict):
                """Process document with external metadata"""
                print(f"Processing {doc_path} with metadata...")
                
                # Generate intent based on metadata
                intent = self._generate_intent_from_metadata(metadata)
                print(f"Generated intent: {intent}")
                
                # Process document
                result = self.pipeline.process_document(doc_path, intent)
                
                # Store result with metadata
                processing_record = {
                    "document_path": doc_path,
                    "metadata": metadata,
                    "intent": intent,
                    "result": {
                        "status": result.status,
                        "output_path": result.output_path,
                        "audit_directory": result.audit_directory,
                        "execution_time": result.execution_time
                    },
                    "timestamp": time.time()
                }
                
                self.processed_documents.append(processing_record)
                
                return processing_record
            
            def _generate_intent_from_metadata(self, metadata: Dict) -> str:
                """Generate processing intent based on document metadata"""
                intents = []
                
                if metadata.get("remove_abstract", False):
                    intents.append("remove abstract section")
                
                if metadata.get("remove_references", False):
                    intents.append("remove references section")
                
                if metadata.get("update_toc", True):
                    intents.append("update table of contents")
                
                if metadata.get("apply_formatting", False):
                    style = metadata.get("formatting_style", "standard")
                    intents.append(f"apply {style} formatting")
                
                return ", ".join(intents) if intents else "standardize document formatting"
            
            def get_processing_summary(self) -> Dict:
                """Get summary of all processed documents"""
                total = len(self.processed_documents)
                successful = sum(1 for doc in self.processed_documents 
                               if doc["result"]["status"] == "SUCCESS")
                
                return {
                    "total_processed": total,
                    "successful": successful,
                    "failed": total - successful,
                    "success_rate": successful / total if total > 0 else 0,
                    "average_time": sum(doc["result"]["execution_time"] or 0 
                                      for doc in self.processed_documents) / total if total > 0 else 0
                }
        
        # Example usage
        workflow = DocumentWorkflow()
        
        # Simulate processing documents with different metadata
        test_documents = [
            {
                "path": "research_paper.docx",
                "metadata": {
                    "document_type": "research_paper",
                    "remove_abstract": True,
                    "remove_references": True,
                    "update_toc": True,
                    "apply_formatting": True,
                    "formatting_style": "academic"
                }
            },
            {
                "path": "report.docx",
                "metadata": {
                    "document_type": "business_report",
                    "remove_abstract": False,
                    "remove_references": False,
                    "update_toc": True,
                    "apply_formatting": True,
                    "formatting_style": "corporate"
                }
            }
        ]
        
        # Process documents (using dry run for example)
        for doc_info in test_documents:
            try:
                # Simulate processing
                intent = workflow._generate_intent_from_metadata(doc_info["metadata"])
                print(f"Would process {doc_info['path']} with intent: {intent}")
                
            except Exception as e:
                print(f"Error processing {doc_info['path']}: {e}")
        
        print("✓ Integration workflow example completed")
        
    except Exception as e:
        print(f"✗ Integration error: {e}")


def main():
    """
    Run all advanced usage examples
    """
    print("AutoWord vNext Advanced Usage Examples")
    print("=" * 50)
    
    # Note about sample documents
    print("Note: These examples demonstrate advanced patterns.")
    print("Some examples use simulated data to avoid requiring specific documents.")
    print()
    
    # Run examples
    module_level_api_usage()
    custom_configuration_management()
    performance_optimization_example()
    custom_validation_rules()
    localization_and_fonts()
    integration_with_external_systems()
    
    print("\n" + "=" * 50)
    print("Advanced usage examples completed!")
    print("\nNext steps:")
    print("- Adapt these patterns to your specific use cases")
    print("- Explore the API Reference for detailed documentation")
    print("- Consider performance implications for your workload")
    print("- Test thoroughly with your document types and requirements")


if __name__ == "__main__":
    main()