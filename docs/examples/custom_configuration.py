#!/usr/bin/env python3
"""
AutoWord vNext - Custom Configuration Examples

This script demonstrates various configuration options for AutoWord vNext,
including LLM settings, localization, validation rules, and audit options.
"""

import os
import sys
from pathlib import Path
from typing import Dict, Any

# Add autoword to path for examples
sys.path.insert(0, str(Path(__file__).parent.parent.parent))

from autoword.vnext import VNextPipeline
from autoword.vnext.core import (
    VNextConfig, LLMConfig, LocalizationConfig, 
    ValidationConfig, AuditConfig, ExecutorConfig
)


def openai_configuration_example():
    """Configure pipeline for OpenAI GPT models"""
    print("=== OpenAI Configuration ===")
    
    # OpenAI configuration
    config = VNextConfig(
        llm=LLMConfig(
            provider="openai",
            model="gpt-4",
            api_key=os.getenv("OPENAI_API_KEY", "sk-your-key-here"),
            base_url="https://api.openai.com/v1",
            temperature=0.1,
            max_tokens=4000,
            timeout=30,
            retry_attempts=3
        )
    )
    
    pipeline = VNextPipeline(config)
    
    print("OpenAI Configuration:")
    print(f"  Model: {config.llm.model}")
    print(f"  Temperature: {config.llm.temperature}")
    print(f"  Max tokens: {config.llm.max_tokens}")
    print(f"  Timeout: {config.llm.timeout}s")
    
    # Test with sample document
    try:
        result = pipeline.generate_plan(
            "sample_document.docx",
            "Delete abstract section and update TOC"
        )
        
        if result.is_valid:
            print(f"✅ Plan generated successfully with {len(result.plan.ops)} operations")
        else:
            print(f"❌ Plan generation failed: {result.errors}")
            
    except Exception as e:
        print(f"❌ Configuration test failed: {e}")


def anthropic_configuration_example():
    """Configure pipeline for Anthropic Claude models"""
    print("\n=== Anthropic Claude Configuration ===")
    
    # Anthropic configuration
    config = VNextConfig(
        llm=LLMConfig(
            provider="anthropic",
            model="claude-3-sonnet-20240229",
            api_key=os.getenv("ANTHROPIC_API_KEY", "sk-ant-your-key-here"),
            temperature=0.1,
            max_tokens=4000,
            timeout=45,
            retry_attempts=2
        )
    )
    
    pipeline = VNextPipeline(config)
    
    print("Anthropic Configuration:")
    print(f"  Model: {config.llm.model}")
    print(f"  Provider: {config.llm.provider}")
    print(f"  Temperature: {config.llm.temperature}")
    
    # Note: This would require actual API key to test
    print("✅ Configuration created (API key required for testing)")


def chinese_localization_example():
    """Configure pipeline for Chinese document processing"""
    print("\n=== Chinese Localization Configuration ===")
    
    config = VNextConfig(
        llm=LLMConfig(
            provider="openai",
            model="gpt-4",
            api_key=os.getenv("OPENAI_API_KEY", "sk-your-key-here")
        ),
        localization=LocalizationConfig(
            language="zh-CN",
            style_aliases={
                "Heading 1": "标题 1",
                "Heading 2": "标题 2", 
                "Heading 3": "标题 3",
                "Normal": "正文",
                "Title": "标题",
                "Subtitle": "副标题"
            },
            font_fallbacks={
                "楷体": ["楷体", "楷体_GB2312", "STKaiti", "KaiTi"],
                "宋体": ["宋体", "SimSun", "NSimSun"],
                "黑体": ["黑体", "SimHei", "Microsoft YaHei"],
                "仿宋": ["仿宋", "FangSong", "FangSong_GB2312"],
                "微软雅黑": ["微软雅黑", "Microsoft YaHei", "SimHei"]
            },
            enable_font_fallback=True,
            log_fallbacks=True
        )
    )
    
    pipeline = VNextPipeline(config)
    
    print("Chinese Localization:")
    print("  Style aliases:")
    for en, zh in config.localization.style_aliases.items():
        print(f"    {en} → {zh}")
    
    print("  Font fallbacks:")
    for font, fallbacks in config.localization.font_fallbacks.items():
        print(f"    {font}: {' → '.join(fallbacks)}")
    
    # Test Chinese document processing
    user_intent = "删除摘要部分，更新目录，设置标题1为楷体字体"
    print(f"\nTesting with Chinese intent: {user_intent}")
    print("✅ Configuration ready for Chinese documents")


def strict_validation_example():
    """Configure pipeline with strict validation rules"""
    print("\n=== Strict Validation Configuration ===")
    
    config = VNextConfig(
        llm=LLMConfig(
            provider="openai",
            model="gpt-4",
            api_key=os.getenv("OPENAI_API_KEY", "sk-your-key-here")
        ),
        validation=ValidationConfig(
            strict_mode=True,
            rollback_on_failure=True,
            chapter_assertions=True,
            style_assertions=True,
            toc_assertions=True,
            pagination_assertions=True,
            forbidden_level1_headings=[
                "摘要", "Abstract", 
                "参考文献", "References", "Bibliography",
                "附录", "Appendix", "Appendices"
            ]
        )
    )
    
    pipeline = VNextPipeline(config)
    
    print("Strict Validation Rules:")
    print(f"  Strict mode: {config.validation.strict_mode}")
    print(f"  Rollback on failure: {config.validation.rollback_on_failure}")
    print(f"  Chapter assertions: {config.validation.chapter_assertions}")
    print(f"  Style assertions: {config.validation.style_assertions}")
    print(f"  TOC assertions: {config.validation.toc_assertions}")
    
    print("  Forbidden level 1 headings:")
    for heading in config.validation.forbidden_level1_headings:
        print(f"    - {heading}")
    
    print("✅ Strict validation configured")


def comprehensive_audit_example():
    """Configure pipeline with comprehensive audit settings"""
    print("\n=== Comprehensive Audit Configuration ===")
    
    config = VNextConfig(
        llm=LLMConfig(
            provider="openai",
            model="gpt-4",
            api_key=os.getenv("OPENAI_API_KEY", "sk-your-key-here")
        ),
        audit=AuditConfig(
            save_snapshots=True,
            generate_diff_reports=True,
            output_directory="./comprehensive_audit",
            keep_audit_days=90,
            compress_old_audits=True
        )
    )
    
    pipeline = VNextPipeline(config)
    
    print("Comprehensive Audit Settings:")
    print(f"  Save snapshots: {config.audit.save_snapshots}")
    print(f"  Generate diff reports: {config.audit.generate_diff_reports}")
    print(f"  Output directory: {config.audit.output_directory}")
    print(f"  Keep audit days: {config.audit.keep_audit_days}")
    print(f"  Compress old audits: {config.audit.compress_old_audits}")
    
    # Create audit directory if it doesn't exist
    audit_dir = Path(config.audit.output_directory)
    audit_dir.mkdir(exist_ok=True)
    print(f"✅ Audit directory created: {audit_dir.absolute()}")


def performance_optimized_example():
    """Configure pipeline for performance optimization"""
    print("\n=== Performance Optimized Configuration ===")
    
    config = VNextConfig(
        llm=LLMConfig(
            provider="openai",
            model="gpt-4",
            api_key=os.getenv("OPENAI_API_KEY", "sk-your-key-here"),
            temperature=0.05,  # Lower temperature for consistency
            max_tokens=2000,   # Reduced for faster responses
            timeout=20         # Shorter timeout
        ),
        executor=ExecutorConfig(
            batch_operations=True,      # Batch similar operations
            com_timeout=20,             # Shorter COM timeout
            retry_com_errors=True,      # Retry COM errors
            max_com_retries=2,          # Fewer retries
            enable_localization=True
        ),
        audit=AuditConfig(
            save_snapshots=True,
            generate_diff_reports=False,  # Skip diff reports for speed
            compress_old_audits=True
        )
    )
    
    pipeline = VNextPipeline(config)
    
    print("Performance Optimizations:")
    print(f"  LLM temperature: {config.llm.temperature} (lower = more consistent)")
    print(f"  LLM max tokens: {config.llm.max_tokens} (reduced for speed)")
    print(f"  Batch operations: {config.executor.batch_operations}")
    print(f"  COM timeout: {config.executor.com_timeout}s")
    print(f"  Skip diff reports: {not config.audit.generate_diff_reports}")
    
    print("✅ Performance optimized configuration ready")


def development_debug_example():
    """Configure pipeline for development and debugging"""
    print("\n=== Development/Debug Configuration ===")
    
    config = VNextConfig(
        llm=LLMConfig(
            provider="openai",
            model="gpt-4",
            api_key=os.getenv("OPENAI_API_KEY", "sk-your-key-here"),
            temperature=0.2,    # Slightly higher for variety
            timeout=60          # Longer timeout for debugging
        ),
        validation=ValidationConfig(
            strict_mode=False,          # Relaxed for testing
            rollback_on_failure=False,  # Keep failed attempts
            chapter_assertions=True,
            style_assertions=False,     # Skip style validation
            toc_assertions=True
        ),
        audit=AuditConfig(
            save_snapshots=True,
            generate_diff_reports=True,
            output_directory="./debug_audit",
            keep_audit_days=7,          # Short retention for testing
            compress_old_audits=False   # Keep uncompressed for inspection
        ),
        executor=ExecutorConfig(
            retry_com_errors=True,
            max_com_retries=5,          # More retries for debugging
            enable_localization=True
        )
    )
    
    pipeline = VNextPipeline(config)
    
    print("Development/Debug Settings:")
    print(f"  Strict validation: {config.validation.strict_mode}")
    print(f"  Rollback on failure: {config.validation.rollback_on_failure}")
    print(f"  Max COM retries: {config.executor.max_com_retries}")
    print(f"  Audit retention: {config.audit.keep_audit_days} days")
    print(f"  Compress audits: {config.audit.compress_old_audits}")
    
    print("✅ Debug configuration ready")


def load_config_from_file_example():
    """Load configuration from JSON file"""
    print("\n=== Load Configuration from File ===")
    
    # Sample configuration file content
    config_content = {
        "llm": {
            "provider": "openai",
            "model": "gpt-4",
            "api_key": "sk-your-key-here",
            "temperature": 0.1,
            "max_tokens": 4000
        },
        "localization": {
            "language": "zh-CN",
            "style_aliases": {
                "Heading 1": "标题 1",
                "Normal": "正文"
            },
            "font_fallbacks": {
                "楷体": ["楷体", "楷体_GB2312", "STKaiti"]
            }
        },
        "validation": {
            "strict_mode": True,
            "rollback_on_failure": True
        },
        "audit": {
            "save_snapshots": True,
            "output_directory": "./file_config_audit"
        }
    }
    
    # Save sample config file
    import json
    config_file = "vnext_config.json"
    
    with open(config_file, 'w', encoding='utf-8') as f:
        json.dump(config_content, f, indent=2, ensure_ascii=False)
    
    print(f"Sample configuration saved to: {config_file}")
    
    # Load configuration (this would be implemented in the actual system)
    print("Configuration file structure:")
    print(json.dumps(config_content, indent=2, ensure_ascii=False))
    
    print("✅ Configuration file example created")


def environment_specific_configs():
    """Show different configurations for different environments"""
    print("\n=== Environment-Specific Configurations ===")
    
    environments = {
        "development": {
            "description": "Development environment with debug features",
            "config": VNextConfig(
                validation=ValidationConfig(strict_mode=False, rollback_on_failure=False),
                audit=AuditConfig(keep_audit_days=7, compress_old_audits=False)
            )
        },
        "testing": {
            "description": "Testing environment with comprehensive validation",
            "config": VNextConfig(
                validation=ValidationConfig(strict_mode=True, rollback_on_failure=True),
                audit=AuditConfig(save_snapshots=True, generate_diff_reports=True)
            )
        },
        "production": {
            "description": "Production environment optimized for performance",
            "config": VNextConfig(
                llm=LLMConfig(temperature=0.05, max_tokens=2000),
                executor=ExecutorConfig(batch_operations=True, com_timeout=20),
                audit=AuditConfig(keep_audit_days=30, compress_old_audits=True)
            )
        }
    }
    
    for env_name, env_info in environments.items():
        print(f"\n{env_name.upper()} Environment:")
        print(f"  Description: {env_info['description']}")
        
        config = env_info['config']
        if hasattr(config, 'validation') and config.validation:
            print(f"  Strict mode: {config.validation.strict_mode}")
        if hasattr(config, 'audit') and config.audit:
            print(f"  Audit retention: {config.audit.keep_audit_days} days")
        if hasattr(config, 'llm') and config.llm:
            print(f"  LLM temperature: {config.llm.temperature}")


def main():
    """Run all configuration examples"""
    print("AutoWord vNext - Custom Configuration Examples")
    print("=" * 70)
    
    try:
        # Run configuration examples
        openai_configuration_example()
        anthropic_configuration_example()
        chinese_localization_example()
        strict_validation_example()
        comprehensive_audit_example()
        performance_optimized_example()
        development_debug_example()
        load_config_from_file_example()
        environment_specific_configs()
        
        print("\n" + "=" * 70)
        print("Configuration examples completed.")
        print("Customize these examples for your specific needs.")
        
    except ImportError as e:
        print(f"❌ Import error: {e}")
        print("Please ensure AutoWord vNext is properly installed.")
        
    except Exception as e:
        print(f"❌ Unexpected error: {e}")


if __name__ == "__main__":
    main()