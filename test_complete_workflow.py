#!/usr/bin/env python3
"""
完整工作流程测试
"""

import sys
import os
from pathlib import Path

# 添加项目根目录到Python路径
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

def test_document_processing():
    """测试文档处理完整流程"""
    print("========================================")
    print("      AutoWord 完整流程测试")
    print("========================================")
    print()
    
    # 测试文档路径
    test_doc = r"C:\Users\y\Desktop\郭宇睿-现代汉语言文学在社会发展中的作用与影响研究2稿批注.docx"
    
    # 检查文档是否存在
    if not Path(test_doc).exists():
        print(f"❌ 测试文档不存在: {test_doc}")
        return False
    
    print(f"✅ 找到测试文档: {Path(test_doc).name}")
    print()
    
    try:
        # 导入必要的模块
        from autoword.core.pipeline import DocumentProcessor, PipelineConfig
        from autoword.core.llm_client import ModelType
        from autoword.core.word_executor import ExecutionMode
        
        print("[1/6] 模块导入成功")
        
        # 创建输出目录
        output_dir = Path(test_doc).parent / "+process"
        output_dir.mkdir(exist_ok=True)
        print(f"✅ 输出目录: {output_dir}")
        
        # 配置API密钥
        api_keys = {
            "claude": "sk-3w1JFbWUq7tKjpLlopdkISQ9F6fpLhHx5viD0frh43ESE9Io",
            "gpt": "sk-NhjnJtqlZMx4PGTqvkGlH4POT82HHBrBnBbWOat99Bs5VZXi"
        }
        print("[2/6] API密钥配置完成")
        
        # 创建管道配置 - 使用Claude 3.7
        config = PipelineConfig(
            model=ModelType.CLAUDE37,
            execution_mode=ExecutionMode.DRY_RUN,  # 先用试运行模式
            create_backup=True,
            enable_validation=True,
            export_results=True,
            output_dir=str(output_dir),
            visible_word=False,
            max_retries=3
        )
        print("[3/6] 管道配置创建完成")
        
        # 创建处理器
        processor = DocumentProcessor(config, api_keys=api_keys)
        print("[4/6] 文档处理器创建完成")
        
        # 添加进度回调
        def progress_callback(progress):
            print(f"    进度: [{progress.stage.value}] {progress.progress:.1%} - {progress.message}")
        
        processor.add_progress_callback(progress_callback)
        print("[5/6] 进度回调设置完成")
        
        # 执行处理
        print()
        print("开始处理文档...")
        print("=" * 50)
        
        result = processor.process_document(test_doc)
        
        print("=" * 50)
        print()
        
        if result.success:
            print("✅ 文档处理成功!")
            print(f"   耗时: {result.total_time:.2f}秒")
            print(f"   完成阶段: {len(result.stages_completed)} 个")
            
            if result.task_plan:
                print(f"   生成任务: {len(result.task_plan.tasks)} 个")
            
            if result.exported_files:
                print(f"   导出文件: {len(result.exported_files)} 个")
                for file_type, file_path in result.exported_files.items():
                    print(f"     - {file_type}: {file_path}")
            
            print()
            print("🎉 完整流程测试成功!")
            return True
            
        else:
            print("❌ 文档处理失败:")
            print(f"   错误: {result.error_message}")
            return False
            
    except Exception as e:
        print(f"❌ 测试过程中发生异常: {e}")
        import traceback
        traceback.print_exc()
        return False
    
    finally:
        try:
            processor.close()
        except:
            pass


if __name__ == "__main__":
    try:
        success = test_document_processing()
        if success:
            print()
            print("✅ 所有测试通过！GUI应用程序已准备就绪。")
            print("   请运行 'AutoWord启动器.bat' 启动GUI界面。")
            sys.exit(0)
        else:
            print()
            print("❌ 测试失败，请检查配置。")
            sys.exit(1)
    except KeyboardInterrupt:
        print("\n测试被用户中断")
        sys.exit(1)
    except Exception as e:
        print(f"测试过程中发生异常: {e}")
        sys.exit(1)