#!/usr/bin/env python3
"""
测试输出文件功能
"""

import sys
from pathlib import Path

# 添加项目根目录到Python路径
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

def test_output_file_functionality():
    """测试输出文件功能"""
    print("========================================")
    print("      测试输出文件功能")
    print("========================================")
    print()
    
    # 测试文档路径
    test_doc = r"C:\Users\y\Desktop\郭宇睿-现代汉语言文学在社会发展中的作用与影响研究2稿批注.docx"
    
    # 检查文档是否存在
    if not Path(test_doc).exists():
        print(f"❌ 测试文档不存在: {test_doc}")
        return False
    
    print(f"✅ 找到测试文档: {Path(test_doc).name}")
    
    # 生成输出文件路径
    input_path = Path(test_doc)
    output_file_path = input_path.parent / f"{input_path.stem}.process{input_path.suffix}"
    
    print(f"📁 输入文件: {input_path}")
    print(f"📄 输出文件: {output_file_path}")
    print()
    
    try:
        # 导入必要的模块
        from autoword.core.pipeline import DocumentProcessor, PipelineConfig
        from autoword.core.llm_client import ModelType
        from autoword.core.word_executor import ExecutionMode
        
        print("[1/4] 模块导入成功")
        
        # 配置API密钥
        api_keys = {
            "claude": "sk-3w1JFbWUq7tKjpLlopdkISQ9F6fpLhHx5viD0frh43ESE9Io",
            "gpt": "sk-NhjnJtqlZMx4PGTqvkGlH4POT82HHBrBnBbWOat99Bs5VZXi"
        }
        print("[2/4] API密钥配置完成")
        
        # 创建管道配置
        config = PipelineConfig(
            model=ModelType.CLAUDE37,
            execution_mode=ExecutionMode.NORMAL,
            create_backup=True,
            enable_validation=True,
            export_results=True,
            output_dir=str(output_file_path.parent),
            visible_word=False,
            max_retries=3
        )
        print("[3/4] 管道配置创建完成")
        
        # 创建处理器
        processor = DocumentProcessor(config, api_keys=api_keys)
        print("[4/4] 文档处理器创建完成")
        
        # 添加进度回调
        def progress_callback(progress):
            print(f"    进度: [{progress.stage.value}] {progress.progress:.1%} - {progress.message}")
        
        processor.add_progress_callback(progress_callback)
        
        # 执行处理
        print()
        print("开始处理文档...")
        print("=" * 50)
        
        result = processor.process_document(str(input_path), str(output_file_path))
        
        print("=" * 50)
        print()
        
        if result.success:
            print("✅ 文档处理成功!")
            print(f"   耗时: {result.total_time:.2f}秒")
            print(f"   完成阶段: {len(result.stages_completed)} 个")
            
            # 检查输出文件是否存在
            if output_file_path.exists():
                print(f"✅ 输出文件已创建: {output_file_path}")
                print(f"   文件大小: {output_file_path.stat().st_size} 字节")
            else:
                print(f"❌ 输出文件未创建: {output_file_path}")
                return False
            
            if result.task_plan:
                print(f"   生成任务: {len(result.task_plan.tasks)} 个")
            
            if result.exported_files:
                print(f"   导出文件: {len(result.exported_files)} 个")
                for file_type, file_path in result.exported_files.items():
                    print(f"     - {file_type}: {file_path}")
            
            print()
            print("🎉 输出文件功能测试成功!")
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
        success = test_output_file_functionality()
        if success:
            print()
            print("✅ 输出文件功能测试通过！")
            print("   现在GUI应用程序可以正确生成输出文件了。")
            sys.exit(0)
        else:
            print()
            print("❌ 输出文件功能测试失败。")
            sys.exit(1)
    except KeyboardInterrupt:
        print("\n测试被用户中断")
        sys.exit(1)
    except Exception as e:
        print(f"测试过程中发生异常: {e}")
        sys.exit(1)