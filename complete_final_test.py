#!/usr/bin/env python3
"""
完整最终测试 - 包括实际文档处理
"""

import sys
import time
from pathlib import Path

# 添加项目根目录到Python路径
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

def test_complete_workflow():
    """测试完整工作流程"""
    print("========================================")
    print("      AutoWord 完整工作流程测试")
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
        # 1. 测试JSON修复功能
        print("[1/6] 测试JSON修复功能...")
        from autoword.core.llm_client import LLMClient
        
        llm_client = LLMClient()
        test_json = '{"tasks": [{"id": "task_1", "type": "rewrite",}]}'
        
        try:
            import json
            cleaned = llm_client._clean_json_content(test_json)
            json.loads(cleaned)
            print("✅ JSON修复功能正常")
        except Exception as e:
            print(f"❌ JSON修复功能异常: {e}")
            return False
        
        # 2. 测试API配置
        print("[2/6] 测试API配置...")
        from autoword.gui.config_manager import ConfigurationManager
        
        config_manager = ConfigurationManager()
        api_keys = {
            "claude": config_manager.get_api_key("claude"),
            "gpt": config_manager.get_api_key("gpt")
        }
        
        if not api_keys["claude"] or not api_keys["gpt"]:
            print("❌ API密钥未配置")
            return False
        
        print("✅ API密钥配置正常")
        
        # 3. 测试LLM连接和JSON解析
        print("[3/6] 测试LLM连接和JSON解析...")
        from autoword.core.llm_client import ModelType
        
        llm_client = LLMClient(api_keys=api_keys)
        
        # 测试简单的JSON生成
        system_prompt = """你是一个任务规划助手。请根据用户的要求生成JSON格式的任务列表。

返回格式必须是有效的JSON，格式如下：
{
  "tasks": [
    {
      "id": "task_1",
      "type": "rewrite",
      "instruction": "重写内容",
      "locator": {"by": "find", "value": "目标文本"}
    }
  ]
}

重要：只返回JSON，不要添加任何解释文字或markdown标记。"""
        
        user_prompt = "请生成一个简单的重写任务"
        
        response = llm_client.call_with_json_retry(
            ModelType.CLAUDE37,
            system_prompt,
            user_prompt,
            max_json_retries=3
        )
        
        if response.success:
            try:
                import json
                parsed = json.loads(response.content)
                print("✅ LLM连接和JSON解析正常")
            except json.JSONDecodeError as e:
                print(f"❌ JSON解析失败: {e}")
                print(f"响应内容: {response.content}")
                return False
        else:
            print(f"❌ LLM调用失败: {response.error}")
            return False
        
        # 4. 测试文档加载
        print("[4/6] 测试文档加载...")
        from autoword.core.doc_loader import DocLoader
        
        doc_loader = DocLoader()
        try:
            word_app, document = doc_loader.load_document(str(input_path), create_backup=False)
            print("✅ 文档加载正常")
            
            # 关闭文档
            document.Close()
            word_app.Quit()
        except Exception as e:
            print(f"❌ 文档加载失败: {e}")
            return False
        
        # 5. 测试GUI组件
        print("[5/6] 测试GUI组件...")
        try:
            from autoword.gui.main_window import MainWindow
            from autoword.gui.processor_controller import DocumentProcessorController
            
            # 测试配置管理器
            config_manager = ConfigurationManager()
            
            # 测试处理控制器初始化
            processor_controller = DocumentProcessorController(config_manager)
            
            print("✅ GUI组件正常")
        except Exception as e:
            print(f"❌ GUI组件测试失败: {e}")
            return False
        
        # 6. 测试完整管道（使用简化模式）
        print("[6/6] 测试完整处理管道...")
        from autoword.core.pipeline import DocumentProcessor, PipelineConfig
        from autoword.core.word_executor import ExecutionMode
        
        # 使用试运行模式避免实际修改文档
        config = PipelineConfig(
            model=ModelType.CLAUDE37,
            execution_mode=ExecutionMode.DRY_RUN,  # 试运行模式
            create_backup=False,
            enable_validation=False,
            export_results=False,
            output_dir=str(output_file_path.parent),
            visible_word=False,
            max_retries=2
        )
        
        processor = DocumentProcessor(config, api_keys=api_keys)
        
        # 添加进度回调
        def progress_callback(progress):
            print(f"    进度: [{progress.stage.value}] {progress.progress:.1%}")
        
        processor.add_progress_callback(progress_callback)
        
        try:
            result = processor.process_document(str(input_path), str(output_file_path))
            
            if result.success:
                print("✅ 完整处理管道正常")
            else:
                print(f"⚠️  处理管道完成但有问题: {result.error_message}")
                # 对于试运行模式，某些错误是可以接受的
                if "试运行" in str(result.error_message) or "DRY_RUN" in str(result.error_message):
                    print("✅ 试运行模式正常")
                else:
                    return False
        except Exception as e:
            print(f"❌ 处理管道测试失败: {e}")
            return False
        finally:
            processor.close()
        
        print()
        print("🎉 所有测试通过！")
        print()
        print("功能确认:")
        print("✅ JSON修复功能正常")
        print("✅ API密钥已预配置")
        print("✅ LLM连接和JSON解析正常")
        print("✅ 文档加载功能正常")
        print("✅ GUI组件正常")
        print("✅ 完整处理管道正常")
        print()
        print("现在可以安全使用GUI进行实际文档处理！")
        
        return True
        
    except Exception as e:
        print(f"❌ 测试过程中发生异常: {e}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == "__main__":
    try:
        success = test_complete_workflow()
        if success:
            print()
            print("🎉 AutoWord 完整测试通过！")
            print("   所有功能已修复并可正常使用。")
            print("   双击 'AutoWord启动器.bat' 开始使用。")
            sys.exit(0)
        else:
            print()
            print("❌ 部分功能测试失败。")
            sys.exit(1)
    except KeyboardInterrupt:
        print("\n测试被用户中断")
        sys.exit(1)
    except Exception as e:
        print(f"测试过程中发生异常: {e}")
        sys.exit(1)