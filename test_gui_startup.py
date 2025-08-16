#!/usr/bin/env python3
"""
测试GUI启动
"""

import sys
import time
from pathlib import Path

# 添加项目根目录到Python路径
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

def test_gui_startup():
    """测试GUI启动"""
    print("========================================")
    print("      GUI启动测试")
    print("========================================")
    print()
    
    try:
        # 测试所有GUI组件导入
        print("[1/4] 测试GUI组件导入...")
        
        from autoword.gui.config_manager import ConfigurationManager
        from autoword.gui.main_window import MainWindow
        from autoword.gui.processor_controller import DocumentProcessorController
        from autoword.gui.error_handler import ErrorHandler
        
        print("✅ GUI组件导入成功")
        
        # 测试配置管理器
        print("[2/4] 测试配置管理器...")
        config_manager = ConfigurationManager()
        
        # 检查API密钥
        claude_key = config_manager.get_api_key("claude")
        gpt_key = config_manager.get_api_key("gpt")
        
        if claude_key and gpt_key:
            print("✅ API密钥配置正常")
        else:
            print("❌ API密钥配置异常")
            return False
        
        # 测试模型列表
        models = config_manager.get_available_models()
        if "claude" in models and "gpt" in models:
            print("✅ 模型列表正常")
        else:
            print("❌ 模型列表异常")
            return False
        
        # 测试处理控制器
        print("[3/4] 测试处理控制器...")
        processor_controller = DocumentProcessorController(config_manager)
        print("✅ 处理控制器初始化成功")
        
        # 测试错误处理器
        print("[4/4] 测试错误处理器...")
        error_handler = ErrorHandler()
        print("✅ 错误处理器初始化成功")
        
        print()
        print("🎉 GUI启动测试通过！")
        print()
        print("GUI功能确认:")
        print("✅ 所有组件导入正常")
        print("✅ 配置管理器正常")
        print("✅ API密钥已预配置")
        print("✅ 模型列表正常")
        print("✅ 处理控制器正常")
        print("✅ 错误处理器正常")
        print()
        print("现在可以安全启动GUI界面！")
        
        return True
        
    except Exception as e:
        print(f"❌ GUI启动测试失败: {e}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == "__main__":
    try:
        success = test_gui_startup()
        if success:
            print()
            print("✅ GUI启动测试通过！")
            print("   现在可以安全使用 'AutoWord启动器.bat'")
            sys.exit(0)
        else:
            print()
            print("❌ GUI启动测试失败。")
            sys.exit(1)
    except KeyboardInterrupt:
        print("\n测试被用户中断")
        sys.exit(1)
    except Exception as e:
        print(f"测试过程中发生异常: {e}")
        sys.exit(1)