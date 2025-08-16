#!/usr/bin/env python3
"""
AutoWord GUI 快速启动脚本
用于开发和测试GUI界面
"""

import sys
import os
from pathlib import Path

# 添加项目根目录到Python路径
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

def main():
    """主函数"""
    try:
        # 检查GUI依赖
        try:
            from PySide6.QtWidgets import QApplication
            import qdarktheme
            from qframelesswindow import FramelessWindow
        except ImportError as e:
            print(f"GUI依赖缺失: {e}")
            print("请运行: pip install PySide6 qframelesswindow qdarktheme cryptography")
            return 1
        
        # 设置应用程序
        QApplication.setApplicationName("AutoWord")
        QApplication.setApplicationVersion("1.0.0")
        
        app = QApplication(sys.argv)
        
        # 应用深色主题
        qdarktheme.setup_theme("dark")
        
        # 创建配置管理器
        from autoword.gui.config_manager import ConfigurationManager
        config_manager = ConfigurationManager()
        
        # 创建主窗口
        from autoword.gui.main_window import MainWindow
        main_window = MainWindow(config_manager)
        main_window.show()
        
        print("AutoWord GUI 启动成功")
        
        # 运行应用程序
        return app.exec()
        
    except Exception as e:
        print(f"启动失败: {e}")
        import traceback
        traceback.print_exc()
        return 1

if __name__ == "__main__":
    sys.exit(main())