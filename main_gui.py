#!/usr/bin/env python3
"""
AutoWord GUI Application Entry Point
AutoWord图形界面应用程序入口
"""

import sys
import os
import logging
from pathlib import Path

# 添加项目根目录到Python路径
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

try:
    from PySide6.QtWidgets import QApplication, QMessageBox
    from PySide6.QtCore import Qt, QDir
    from PySide6.QtGui import QIcon
    import qdarkstyle
except ImportError as e:
    print(f"GUI依赖缺失: {e}")
    print("请运行: pip install PySide6 qdarkstyle cryptography")
    sys.exit(1)

from autoword.gui import MainWindow, ConfigurationManager


def setup_application():
    """设置应用程序"""
    # 设置应用程序属性
    QApplication.setApplicationName("AutoWord")
    QApplication.setApplicationVersion("1.0.0")
    QApplication.setOrganizationName("AutoWord Team")
    QApplication.setOrganizationDomain("autoword.com")
    
    # 创建应用程序实例
    app = QApplication(sys.argv)
    
    # 设置应用程序图标
    icon_path = project_root / "autoword" / "gui" / "resources" / "icons" / "app.ico"
    if icon_path.exists():
        app.setWindowIcon(QIcon(str(icon_path)))
    
    # 应用深色主题
    app.setStyleSheet(qdarkstyle.load_stylesheet_pyside6())
    
    return app


def check_dependencies():
    """检查依赖和环境"""
    try:
        # 检查Word是否安装
        import win32com.client
        try:
            word_app = win32com.client.Dispatch("Word.Application")
            word_app.Quit()
        except Exception:
            QMessageBox.warning(
                None,
                "依赖检查",
                "未检测到Microsoft Word，请确保已安装Word应用程序。"
            )
            return False
        
        return True
        
    except ImportError:
        QMessageBox.critical(
            None,
            "依赖错误", 
            "缺少必要的依赖包，请运行: pip install -r requirements.txt"
        )
        return False


def main():
    """主函数"""
    try:
        # 设置日志
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
        logger = logging.getLogger(__name__)
        
        logger.info("启动AutoWord GUI应用程序")
        
        # 创建应用程序
        app = setup_application()
        
        # 检查依赖
        if not check_dependencies():
            return 1
        
        # 创建配置管理器
        config_manager = ConfigurationManager()
        
        # 创建主窗口
        main_window = MainWindow(config_manager)
        main_window.show()
        
        logger.info("GUI应用程序启动完成")
        
        # 运行应用程序
        return app.exec()
        
    except Exception as e:
        logging.error(f"应用程序启动失败: {e}")
        
        # 显示错误对话框
        if 'app' in locals():
            QMessageBox.critical(
                None,
                "启动错误",
                f"应用程序启动失败:\n{str(e)}\n\n请检查日志文件获取详细信息。"
            )
        else:
            print(f"应用程序启动失败: {e}")
        
        return 1


if __name__ == "__main__":
    sys.exit(main())