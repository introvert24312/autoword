"""
Error Handler
错误处理器 - 统一处理和显示错误信息
"""

import logging
import traceback
from typing import Optional, Dict, Any
from datetime import datetime

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QTextEdit, QScrollArea, QWidget, QFrame, QApplication
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont, QIcon, QPixmap


logger = logging.getLogger(__name__)


class ErrorDialog(QDialog):
    """错误对话框"""
    
    def __init__(self, error_type: str, user_message: str, technical_details: str = None, parent=None):
        super().__init__(parent)
        
        self.error_type = error_type
        self.user_message = user_message
        self.technical_details = technical_details or ""
        
        self.setup_ui()
        self.setup_connections()
        
        # 设置窗口属性
        self.setWindowTitle("错误报告")
        self.setModal(True)
        self.setMinimumSize(600, 400)
        self.resize(700, 500)
    
    def setup_ui(self):
        """设置用户界面"""
        layout = QVBoxLayout(self)
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # 创建头部区域
        self.create_header(layout)
        
        # 创建消息区域
        self.create_message_area(layout)
        
        # 创建建议区域
        self.create_suggestions_area(layout)
        
        # 创建技术详情区域
        self.create_details_area(layout)
        
        # 创建按钮区域
        self.create_buttons_area(layout)
        
        # 应用样式
        self.apply_styles()
    
    def create_header(self, parent_layout):
        """创建头部区域"""
        header_layout = QHBoxLayout()
        
        # 错误图标
        icon_label = QLabel()
        icon_label.setText("⚠️")
        icon_label.setFont(QFont("Segoe UI Emoji", 24))
        icon_label.setFixedSize(40, 40)
        icon_label.setAlignment(Qt.AlignCenter)
        
        # 标题
        title_label = QLabel("处理错误")
        title_label.setFont(QFont("Microsoft YaHei", 16, QFont.Bold))
        title_label.setStyleSheet("color: #f38ba8;")
        
        header_layout.addWidget(icon_label)
        header_layout.addWidget(title_label)
        header_layout.addStretch()
        
        parent_layout.addLayout(header_layout)
    
    def create_message_area(self, parent_layout):
        """创建消息区域"""
        # 主要错误消息
        message_label = QLabel("文档处理过程中发生错误，请检查以下信息：")
        message_label.setFont(QFont("Microsoft YaHei", 10))
        message_label.setWordWrap(True)
        
        # 错误类型和描述
        error_frame = QFrame()
        error_frame.setFrameStyle(QFrame.Box)
        error_frame.setStyleSheet("""
            QFrame {
                border: 1px solid #f38ba8;
                border-radius: 8px;
                background-color: #2d1b1b;
                padding: 10px;
            }
        """)
        
        error_layout = QVBoxLayout(error_frame)
        
        # 错误类型
        type_label = QLabel(f"错误类型: {self.get_error_type_display()}")
        type_label.setFont(QFont("Microsoft YaHei", 10, QFont.Bold))
        type_label.setStyleSheet("color: #f38ba8;")
        
        # 错误描述
        desc_label = QLabel(f"错误描述: {self.user_message}")
        desc_label.setFont(QFont("Microsoft YaHei", 10))
        desc_label.setWordWrap(True)
        desc_label.setStyleSheet("color: #ffffff;")
        
        error_layout.addWidget(type_label)
        error_layout.addWidget(desc_label)
        
        parent_layout.addWidget(message_label)
        parent_layout.addWidget(error_frame)
    
    def create_suggestions_area(self, parent_layout):
        """创建建议区域"""
        suggestions = self.get_error_suggestions()
        if not suggestions:
            return
        
        suggestions_label = QLabel("建议解决方案:")
        suggestions_label.setFont(QFont("Microsoft YaHei", 10, QFont.Bold))
        
        suggestions_frame = QFrame()
        suggestions_frame.setFrameStyle(QFrame.Box)
        suggestions_frame.setStyleSheet("""
            QFrame {
                border: 1px solid #89b4fa;
                border-radius: 8px;
                background-color: #1b1d2d;
                padding: 10px;
            }
        """)
        
        suggestions_layout = QVBoxLayout(suggestions_frame)
        
        for suggestion in suggestions:
            suggestion_label = QLabel(f"• {suggestion}")
            suggestion_label.setFont(QFont("Microsoft YaHei", 9))
            suggestion_label.setWordWrap(True)
            suggestion_label.setStyleSheet("color: #ffffff;")
            suggestions_layout.addWidget(suggestion_label)
        
        parent_layout.addWidget(suggestions_label)
        parent_layout.addWidget(suggestions_frame)
    
    def create_details_area(self, parent_layout):
        """创建技术详情区域"""
        if not self.technical_details:
            return
        
        details_label = QLabel("技术详情:")
        details_label.setFont(QFont("Microsoft YaHei", 10, QFont.Bold))
        
        self.details_text = QTextEdit()
        self.details_text.setPlainText(self.technical_details)
        self.details_text.setReadOnly(True)
        self.details_text.setFont(QFont("Consolas", 9))
        self.details_text.setMaximumHeight(150)
        self.details_text.setStyleSheet("""
            QTextEdit {
                background-color: #1e1e1e;
                color: #ffffff;
                border: 1px solid #3c3c3c;
                border-radius: 4px;
                padding: 8px;
            }
        """)
        
        parent_layout.addWidget(details_label)
        parent_layout.addWidget(self.details_text)
    
    def create_buttons_area(self, parent_layout):
        """创建按钮区域"""
        buttons_layout = QHBoxLayout()
        buttons_layout.addStretch()
        
        # 复制错误信息按钮
        self.copy_btn = QPushButton("复制错误信息")
        self.copy_btn.setMinimumHeight(35)
        self.copy_btn.setMinimumWidth(120)
        
        # 保存日志按钮
        self.save_btn = QPushButton("保存日志")
        self.save_btn.setMinimumHeight(35)
        self.save_btn.setMinimumWidth(100)
        
        # 确定按钮
        self.ok_btn = QPushButton("确定")
        self.ok_btn.setMinimumHeight(35)
        self.ok_btn.setMinimumWidth(80)
        self.ok_btn.setDefault(True)
        
        buttons_layout.addWidget(self.copy_btn)
        buttons_layout.addWidget(self.save_btn)
        buttons_layout.addWidget(self.ok_btn)
        
        parent_layout.addLayout(buttons_layout)
    
    def apply_styles(self):
        """应用样式"""
        self.setStyleSheet("""
            QDialog {
                background-color: #2b2b2b;
                color: #ffffff;
            }
            
            QLabel {
                color: #ffffff;
            }
            
            QPushButton {
                background-color: #404040;
                border: 1px solid #555555;
                border-radius: 4px;
                padding: 8px 16px;
                color: #ffffff;
                font-family: "Microsoft YaHei";
            }
            
            QPushButton:hover {
                background-color: #4a4a4a;
            }
            
            QPushButton:pressed {
                background-color: #353535;
            }
            
            QPushButton:default {
                background-color: #0078d4;
                border-color: #0078d4;
            }
            
            QPushButton:default:hover {
                background-color: #106ebe;
            }
        """)
    
    def setup_connections(self):
        """设置信号连接"""
        self.copy_btn.clicked.connect(self.copy_error_info)
        self.save_btn.clicked.connect(self.save_error_log)
        self.ok_btn.clicked.connect(self.accept)
    
    def get_error_type_display(self) -> str:
        """获取错误类型显示名称"""
        type_names = {
            "ValidationError": "配置验证错误",
            "ProcessingError": "文档处理错误",
            "NetworkError": "网络连接错误",
            "APIError": "API调用错误",
            "FileError": "文件操作错误",
            "UnexpectedError": "未知错误",
            "StartupError": "启动错误"
        }
        return type_names.get(self.error_type, self.error_type)
    
    def get_error_suggestions(self) -> list:
        """获取错误解决建议"""
        suggestions_map = {
            "ValidationError": [
                "检查API密钥是否正确输入",
                "确认选择的模型类型是否正确",
                "验证输入文件路径是否存在"
            ],
            "ProcessingError": [
                "检查Word文档是否已被其他程序占用",
                "确认文档格式是否为标准的.docx或.doc格式",
                "尝试重新启动应用程序"
            ],
            "NetworkError": [
                "检查网络连接是否正常",
                "确认防火墙是否阻止了网络访问",
                "尝试稍后重试"
            ],
            "APIError": [
                "检查API密钥是否有效且未过期",
                "确认API服务是否正常运行",
                "检查API调用频率是否超出限制"
            ],
            "FileError": [
                "检查文件是否存在且可访问",
                "确认是否有足够的磁盘空间",
                "验证文件权限是否正确"
            ],
            "UnexpectedError": [
                "尝试重新启动应用程序",
                "检查系统资源是否充足",
                "联系技术支持获取帮助"
            ],
            "StartupError": [
                "检查所有依赖是否正确安装",
                "确认Microsoft Word是否已安装",
                "尝试以管理员权限运行"
            ]
        }
        return suggestions_map.get(self.error_type, [])
    
    def copy_error_info(self):
        """复制错误信息到剪贴板"""
        error_info = self.format_error_info()
        
        clipboard = QApplication.clipboard()
        clipboard.setText(error_info)
        
        # 临时更改按钮文本
        original_text = self.copy_btn.text()
        self.copy_btn.setText("已复制!")
        
        # 2秒后恢复原文本
        from PySide6.QtCore import QTimer
        QTimer.singleShot(2000, lambda: self.copy_btn.setText(original_text))
    
    def save_error_log(self):
        """保存错误日志到文件"""
        from PySide6.QtWidgets import QFileDialog
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_filename = f"autoword_error_{timestamp}.txt"
        
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "保存错误日志",
            default_filename,
            "文本文件 (*.txt);;所有文件 (*.*)"
        )
        
        if file_path:
            try:
                error_info = self.format_error_info()
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(error_info)
                
                # 临时更改按钮文本
                original_text = self.save_btn.text()
                self.save_btn.setText("已保存!")
                
                # 2秒后恢复原文本
                from PySide6.QtCore import QTimer
                QTimer.singleShot(2000, lambda: self.save_btn.setText(original_text))
                
            except Exception as e:
                from PySide6.QtWidgets import QMessageBox
                QMessageBox.warning(self, "保存失败", f"无法保存错误日志: {str(e)}")
    
    def format_error_info(self) -> str:
        """格式化错误信息"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        info_lines = [
            "AutoWord 错误报告",
            "=" * 50,
            f"时间: {timestamp}",
            f"错误类型: {self.get_error_type_display()}",
            f"错误描述: {self.user_message}",
            ""
        ]
        
        # 添加建议
        suggestions = self.get_error_suggestions()
        if suggestions:
            info_lines.append("建议解决方案:")
            for suggestion in suggestions:
                info_lines.append(f"• {suggestion}")
            info_lines.append("")
        
        # 添加技术详情
        if self.technical_details:
            info_lines.extend([
                "技术详情:",
                "-" * 30,
                self.technical_details,
                ""
            ])
        
        info_lines.extend([
            "=" * 50,
            "如需技术支持，请将此错误报告发送给开发团队。"
        ])
        
        return "\n".join(info_lines)


class ErrorHandler:
    """错误处理器"""
    
    def __init__(self):
        self.error_history = []
        logger.info("错误处理器初始化完成")
    
    def show_error_dialog(self, error_type: str, error_message: str, details: str = None, parent=None):
        """显示错误对话框"""
        try:
            # 记录错误
            self.log_error_info(error_type, error_message, details)
            
            # 创建并显示错误对话框
            dialog = ErrorDialog(error_type, error_message, details, parent)
            dialog.exec()
            
        except Exception as e:
            logger.error(f"显示错误对话框失败: {e}")
            # 降级到简单消息框
            from PySide6.QtWidgets import QMessageBox
            QMessageBox.critical(parent, "错误", f"{error_type}: {error_message}")
    
    def log_error(self, error: Exception, context: str = None):
        """记录异常错误"""
        error_type = type(error).__name__
        error_message = str(error)
        details = traceback.format_exc()
        
        if context:
            error_message = f"{context}: {error_message}"
        
        self.log_error_info(error_type, error_message, details)
        logger.error(f"异常错误 [{error_type}]: {error_message}")
    
    def log_error_info(self, error_type: str, error_message: str, details: str = None):
        """记录错误信息"""
        error_info = {
            "timestamp": datetime.now(),
            "type": error_type,
            "message": error_message,
            "details": details or ""
        }
        
        self.error_history.append(error_info)
        
        # 限制历史记录数量
        if len(self.error_history) > 100:
            self.error_history = self.error_history[-50:]
        
        logger.error(f"错误记录 [{error_type}]: {error_message}")
    
    def export_error_log(self, file_path: str):
        """导出错误日志"""
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write("AutoWord 错误日志\n")
                f.write("=" * 50 + "\n\n")
                
                for error_info in self.error_history:
                    f.write(f"时间: {error_info['timestamp'].strftime('%Y-%m-%d %H:%M:%S')}\n")
                    f.write(f"类型: {error_info['type']}\n")
                    f.write(f"消息: {error_info['message']}\n")
                    if error_info['details']:
                        f.write(f"详情:\n{error_info['details']}\n")
                    f.write("-" * 30 + "\n\n")
            
            logger.info(f"错误日志已导出到: {file_path}")
            
        except Exception as e:
            logger.error(f"导出错误日志失败: {e}")
            raise
    
    def get_error_statistics(self) -> Dict[str, Any]:
        """获取错误统计信息"""
        if not self.error_history:
            return {"total": 0, "by_type": {}, "recent": []}
        
        # 按类型统计
        by_type = {}
        for error_info in self.error_history:
            error_type = error_info['type']
            by_type[error_type] = by_type.get(error_type, 0) + 1
        
        # 最近的错误
        recent = self.error_history[-10:] if len(self.error_history) > 10 else self.error_history
        
        return {
            "total": len(self.error_history),
            "by_type": by_type,
            "recent": [
                {
                    "timestamp": error_info['timestamp'].strftime('%Y-%m-%d %H:%M:%S'),
                    "type": error_info['type'],
                    "message": error_info['message'][:100] + "..." if len(error_info['message']) > 100 else error_info['message']
                }
                for error_info in recent
            ]
        }
    
    def clear_error_history(self):
        """清空错误历史"""
        self.error_history.clear()
        logger.info("错误历史已清空")