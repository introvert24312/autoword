"""
Main Window
主窗口 - AutoWord GUI应用程序的主界面
"""

import logging
from typing import Optional
from pathlib import Path

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QComboBox, QLineEdit, QPushButton, QTextEdit,
    QProgressBar, QFileDialog, QMessageBox, QFrame,
    QSizePolicy, QSpacerItem, QMainWindow
)
from PySide6.QtCore import Qt, Signal, QTimer, QRect
from PySide6.QtGui import QFont, QIcon, QPalette

from .config_manager import ConfigurationManager
from .processor_controller import DocumentProcessorController
from .error_handler import ErrorHandler
from .progress_widget import ProgressWidget


logger = logging.getLogger(__name__)


class MainWindow(QMainWindow):
    """主窗口类"""
    
    def __init__(self, config_manager: ConfigurationManager):
        super().__init__()
        
        self.config_manager = config_manager
        self.processor_controller = None
        self.error_handler = ErrorHandler()
        
        # 设置窗口属性
        self.setWindowTitle("AutoWord - 文档自动化处理工具")
        self.setMinimumSize(800, 600)
        
        # 加载窗口几何信息
        geometry = self.config_manager.get_window_geometry()
        self.setGeometry(geometry['x'], geometry['y'], geometry['width'], geometry['height'])
        
        # 初始化UI
        self.setup_ui()
        self.setup_connections()
        self.load_settings()
        
        logger.info("主窗口初始化完成")
    
    def setup_ui(self):
        """设置用户界面"""
        # 创建中央窗口部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 主布局
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)
        
        # 创建各个区域
        self.create_config_section(main_layout)
        self.create_file_section(main_layout)
        self.create_control_section(main_layout)
        self.create_progress_section(main_layout)
        self.create_log_section(main_layout)
        
        # 设置样式
        self.apply_styles()
    
    def create_config_section(self, parent_layout):
        """创建配置区域"""
        # 配置框架
        config_frame = QFrame()
        config_frame.setFrameStyle(QFrame.Box)
        config_frame.setStyleSheet("QFrame { border: 1px solid #3c3c3c; border-radius: 8px; padding: 10px; }")
        
        config_layout = QGridLayout(config_frame)
        config_layout.setSpacing(10)
        
        # 模型选择
        model_label = QLabel("AI模型:")
        model_label.setFont(QFont("Microsoft YaHei", 10))
        self.model_combo = QComboBox()
        self.model_combo.setMinimumHeight(35)
        
        # 添加模型选项
        models = self.config_manager.get_available_models()
        for model_key, model_name in models.items():
            self.model_combo.addItem(model_name, model_key)
        
        # API密钥输入
        api_label = QLabel("API密钥:")
        api_label.setFont(QFont("Microsoft YaHei", 10))
        self.api_key_input = QLineEdit()
        self.api_key_input.setMinimumHeight(35)
        self.api_key_input.setEchoMode(QLineEdit.Password)
        self.api_key_input.setPlaceholderText("请输入API密钥...")
        
        # 显示/隐藏密钥按钮
        self.show_key_btn = QPushButton("👁")
        self.show_key_btn.setMaximumWidth(40)
        self.show_key_btn.setMinimumHeight(35)
        self.show_key_btn.setCheckable(True)
        self.show_key_btn.setToolTip("显示/隐藏API密钥")
        
        # 布局
        config_layout.addWidget(model_label, 0, 0)
        config_layout.addWidget(self.model_combo, 0, 1, 1, 2)
        config_layout.addWidget(api_label, 1, 0)
        config_layout.addWidget(self.api_key_input, 1, 1)
        config_layout.addWidget(self.show_key_btn, 1, 2)
        
        parent_layout.addWidget(config_frame)
    
    def create_file_section(self, parent_layout):
        """创建文件选择区域"""
        # 文件框架
        file_frame = QFrame()
        file_frame.setFrameStyle(QFrame.Box)
        file_frame.setStyleSheet("QFrame { border: 1px solid #3c3c3c; border-radius: 8px; padding: 10px; }")
        
        file_layout = QGridLayout(file_frame)
        file_layout.setSpacing(10)
        
        # 输入文件
        input_label = QLabel("输入文件:")
        input_label.setFont(QFont("Microsoft YaHei", 10))
        self.input_path_display = QLineEdit()
        self.input_path_display.setMinimumHeight(35)
        self.input_path_display.setReadOnly(True)
        self.input_path_display.setPlaceholderText("请选择Word文档...")
        
        self.select_input_btn = QPushButton("选择文件")
        self.select_input_btn.setMinimumHeight(35)
        self.select_input_btn.setMinimumWidth(100)
        
        # 输出文件
        output_label = QLabel("输出文件:")
        output_label.setFont(QFont("Microsoft YaHei", 10))
        self.output_path_display = QLineEdit()
        self.output_path_display.setMinimumHeight(35)
        self.output_path_display.setReadOnly(True)
        self.output_path_display.setPlaceholderText("自动生成输出文件名...")
        
        self.select_output_btn = QPushButton("选择文件")
        self.select_output_btn.setMinimumHeight(35)
        self.select_output_btn.setMinimumWidth(100)
        
        # 布局
        file_layout.addWidget(input_label, 0, 0)
        file_layout.addWidget(self.input_path_display, 0, 1)
        file_layout.addWidget(self.select_input_btn, 0, 2)
        file_layout.addWidget(output_label, 1, 0)
        file_layout.addWidget(self.output_path_display, 1, 1)
        file_layout.addWidget(self.select_output_btn, 1, 2)
        
        parent_layout.addWidget(file_frame)
    
    def create_control_section(self, parent_layout):
        """创建控制按钮区域"""
        control_layout = QHBoxLayout()
        
        # 添加弹性空间
        control_layout.addStretch()
        
        # 开始处理按钮
        self.start_btn = QPushButton("开始处理")
        self.start_btn.setMinimumHeight(45)
        self.start_btn.setMinimumWidth(150)
        self.start_btn.setFont(QFont("Microsoft YaHei", 12, QFont.Bold))
        self.start_btn.setStyleSheet("""
            QPushButton {
                background-color: #0078d4;
                color: white;
                border: none;
                border-radius: 8px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #106ebe;
            }
            QPushButton:pressed {
                background-color: #005a9e;
            }
            QPushButton:disabled {
                background-color: #3c3c3c;
                color: #666666;
            }
        """)
        
        # 停止处理按钮
        self.stop_btn = QPushButton("停止处理")
        self.stop_btn.setMinimumHeight(45)
        self.stop_btn.setMinimumWidth(150)
        self.stop_btn.setFont(QFont("Microsoft YaHei", 12))
        self.stop_btn.setEnabled(False)
        self.stop_btn.setStyleSheet("""
            QPushButton {
                background-color: #d13438;
                color: white;
                border: none;
                border-radius: 8px;
            }
            QPushButton:hover {
                background-color: #b71c1c;
            }
            QPushButton:pressed {
                background-color: #8e0000;
            }
            QPushButton:disabled {
                background-color: #3c3c3c;
                color: #666666;
            }
        """)
        
        control_layout.addWidget(self.start_btn)
        control_layout.addSpacing(10)
        control_layout.addWidget(self.stop_btn)
        control_layout.addStretch()
        
        parent_layout.addLayout(control_layout)
    
    def create_progress_section(self, parent_layout):
        """创建进度显示区域"""
        # 进度框架
        progress_frame = QFrame()
        progress_frame.setFrameStyle(QFrame.Box)
        progress_frame.setStyleSheet("QFrame { border: 1px solid #3c3c3c; border-radius: 8px; padding: 10px; }")
        
        progress_layout = QVBoxLayout(progress_frame)
        progress_layout.setSpacing(8)
        
        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimumHeight(25)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 1px solid #3c3c3c;
                border-radius: 8px;
                text-align: center;
                font-weight: bold;
            }
            QProgressBar::chunk {
                background-color: #0078d4;
                border-radius: 7px;
            }
        """)
        
        # 状态标签
        self.status_label = QLabel("就绪")
        self.status_label.setFont(QFont("Microsoft YaHei", 10))
        self.status_label.setStyleSheet("color: #888888;")
        
        progress_layout.addWidget(self.progress_bar)
        progress_layout.addWidget(self.status_label)
        
        parent_layout.addWidget(progress_frame)
    
    def create_log_section(self, parent_layout):
        """创建日志显示区域"""
        # 日志框架
        log_frame = QFrame()
        log_frame.setFrameStyle(QFrame.Box)
        log_frame.setStyleSheet("QFrame { border: 1px solid #3c3c3c; border-radius: 8px; padding: 10px; }")
        
        log_layout = QVBoxLayout(log_frame)
        log_layout.setSpacing(8)
        
        # 日志标题和按钮
        log_header = QHBoxLayout()
        log_title = QLabel("处理日志:")
        log_title.setFont(QFont("Microsoft YaHei", 10, QFont.Bold))
        
        self.save_log_btn = QPushButton("保存日志")
        self.save_log_btn.setMaximumWidth(100)
        self.save_log_btn.setMinimumHeight(30)
        
        self.clear_log_btn = QPushButton("清空日志")
        self.clear_log_btn.setMaximumWidth(100)
        self.clear_log_btn.setMinimumHeight(30)
        
        log_header.addWidget(log_title)
        log_header.addStretch()
        log_header.addWidget(self.save_log_btn)
        log_header.addWidget(self.clear_log_btn)
        
        # 日志文本区域
        self.log_text = QTextEdit()
        self.log_text.setMinimumHeight(200)
        self.log_text.setReadOnly(True)
        self.log_text.setFont(QFont("Consolas", 9))
        self.log_text.setStyleSheet("""
            QTextEdit {
                background-color: #1e1e1e;
                color: #ffffff;
                border: 1px solid #3c3c3c;
                border-radius: 4px;
                padding: 8px;
            }
        """)
        
        log_layout.addLayout(log_header)
        log_layout.addWidget(self.log_text)
        
        parent_layout.addWidget(log_frame)
    
    def apply_styles(self):
        """应用样式"""
        # 设置整体样式
        self.setStyleSheet("""
            QWidget {
                background-color: #2b2b2b;
                color: #ffffff;
                font-family: "Microsoft YaHei";
            }
            
            QLabel {
                color: #ffffff;
            }
            
            QComboBox {
                background-color: #3c3c3c;
                border: 1px solid #555555;
                border-radius: 4px;
                padding: 5px;
                color: #ffffff;
            }
            
            QComboBox::drop-down {
                border: none;
            }
            
            QComboBox::down-arrow {
                image: none;
                border-left: 5px solid transparent;
                border-right: 5px solid transparent;
                border-top: 5px solid #ffffff;
                margin-right: 10px;
            }
            
            QLineEdit {
                background-color: #3c3c3c;
                border: 1px solid #555555;
                border-radius: 4px;
                padding: 8px;
                color: #ffffff;
            }
            
            QLineEdit:focus {
                border-color: #0078d4;
            }
            
            QPushButton {
                background-color: #404040;
                border: 1px solid #555555;
                border-radius: 4px;
                padding: 8px 16px;
                color: #ffffff;
            }
            
            QPushButton:hover {
                background-color: #4a4a4a;
            }
            
            QPushButton:pressed {
                background-color: #353535;
            }
            
            QPushButton:disabled {
                background-color: #2a2a2a;
                color: #666666;
            }
        """)
    
    def setup_connections(self):
        """设置信号连接"""
        # 模型选择变化
        self.model_combo.currentTextChanged.connect(self.on_model_changed)
        
        # API密钥输入
        self.api_key_input.textChanged.connect(self.on_api_key_changed)
        self.show_key_btn.toggled.connect(self.on_show_key_toggled)
        
        # 文件选择
        self.select_input_btn.clicked.connect(self.select_input_file)
        self.select_output_btn.clicked.connect(self.select_output_path)
        
        # 控制按钮
        self.start_btn.clicked.connect(self.start_processing)
        self.stop_btn.clicked.connect(self.stop_processing)
        
        # 日志按钮
        self.save_log_btn.clicked.connect(self.save_log)
        self.clear_log_btn.clicked.connect(self.clear_log)
    
    def load_settings(self):
        """加载设置"""
        # 加载模型选择
        selected_model = self.config_manager.get_selected_model()
        for i in range(self.model_combo.count()):
            if self.model_combo.itemData(i) == selected_model:
                self.model_combo.setCurrentIndex(i)
                break
        
        # 加载API密钥
        api_key = self.config_manager.get_api_key(selected_model)
        self.api_key_input.setText(api_key)
        
        # 加载路径
        paths = self.config_manager.get_last_paths()
        if paths['input']:
            self.input_path_display.setText(paths['input'])
        if paths['output']:
            self.output_path_display.setText(paths['output'])
    
    def on_model_changed(self):
        """模型选择变化处理"""
        current_model = self.model_combo.currentData()
        if current_model:
            # 保存当前API密钥
            if self.api_key_input.text():
                old_model = self.config_manager.get_selected_model()
                self.config_manager.set_api_key(old_model, self.api_key_input.text())
            
            # 更新选中的模型
            self.config_manager.set_selected_model(current_model)
            
            # 加载新模型的API密钥
            api_key = self.config_manager.get_api_key(current_model)
            self.api_key_input.setText(api_key)
    
    def on_api_key_changed(self):
        """API密钥变化处理"""
        current_model = self.model_combo.currentData()
        if current_model:
            self.config_manager.set_api_key(current_model, self.api_key_input.text())
    
    def on_show_key_toggled(self, checked):
        """显示/隐藏密钥切换"""
        if checked:
            self.api_key_input.setEchoMode(QLineEdit.Normal)
            self.show_key_btn.setText("🙈")
        else:
            self.api_key_input.setEchoMode(QLineEdit.Password)
            self.show_key_btn.setText("👁")
    
    def select_input_file(self):
        """选择输入文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择Word文档",
            self.config_manager.get_last_paths()['input'],
            "Word文档 (*.docx *.doc);;所有文件 (*.*)"
        )
        
        if file_path:
            self.input_path_display.setText(file_path)
            self.config_manager.set_last_paths(input_path=file_path)
            
            # 自动生成输出文件名：原文件名.process.扩展名
            input_path = Path(file_path)
            output_filename = f"{input_path.stem}.process{input_path.suffix}"
            output_file_path = input_path.parent / output_filename
            self.output_path_display.setText(str(output_file_path))
            self.config_manager.set_last_paths(output_path=str(output_file_path))
    
    def select_output_path(self):
        """选择输出文件"""
        # 获取当前输出路径作为默认路径
        current_output = self.output_path_display.text()
        default_dir = str(Path(current_output).parent) if current_output else ""
        default_name = Path(current_output).name if current_output else ""
        
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "选择输出文件",
            str(Path(default_dir) / default_name) if default_dir and default_name else "",
            "Word文档 (*.docx);;Word文档 (*.doc);;所有文件 (*.*)"
        )
        
        if file_path:
            self.output_path_display.setText(file_path)
            self.config_manager.set_last_paths(output_path=file_path)
    
    def start_processing(self):
        """开始处理"""
        # 验证配置
        errors = self.validate_inputs()
        if errors:
            self.error_handler.show_error_dialog(
                "配置错误",
                "请检查以下配置项：\n" + "\n".join(f"• {error}" for error in errors)
            )
            return
        
        # 创建处理控制器
        if not self.processor_controller:
            self.processor_controller = DocumentProcessorController(self.config_manager)
            self.setup_processor_connections()
        
        # 获取参数
        input_path = self.input_path_display.text()
        output_path = self.output_path_display.text()
        model_type = self.model_combo.currentData()
        api_key = self.api_key_input.text()
        
        # 如果没有输出路径，自动生成
        if not output_path:
            input_file = Path(input_path)
            output_path = str(input_file.parent / f"{input_file.stem}.process{input_file.suffix}")
        
        # 更新UI状态
        self.set_processing_state(True)
        
        # 开始处理
        self.processor_controller.start_processing(input_path, output_path, model_type, api_key)
        
        self.add_log("开始处理文档...")
    
    def stop_processing(self):
        """停止处理"""
        if self.processor_controller:
            self.processor_controller.stop_processing()
        
        self.set_processing_state(False)
        self.add_log("用户取消处理")
    
    def validate_inputs(self) -> list:
        """验证输入"""
        errors = []
        
        # 检查输入文件
        if not self.input_path_display.text():
            errors.append("请选择输入文件")
        elif not Path(self.input_path_display.text()).exists():
            errors.append("输入文件不存在")
        
        # 检查API密钥
        if not self.api_key_input.text().strip():
            errors.append("请输入API密钥")
        
        # 检查模型选择
        if not self.model_combo.currentData():
            errors.append("请选择AI模型")
        
        return errors
    
    def set_processing_state(self, processing: bool):
        """设置处理状态"""
        # 更新按钮状态
        self.start_btn.setEnabled(not processing)
        self.stop_btn.setEnabled(processing)
        
        # 更新输入控件状态
        self.model_combo.setEnabled(not processing)
        self.api_key_input.setEnabled(not processing)
        self.select_input_btn.setEnabled(not processing)
        self.select_output_btn.setEnabled(not processing)
        
        if not processing:
            self.progress_bar.setValue(0)
            self.status_label.setText("就绪")
    
    def setup_processor_connections(self):
        """设置处理器信号连接"""
        if self.processor_controller:
            self.processor_controller.progress_updated.connect(self.on_progress_updated)
            self.processor_controller.processing_completed.connect(self.on_processing_completed)
            self.processor_controller.error_occurred.connect(self.on_error_occurred)
    
    def on_progress_updated(self, stage: str, progress: float):
        """进度更新处理"""
        self.progress_bar.setValue(int(progress * 100))
        self.status_label.setText(f"正在{stage}...")
        self.add_log(f"[{stage}] 进度: {progress:.1%}")
    
    def on_processing_completed(self, success: bool, message: str):
        """处理完成处理"""
        self.set_processing_state(False)
        
        if success:
            self.add_log(f"处理完成: {message}")
            QMessageBox.information(self, "处理完成", message)
        else:
            self.add_log(f"处理失败: {message}")
            self.error_handler.show_error_dialog("处理失败", message)
    
    def on_error_occurred(self, error_type: str, error_message: str):
        """错误发生处理"""
        self.set_processing_state(False)
        self.add_log(f"错误: {error_message}")
        self.error_handler.show_error_dialog(error_type, error_message)
    
    def add_log(self, message: str):
        """添加日志"""
        from datetime import datetime
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = f"[{timestamp}] {message}"
        self.log_text.append(log_entry)
        
        # 自动滚动到底部
        cursor = self.log_text.textCursor()
        cursor.movePosition(cursor.MoveOperation.End)
        self.log_text.setTextCursor(cursor)
    
    def save_log(self):
        """保存日志"""
        from datetime import datetime
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "保存日志",
            f"autoword_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
            "文本文件 (*.txt);;所有文件 (*.*)"
        )
        
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(self.log_text.toPlainText())
                QMessageBox.information(self, "保存成功", f"日志已保存到: {file_path}")
            except Exception as e:
                self.error_handler.show_error_dialog("保存失败", f"无法保存日志文件: {str(e)}")
    
    def clear_log(self):
        """清空日志"""
        self.log_text.clear()
    
    def closeEvent(self, event):
        """窗口关闭事件"""
        # 保存窗口几何信息
        geometry = self.geometry()
        self.config_manager.set_window_geometry(
            geometry.x(), geometry.y(), geometry.width(), geometry.height()
        )
        
        # 停止处理
        if self.processor_controller:
            self.processor_controller.stop_processing()
        
        event.accept()
        logger.info("主窗口已关闭")