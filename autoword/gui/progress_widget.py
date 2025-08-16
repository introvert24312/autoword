"""
Progress Widget
进度显示组件 - 显示处理进度和状态信息
"""

import logging
from typing import Optional
from datetime import datetime, timedelta

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, 
    QProgressBar, QFrame
)
from PySide6.QtCore import Qt, QTimer, Signal
from PySide6.QtGui import QFont, QPalette


logger = logging.getLogger(__name__)


class ProgressWidget(QWidget):
    """进度显示组件"""
    
    # 信号
    cancel_requested = Signal()  # 取消请求信号
    
    def __init__(self, parent=None):
        super().__init__(parent)
        
        self.start_time: Optional[datetime] = None
        self.current_stage = ""
        self.current_progress = 0.0
        self.estimated_total_time = 0.0
        
        # 定时器用于更新时间显示
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_time_display)
        
        self.setup_ui()
        self.reset()
        
        logger.debug("进度显示组件初始化完成")
    
    def setup_ui(self):
        """设置用户界面"""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)
        
        # 创建进度条
        self.create_progress_bar(layout)
        
        # 创建状态信息
        self.create_status_info(layout)
        
        # 创建时间信息
        self.create_time_info(layout)
        
        # 应用样式
        self.apply_styles()
    
    def create_progress_bar(self, parent_layout):
        """创建进度条"""
        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimumHeight(25)
        self.progress_bar.setMaximumHeight(25)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setAlignment(Qt.AlignCenter)
        
        parent_layout.addWidget(self.progress_bar)
    
    def create_status_info(self, parent_layout):
        """创建状态信息"""
        status_layout = QHBoxLayout()
        
        # 当前阶段标签
        stage_title = QLabel("当前阶段:")
        stage_title.setFont(QFont("Microsoft YaHei", 9))
        
        self.stage_label = QLabel("就绪")
        self.stage_label.setFont(QFont("Microsoft YaHei", 9, QFont.Bold))
        
        # 详细状态标签
        self.status_label = QLabel("等待开始...")
        self.status_label.setFont(QFont("Microsoft YaHei", 9))
        self.status_label.setStyleSheet("color: #888888;")
        
        status_layout.addWidget(stage_title)
        status_layout.addWidget(self.stage_label)
        status_layout.addStretch()
        status_layout.addWidget(self.status_label)
        
        parent_layout.addLayout(status_layout)
    
    def create_time_info(self, parent_layout):
        """创建时间信息"""
        time_layout = QHBoxLayout()
        
        # 已用时间
        elapsed_title = QLabel("已用时间:")
        elapsed_title.setFont(QFont("Microsoft YaHei", 9))
        
        self.elapsed_label = QLabel("00:00:00")
        self.elapsed_label.setFont(QFont("Consolas", 9))
        
        # 预计剩余时间
        remaining_title = QLabel("预计剩余:")
        remaining_title.setFont(QFont("Microsoft YaHei", 9))
        
        self.remaining_label = QLabel("--:--:--")
        self.remaining_label.setFont(QFont("Consolas", 9))
        
        time_layout.addWidget(elapsed_title)
        time_layout.addWidget(self.elapsed_label)
        time_layout.addStretch()
        time_layout.addWidget(remaining_title)
        time_layout.addWidget(self.remaining_label)
        
        parent_layout.addLayout(time_layout)
    
    def apply_styles(self):
        """应用样式"""
        self.setStyleSheet("""
            QProgressBar {
                border: 1px solid #3c3c3c;
                border-radius: 8px;
                text-align: center;
                font-weight: bold;
                background-color: #2a2a2a;
                color: #ffffff;
            }
            
            QProgressBar::chunk {
                background-color: qlineargradient(
                    x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 #0078d4,
                    stop: 1 #106ebe
                );
                border-radius: 7px;
            }
            
            QLabel {
                color: #ffffff;
            }
        """)
    
    def update_progress(self, stage: str, progress: float):
        """更新进度"""
        self.current_stage = stage
        self.current_progress = progress
        
        # 更新进度条
        progress_percent = int(progress * 100)
        self.progress_bar.setValue(progress_percent)
        
        # 更新阶段标签
        self.stage_label.setText(stage)
        
        # 更新状态标签
        if progress >= 1.0:
            self.status_label.setText("完成")
            self.status_label.setStyleSheet("color: #a6e3a1;")  # 绿色
        elif progress > 0:
            self.status_label.setText(f"进行中... ({progress:.1%})")
            self.status_label.setStyleSheet("color: #89b4fa;")  # 蓝色
        else:
            self.status_label.setText("准备中...")
            self.status_label.setStyleSheet("color: #f9e2af;")  # 黄色
        
        # 启动定时器（如果还没启动）
        if not self.timer.isActive() and progress > 0:
            self.start_time = datetime.now()
            self.timer.start(1000)  # 每秒更新一次
        
        # 更新时间显示
        self.update_time_display()
        
        logger.debug(f"进度更新: {stage} - {progress:.1%}")
    
    def set_status_message(self, message: str):
        """设置状态消息"""
        self.status_label.setText(message)
        logger.debug(f"状态消息: {message}")
    
    def show_completion_message(self, success: bool, message: str):
        """显示完成消息"""
        if success:
            self.progress_bar.setValue(100)
            self.stage_label.setText("完成")
            self.status_label.setText(message)
            self.status_label.setStyleSheet("color: #a6e3a1;")  # 绿色
        else:
            self.stage_label.setText("失败")
            self.status_label.setText(message)
            self.status_label.setStyleSheet("color: #f38ba8;")  # 红色
        
        # 停止定时器
        self.timer.stop()
        
        logger.info(f"处理完成: {message}")
    
    def reset(self):
        """重置进度显示"""
        self.current_stage = ""
        self.current_progress = 0.0
        self.start_time = None
        self.estimated_total_time = 0.0
        
        # 重置UI
        self.progress_bar.setValue(0)
        self.stage_label.setText("就绪")
        self.status_label.setText("等待开始...")
        self.status_label.setStyleSheet("color: #888888;")
        self.elapsed_label.setText("00:00:00")
        self.remaining_label.setText("--:--:--")
        
        # 停止定时器
        self.timer.stop()
        
        logger.debug("进度显示已重置")
    
    def update_time_display(self):
        """更新时间显示"""
        if not self.start_time:
            return
        
        # 计算已用时间
        elapsed = datetime.now() - self.start_time
        elapsed_str = self.format_duration(elapsed)
        self.elapsed_label.setText(elapsed_str)
        
        # 计算预计剩余时间
        if self.current_progress > 0.1:  # 至少完成10%才开始预估
            total_estimated = elapsed.total_seconds() / self.current_progress
            remaining_seconds = total_estimated - elapsed.total_seconds()
            
            if remaining_seconds > 0:
                remaining = timedelta(seconds=remaining_seconds)
                remaining_str = self.format_duration(remaining)
                self.remaining_label.setText(remaining_str)
            else:
                self.remaining_label.setText("00:00:00")
        else:
            self.remaining_label.setText("--:--:--")
    
    def format_duration(self, duration: timedelta) -> str:
        """格式化时间间隔"""
        total_seconds = int(duration.total_seconds())
        hours = total_seconds // 3600
        minutes = (total_seconds % 3600) // 60
        seconds = total_seconds % 60
        
        return f"{hours:02d}:{minutes:02d}:{seconds:02d}"
    
    def set_error_state(self, error_message: str):
        """设置错误状态"""
        self.stage_label.setText("错误")
        self.status_label.setText(error_message)
        self.status_label.setStyleSheet("color: #f38ba8;")  # 红色
        
        # 停止定时器
        self.timer.stop()
        
        logger.error(f"进度显示错误状态: {error_message}")
    
    def get_progress_info(self) -> dict:
        """获取进度信息"""
        elapsed_time = 0.0
        if self.start_time:
            elapsed_time = (datetime.now() - self.start_time).total_seconds()
        
        return {
            "stage": self.current_stage,
            "progress": self.current_progress,
            "elapsed_time": elapsed_time,
            "start_time": self.start_time.isoformat() if self.start_time else None,
            "is_running": self.timer.isActive()
        }
    
    def set_indeterminate_progress(self, stage: str):
        """设置不确定进度模式"""
        self.current_stage = stage
        self.stage_label.setText(stage)
        self.status_label.setText("处理中...")
        self.status_label.setStyleSheet("color: #89b4fa;")
        
        # 设置进度条为不确定模式
        self.progress_bar.setRange(0, 0)
        
        logger.debug(f"设置不确定进度模式: {stage}")
    
    def set_determinate_progress(self):
        """设置确定进度模式"""
        self.progress_bar.setRange(0, 100)
        logger.debug("设置确定进度模式")


class StageProgressWidget(QWidget):
    """阶段进度组件 - 显示多个处理阶段的进度"""
    
    def __init__(self, stages: list, parent=None):
        super().__init__(parent)
        
        self.stages = stages
        self.current_stage_index = -1
        self.stage_widgets = {}
        
        self.setup_ui()
    
    def setup_ui(self):
        """设置用户界面"""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(5)
        
        # 创建每个阶段的进度指示器
        for i, stage in enumerate(self.stages):
            stage_widget = self.create_stage_indicator(stage, i)
            self.stage_widgets[stage] = stage_widget
            layout.addWidget(stage_widget)
    
    def create_stage_indicator(self, stage_name: str, index: int) -> QWidget:
        """创建阶段指示器"""
        widget = QFrame()
        widget.setFrameStyle(QFrame.Box)
        widget.setStyleSheet("""
            QFrame {
                border: 1px solid #3c3c3c;
                border-radius: 4px;
                padding: 5px;
                background-color: #2a2a2a;
            }
        """)
        
        layout = QHBoxLayout(widget)
        layout.setContentsMargins(8, 4, 8, 4)
        
        # 阶段编号
        number_label = QLabel(f"{index + 1}")
        number_label.setFixedSize(20, 20)
        number_label.setAlignment(Qt.AlignCenter)
        number_label.setStyleSheet("""
            QLabel {
                background-color: #404040;
                border-radius: 10px;
                color: #ffffff;
                font-weight: bold;
            }
        """)
        
        # 阶段名称
        name_label = QLabel(stage_name)
        name_label.setFont(QFont("Microsoft YaHei", 9))
        
        # 状态指示器
        status_label = QLabel("⏳")
        status_label.setFont(QFont("Segoe UI Emoji", 12))
        
        layout.addWidget(number_label)
        layout.addWidget(name_label)
        layout.addStretch()
        layout.addWidget(status_label)
        
        # 保存引用
        widget.number_label = number_label
        widget.name_label = name_label
        widget.status_label = status_label
        
        return widget
    
    def update_stage_progress(self, stage_name: str, status: str):
        """更新阶段进度"""
        if stage_name not in self.stage_widgets:
            return
        
        widget = self.stage_widgets[stage_name]
        
        # 更新状态图标
        status_icons = {
            "waiting": "⏳",
            "running": "🔄",
            "completed": "✅",
            "error": "❌"
        }
        
        icon = status_icons.get(status, "⏳")
        widget.status_label.setText(icon)
        
        # 更新样式
        if status == "running":
            widget.setStyleSheet("""
                QFrame {
                    border: 2px solid #89b4fa;
                    border-radius: 4px;
                    padding: 5px;
                    background-color: #1b1d2d;
                }
            """)
            widget.number_label.setStyleSheet("""
                QLabel {
                    background-color: #89b4fa;
                    border-radius: 10px;
                    color: #ffffff;
                    font-weight: bold;
                }
            """)
        elif status == "completed":
            widget.setStyleSheet("""
                QFrame {
                    border: 1px solid #a6e3a1;
                    border-radius: 4px;
                    padding: 5px;
                    background-color: #1b2d1b;
                }
            """)
            widget.number_label.setStyleSheet("""
                QLabel {
                    background-color: #a6e3a1;
                    border-radius: 10px;
                    color: #000000;
                    font-weight: bold;
                }
            """)
        elif status == "error":
            widget.setStyleSheet("""
                QFrame {
                    border: 1px solid #f38ba8;
                    border-radius: 4px;
                    padding: 5px;
                    background-color: #2d1b1b;
                }
            """)
            widget.number_label.setStyleSheet("""
                QLabel {
                    background-color: #f38ba8;
                    border-radius: 10px;
                    color: #ffffff;
                    font-weight: bold;
                }
            """)
        else:  # waiting
            widget.setStyleSheet("""
                QFrame {
                    border: 1px solid #3c3c3c;
                    border-radius: 4px;
                    padding: 5px;
                    background-color: #2a2a2a;
                }
            """)
            widget.number_label.setStyleSheet("""
                QLabel {
                    background-color: #404040;
                    border-radius: 10px;
                    color: #ffffff;
                    font-weight: bold;
                }
            """)
    
    def reset(self):
        """重置所有阶段状态"""
        for stage_name in self.stages:
            self.update_stage_progress(stage_name, "waiting")