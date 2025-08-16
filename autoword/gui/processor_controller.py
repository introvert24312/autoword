"""
Document Processor Controller
文档处理控制器 - 协调GUI和AutoWord核心处理逻辑
"""

import logging
from typing import Optional
from datetime import datetime
from dataclasses import dataclass

from PySide6.QtCore import QObject, Signal, QThread

from autoword.core.pipeline import DocumentProcessor, PipelineConfig, PipelineProgress
from autoword.core.llm_client import ModelType, LLMClient
from autoword.core.word_executor import ExecutionMode
from .config_manager import ConfigurationManager


logger = logging.getLogger(__name__)


@dataclass
class ProcessingState:
    """处理状态"""
    is_running: bool = False
    current_stage: str = ""
    progress: float = 0.0
    status_message: str = ""
    start_time: Optional[datetime] = None
    estimated_completion: Optional[datetime] = None


class ProcessingThread(QThread):
    """处理线程"""
    
    progress_updated = Signal(str, float)  # (stage, progress)
    processing_completed = Signal(bool, str)  # (success, message)
    error_occurred = Signal(str, str)  # (error_type, error_message)
    
    def __init__(self, input_path: str, output_file_path: str, model_type: str, api_key: str):
        super().__init__()
        self.input_path = input_path
        self.output_file_path = output_file_path
        self.model_type = model_type
        self.api_key = api_key
        self.should_stop = False
        
    def run(self):
        """运行处理线程"""
        try:
            logger.info(f"开始处理文档: {self.input_path}")
            
            # 创建管道配置
            model_enum = ModelType.CLAUDE37 if self.model_type == "claude" else ModelType.GPT5
            
            # 从输出文件路径提取目录用于日志导出
            from pathlib import Path
            output_file = Path(self.output_file_path)
            output_dir = str(output_file.parent)
            
            config = PipelineConfig(
                model=model_enum,
                execution_mode=ExecutionMode.NORMAL,
                create_backup=True,
                enable_validation=True,
                export_results=True,
                output_dir=output_dir,
                visible_word=False,
                max_retries=3
            )
            
            # 准备API密钥 - 确保包含所有模型的密钥
            api_keys = {
                "claude": "sk-3w1JFbWUq7tKjpLlopdkISQ9F6fpLhHx5viD0frh43ESE9Io",
                "gpt": "sk-NhjnJtqlZMx4PGTqvkGlH4POT82HHBrBnBbWOat99Bs5VZXi"
            }
            # 覆盖当前选择的模型密钥
            api_keys[self.model_type] = self.api_key
            
            # 创建处理器
            processor = DocumentProcessor(config, api_keys=api_keys)
            
            # 添加进度回调
            processor.add_progress_callback(self._on_progress)
            
            # 执行处理
            result = processor.process_document(self.input_path, self.output_file_path)
            
            if self.should_stop:
                self.processing_completed.emit(False, "处理已取消")
                return
            
            if result.success:
                message = f"文档处理完成！\n耗时: {result.total_time:.2f}秒"
                if result.exported_files:
                    message += f"\n输出文件: {len(result.exported_files)} 个"
                self.processing_completed.emit(True, message)
            else:
                self.error_occurred.emit("ProcessingError", result.error_message or "处理失败")
                
        except Exception as e:
            logger.error(f"处理线程异常: {e}")
            self.error_occurred.emit("UnexpectedError", str(e))
        
        finally:
            logger.info("处理线程结束")
    
    def _on_progress(self, progress: PipelineProgress):
        """进度回调"""
        if not self.should_stop:
            stage_names = {
                "loading": "加载文档",
                "inspection": "分析文档",
                "planning": "生成计划",
                "execution": "执行任务",
                "validation": "验证结果",
                "export": "导出结果"
            }
            stage_name = stage_names.get(progress.stage, progress.stage)
            self.progress_updated.emit(stage_name, progress.progress)
    

    
    def stop(self):
        """停止处理"""
        self.should_stop = True
        logger.info("请求停止处理线程")


class DocumentProcessorController(QObject):
    """文档处理控制器"""
    
    # 信号
    progress_updated = Signal(str, float)  # (stage, progress)
    processing_completed = Signal(bool, str)  # (success, message)
    error_occurred = Signal(str, str)  # (error_type, error_message)
    
    def __init__(self, config_manager: ConfigurationManager):
        super().__init__()
        self.config_manager = config_manager
        self.processing_thread: Optional[ProcessingThread] = None
        self.state = ProcessingState()
        
        logger.info("文档处理控制器初始化完成")
    
    def start_processing(self, input_path: str, output_file_path: str, model_type: str, api_key: str):
        """开始处理"""
        if self.state.is_running:
            logger.warning("处理已在运行中")
            return
        
        try:
            # 验证参数
            validation_errors = self._validate_parameters(input_path, output_file_path, model_type, api_key)
            if validation_errors:
                self.error_occurred.emit("ValidationError", "\n".join(validation_errors))
                return
            
            # 更新状态
            self.state.is_running = True
            self.state.start_time = datetime.now()
            self.state.current_stage = "准备中"
            self.state.progress = 0.0
            
            # 创建处理线程
            self.processing_thread = ProcessingThread(input_path, output_file_path, model_type, api_key)
            
            # 连接信号
            self.processing_thread.progress_updated.connect(self._on_progress_updated)
            self.processing_thread.processing_completed.connect(self._on_processing_completed)
            self.processing_thread.error_occurred.connect(self._on_error_occurred)
            self.processing_thread.finished.connect(self._on_thread_finished)
            
            # 启动线程
            self.processing_thread.start()
            
            logger.info(f"开始处理文档: {input_path}")
            
        except Exception as e:
            logger.error(f"启动处理失败: {e}")
            self.state.is_running = False
            self.error_occurred.emit("StartupError", f"启动处理失败: {str(e)}")
    
    def stop_processing(self):
        """停止处理"""
        if not self.state.is_running or not self.processing_thread:
            return
        
        try:
            logger.info("停止文档处理")
            
            # 停止线程
            self.processing_thread.stop()
            
            # 等待线程结束（最多5秒）
            if not self.processing_thread.wait(5000):
                logger.warning("线程未能正常结束，强制终止")
                self.processing_thread.terminate()
                self.processing_thread.wait(2000)
            
            # 更新状态
            self.state.is_running = False
            self.state.current_stage = "已停止"
            
        except Exception as e:
            logger.error(f"停止处理失败: {e}")
    
    def get_processing_status(self) -> dict:
        """获取处理状态"""
        return {
            "is_running": self.state.is_running,
            "current_stage": self.state.current_stage,
            "progress": self.state.progress,
            "status_message": self.state.status_message,
            "start_time": self.state.start_time,
            "estimated_completion": self.state.estimated_completion
        }
    
    def _validate_parameters(self, input_path: str, output_file_path: str, model_type: str, api_key: str) -> list:
        """验证参数"""
        errors = []
        
        # 检查输入文件
        if not input_path:
            errors.append("输入文件路径不能为空")
        else:
            from pathlib import Path
            input_file = Path(input_path)
            if not input_file.exists():
                errors.append(f"输入文件不存在: {input_path}")
            elif not input_file.suffix.lower() in ['.docx', '.doc']:
                errors.append("输入文件必须是Word文档格式")
        
        # 检查输出文件路径
        if not output_file_path:
            errors.append("输出文件路径不能为空")
        else:
            from pathlib import Path
            output_file = Path(output_file_path)
            output_dir = output_file.parent
            
            # 确保输出目录存在
            if not output_dir.exists():
                try:
                    output_dir.mkdir(parents=True, exist_ok=True)
                except Exception as e:
                    errors.append(f"无法创建输出目录: {str(e)}")
            
            # 检查文件扩展名
            if not output_file.suffix.lower() in ['.docx', '.doc']:
                errors.append("输出文件必须是Word文档格式")
        
        # 检查模型类型
        if model_type not in ["claude", "gpt"]:
            errors.append(f"不支持的模型类型: {model_type}")
        
        # 检查API密钥
        if not api_key or not api_key.strip():
            errors.append("API密钥不能为空")
        elif len(api_key.strip()) < 10:
            errors.append("API密钥格式不正确")
        
        return errors
    
    def _on_progress_updated(self, stage: str, progress: float):
        """进度更新处理"""
        self.state.current_stage = stage
        self.state.progress = progress
        self.state.status_message = f"正在{stage}... ({progress:.1%})"
        
        # 转发信号
        self.progress_updated.emit(stage, progress)
    
    def _on_processing_completed(self, success: bool, message: str):
        """处理完成处理"""
        self.state.is_running = False
        self.state.current_stage = "完成" if success else "失败"
        self.state.progress = 1.0 if success else 0.0
        self.state.status_message = message
        
        # 转发信号
        self.processing_completed.emit(success, message)
        
        logger.info(f"处理完成: {message}")
    
    def _on_error_occurred(self, error_type: str, error_message: str):
        """错误发生处理"""
        self.state.is_running = False
        self.state.current_stage = "错误"
        self.state.status_message = f"错误: {error_message}"
        
        # 转发信号
        self.error_occurred.emit(error_type, error_message)
        
        logger.error(f"处理错误 [{error_type}]: {error_message}")
    
    def _on_thread_finished(self):
        """线程结束处理"""
        if self.processing_thread:
            self.processing_thread.deleteLater()
            self.processing_thread = None
        
        logger.info("处理线程已清理")
    
    def __del__(self):
        """析构函数"""
        if self.processing_thread and self.processing_thread.isRunning():
            self.stop_processing()