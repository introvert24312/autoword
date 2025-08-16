"""
AutoWord Pipeline Orchestrator
核心管道编排器，负责协调整个文档处理流程
"""

import logging
import time
from typing import List, Dict, Any, Optional, Callable
from datetime import datetime
from dataclasses import dataclass
from enum import Enum

from .models import (
    Comment, DocumentStructure, TaskPlan, ExecutionResult, 
    TaskResult, ValidationResult
)
from .doc_loader import WordSession, DocLoader
from .doc_inspector import DocInspector
from .prompt_builder import PromptBuilder, PromptContext
from .planner import TaskPlanner, PlanningResult
from .word_executor import WordExecutor, ExecutionMode
from .format_validator import FormatValidator
from .exporter import Exporter, export_execution_report
from .llm_client import ModelType
from .exceptions import (
    DocumentError, LLMError, TaskExecutionError, 
    FormatProtectionError, ValidationError
)


logger = logging.getLogger(__name__)


class PipelineStage(str, Enum):
    """管道阶段枚举"""
    LOADING = "loading"           # 文档加载
    INSPECTION = "inspection"     # 文档检查
    PLANNING = "planning"         # 任务规划
    EXECUTION = "execution"       # 任务执行
    VALIDATION = "validation"     # 结果验证
    EXPORT = "export"            # 结果导出


@dataclass
class PipelineConfig:
    """管道配置"""
    model: ModelType = ModelType.GPT5
    execution_mode: ExecutionMode = ExecutionMode.NORMAL
    create_backup: bool = True
    enable_validation: bool = True
    export_results: bool = True
    output_dir: str = "output"
    visible_word: bool = False
    max_retries: int = 3
    chunk_processing: bool = True


@dataclass
class PipelineProgress:
    """管道进度"""
    stage: PipelineStage
    progress: float  # 0.0 - 1.0
    message: str
    timestamp: datetime
    
    def __post_init__(self):
        if not hasattr(self, 'timestamp') or self.timestamp is None:
            self.timestamp = datetime.now()


@dataclass
class PipelineResult:
    """管道执行结果"""
    success: bool
    document_path: str
    total_time: float
    stages_completed: List[PipelineStage]
    task_plan: Optional[TaskPlan] = None
    execution_result: Optional[ExecutionResult] = None
    validation_result: Optional[ValidationResult] = None
    exported_files: Optional[Dict[str, str]] = None
    error_message: Optional[str] = None
    progress_history: List[PipelineProgress] = None
    
    def __post_init__(self):
        if self.progress_history is None:
            self.progress_history = []


class DocumentProcessor:
    """文档处理器 - 核心管道编排器"""
    
    def __init__(self, config: Optional[PipelineConfig] = None, api_keys: Optional[Dict[str, str]] = None):
        """
        初始化文档处理器
        
        Args:
            config: 管道配置
            api_keys: API密钥字典
        """
        self.config = config or PipelineConfig()
        self.progress_callbacks: List[Callable[[PipelineProgress], None]] = []
        
        # 初始化组件
        self.doc_loader = DocLoader()
        self.doc_inspector = DocInspector()
        self.prompt_builder = PromptBuilder()
        self.task_planner = TaskPlanner(default_model=self.config.model, api_keys=api_keys)
        self.word_executor = WordExecutor(visible=self.config.visible_word)
        self.format_validator = FormatValidator() if self.config.enable_validation else None
        self.exporter = Exporter(self.config.output_dir) if self.config.export_results else None
    
    def add_progress_callback(self, callback: Callable[[PipelineProgress], None]):
        """添加进度回调函数"""
        self.progress_callbacks.append(callback)
    
    def _report_progress(self, stage: PipelineStage, progress: float, message: str):
        """报告进度"""
        progress_info = PipelineProgress(
            stage=stage,
            progress=progress,
            message=message,
            timestamp=datetime.now()
        )
        
        logger.info(f"[{stage.value.upper()}] {progress:.1%} - {message}")
        
        for callback in self.progress_callbacks:
            try:
                callback(progress_info)
            except Exception as e:
                logger.error(f"进度回调失败: {e}")
    
    def process_document(self, document_path: str, output_file_path: Optional[str] = None) -> PipelineResult:
        """
        处理文档的完整管道
        
        Args:
            document_path: 文档路径
            
        Returns:
            管道执行结果
        """
        start_time = time.time()
        stages_completed = []
        progress_history = []
        
        try:
            logger.info(f"开始处理文档: {document_path}")
            
            # 阶段1: 文档加载
            self._report_progress(PipelineStage.LOADING, 0.0, "开始加载文档")
            document, word_app = self._load_document(document_path)
            stages_completed.append(PipelineStage.LOADING)
            self._report_progress(PipelineStage.LOADING, 1.0, "文档加载完成")
            
            try:
                # 阶段2: 文档检查
                self._report_progress(PipelineStage.INSPECTION, 0.0, "开始检查文档")
                comments, structure = self._inspect_document(document)
                stages_completed.append(PipelineStage.INSPECTION)
                self._report_progress(PipelineStage.INSPECTION, 1.0, f"文档检查完成: {len(comments)} 个批注")
                
                # 阶段3: 任务规划
                self._report_progress(PipelineStage.PLANNING, 0.0, "开始生成任务计划")
                planning_result = self._plan_tasks(structure, comments, document_path)
                if not planning_result.success:
                    raise TaskExecutionError(f"任务规划失败: {planning_result.error_message}")
                
                task_plan = planning_result.task_plan
                stages_completed.append(PipelineStage.PLANNING)
                self._report_progress(PipelineStage.PLANNING, 1.0, f"任务规划完成: {len(task_plan.tasks)} 个任务")
                
                # 阶段4: 任务执行
                self._report_progress(PipelineStage.EXECUTION, 0.0, "开始执行任务")
                execution_result = self._execute_tasks(task_plan, document_path, comments, output_file_path)
                stages_completed.append(PipelineStage.EXECUTION)
                self._report_progress(PipelineStage.EXECUTION, 1.0, 
                                    f"任务执行完成: {execution_result.completed_tasks}/{execution_result.total_tasks}")
                
                # 阶段5: 结果验证
                validation_result = None
                if self.config.enable_validation:
                    self._report_progress(PipelineStage.VALIDATION, 0.0, "开始验证结果")
                    validation_result = self._validate_results(task_plan, execution_result, comments)
                    stages_completed.append(PipelineStage.VALIDATION)
                    self._report_progress(PipelineStage.VALIDATION, 1.0, "结果验证完成")
                
                # 阶段6: 结果导出
                exported_files = None
                if self.config.export_results:
                    self._report_progress(PipelineStage.EXPORT, 0.0, "开始导出结果")
                    exported_files = self._export_results(task_plan, execution_result, structure)
                    stages_completed.append(PipelineStage.EXPORT)
                    self._report_progress(PipelineStage.EXPORT, 1.0, "结果导出完成")
                
                total_time = time.time() - start_time
                
                logger.info(f"文档处理完成: {document_path} (耗时: {total_time:.2f}s)")
                
                return PipelineResult(
                    success=True,
                    document_path=document_path,
                    total_time=total_time,
                    stages_completed=stages_completed,
                    task_plan=task_plan,
                    execution_result=execution_result,
                    validation_result=validation_result,
                    exported_files=exported_files,
                    progress_history=progress_history
                )
                
            finally:
                # 确保关闭文档和Word应用程序
                try:
                    if document:
                        document.Close()
                    if word_app:
                        word_app.Quit()
                except:
                    pass
                
        except Exception as e:
            total_time = time.time() - start_time
            error_msg = f"文档处理失败: {e}"
            
            logger.error(error_msg)
            
            return PipelineResult(
                success=False,
                document_path=document_path,
                total_time=total_time,
                stages_completed=stages_completed,
                error_message=error_msg,
                progress_history=progress_history
            )
    
    def _load_document(self, document_path: str):
        """加载文档"""
        try:
            word_app, document = self.doc_loader.load_document(document_path, create_backup=self.config.create_backup)
            return document, word_app
        except Exception as e:
            raise DocumentError(f"文档加载失败: {e}")
    
    def _inspect_document(self, document):
        """检查文档"""
        try:
            comments = self.doc_inspector.extract_comments(document)
            structure = self.doc_inspector.extract_structure(document)
            return comments, structure
        except Exception as e:
            raise DocumentError(f"文档检查失败: {e}")
    
    def _plan_tasks(self, structure: DocumentStructure, comments: List[Comment], document_path: str) -> PlanningResult:
        """规划任务"""
        try:
            return self.task_planner.generate_plan(structure, comments, document_path, self.config.model)
        except Exception as e:
            raise LLMError(f"任务规划失败: {e}")
    
    def _execute_tasks(self, task_plan: TaskPlan, document_path: str, comments: List[Comment], output_file_path: Optional[str] = None) -> ExecutionResult:
        """执行任务"""
        try:
            return self.word_executor.execute_tasks(
                tasks=task_plan.tasks,
                document_path=document_path,
                comments=comments,
                mode=self.config.execution_mode,
                create_backup=self.config.create_backup,
                output_file_path=output_file_path
            )
        except Exception as e:
            raise TaskExecutionError(f"任务执行失败: {e}")
    
    def _validate_results(self, task_plan: TaskPlan, execution_result: ExecutionResult, comments: List[Comment]) -> ValidationResult:
        """验证结果"""
        if not self.format_validator:
            return ValidationResult(is_valid=True, errors=[], warnings=[])
        
        try:
            return self.format_validator.validate_execution_result(task_plan, execution_result, comments)
        except Exception as e:
            logger.error(f"结果验证失败: {e}")
            return ValidationResult(
                is_valid=False,
                errors=[f"验证失败: {e}"],
                warnings=[]
            )
    
    def _export_results(self, task_plan: TaskPlan, execution_result: ExecutionResult, structure: DocumentStructure) -> Dict[str, str]:
        """导出结果"""
        if not self.exporter:
            return {}
        
        try:
            exported_files = {}
            
            # 导出任务计划
            exported_files["plan"] = self.exporter.export_plan(task_plan)
            
            # 导出执行日志
            exported_files["log"] = self.exporter.export_execution_log(execution_result)
            
            return exported_files
            
        except Exception as e:
            logger.error(f"结果导出失败: {e}")
            return {}
    
    def dry_run_document(self, document_path: str) -> PipelineResult:
        """
        试运行文档处理
        
        Args:
            document_path: 文档路径
            
        Returns:
            管道执行结果
        """
        # 临时修改配置为试运行模式
        original_mode = self.config.execution_mode
        original_backup = self.config.create_backup
        
        self.config.execution_mode = ExecutionMode.DRY_RUN
        self.config.create_backup = False
        
        try:
            return self.process_document(document_path)
        finally:
            # 恢复原始配置
            self.config.execution_mode = original_mode
            self.config.create_backup = original_backup
    
    def close(self):
        """关闭处理器资源"""
        if hasattr(self, 'task_planner') and hasattr(self.task_planner, 'close'):
            try:
                self.task_planner.close()
            except:
                pass


# 便捷函数
def process_document_simple(document_path: str, 
                           model: ModelType = ModelType.GPT5,
                           dry_run: bool = False,
                           output_dir: str = "output") -> PipelineResult:
    """
    便捷函数：简单处理文档
    
    Args:
        document_path: 文档路径
        model: LLM 模型
        dry_run: 是否试运行
        output_dir: 输出目录
        
    Returns:
        管道执行结果
    """
    config = PipelineConfig(
        model=model,
        execution_mode=ExecutionMode.DRY_RUN if dry_run else ExecutionMode.NORMAL,
        output_dir=output_dir
    )
    
    processor = DocumentProcessor(config)
    
    try:
        return processor.process_document(document_path)
    finally:
        processor.close()