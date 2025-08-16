"""
AutoWord Enhanced Executor
增强执行器 - 集成完整的四重格式保护防线
"""

import logging
import time
from typing import List, Optional, Tuple
from datetime import datetime

from .word_executor import WordExecutor, ExecutionMode
from .format_validator import FormatValidator, ValidationReport
from .doc_inspector import DocInspector
from .doc_loader import DocLoader
from .models import (
    Task, Comment, ExecutionResult, TaskResult, DocumentSnapshot
)
from .exceptions import FormatProtectionError


logger = logging.getLogger(__name__)


class EnhancedExecutionResult(ExecutionResult):
    """增强的执行结果，包含格式验证信息"""
    
    def __init__(self, *args, **kwargs):
        # 提取新增字段
        self.before_snapshot = kwargs.pop('before_snapshot', None)
        self.after_snapshot = kwargs.pop('after_snapshot', None)
        self.validation_report = kwargs.pop('validation_report', None)
        self.rollback_performed = kwargs.pop('rollback_performed', False)
        self.backup_path = kwargs.pop('backup_path', None)
        
        # 调用父类构造函数
        super().__init__(*args, **kwargs)


class EnhancedWordExecutor:
    """增强的 Word 执行器 - 实现完整的四重格式保护防线"""
    
    def __init__(self, visible: bool = False, enable_validation: bool = True):
        """
        初始化增强执行器
        
        Args:
            visible: 是否显示 Word 窗口
            enable_validation: 是否启用格式验证（第4层防线）
        """
        self.visible = visible
        self.enable_validation = enable_validation
        
        # 初始化组件
        self.base_executor = WordExecutor(visible=visible)
        self.doc_loader = DocLoader(visible=visible)
        self.doc_inspector = DocInspector()
        self.format_validator = FormatValidator()
    
    def execute_tasks_with_protection(self,
                                    tasks: List[Task],
                                    document_path: str,
                                    comments: List[Comment] = None,
                                    mode: ExecutionMode = ExecutionMode.NORMAL,
                                    auto_rollback: bool = True) -> EnhancedExecutionResult:
        """
        执行任务并提供完整的格式保护
        
        Args:
            tasks: 任务列表
            document_path: 文档路径
            comments: 批注列表
            mode: 执行模式
            auto_rollback: 是否自动回滚未授权变更
            
        Returns:
            增强的执行结果
        """
        start_time = time.time()
        backup_path = None
        before_snapshot = None
        after_snapshot = None
        validation_report = None
        rollback_performed = False
        
        try:
            logger.info(f"开始增强执行: {len(tasks)} 个任务")
            
            # 1. 创建备份
            if mode != ExecutionMode.DRY_RUN:
                backup_path = self.doc_loader.create_backup(document_path)
                logger.info(f"备份已创建: {backup_path}")
            
            # 2. 创建执行前快照（如果启用验证）
            if self.enable_validation and mode != ExecutionMode.DRY_RUN:
                before_snapshot = self._create_document_snapshot(document_path, comments)
                logger.info("执行前快照已创建")
            
            # 3. 执行任务（使用基础执行器）
            base_result = self.base_executor.execute_tasks(
                tasks=tasks,
                document_path=document_path,
                comments=comments,
                mode=mode,
                create_backup=False  # 我们已经创建了备份
            )
            
            # 4. 创建执行后快照并验证（如果启用验证且不是试运行）
            if self.enable_validation and mode != ExecutionMode.DRY_RUN:
                after_snapshot = self._create_document_snapshot(document_path, comments)
                logger.info("执行后快照已创建")
                
                # 执行格式验证 - 第4层防线
                validation_report = self.format_validator.validate_execution_result(
                    before_snapshot=before_snapshot,
                    after_snapshot=after_snapshot,
                    executed_tasks=[t for t in tasks if self._task_executed_successfully(t, base_result.task_results)],
                    comments=comments or []
                )
                
                logger.info(f"格式验证完成: {'通过' if validation_report.is_valid else '失败'}")
                
                # 5. 处理验证失败的情况
                if not validation_report.is_valid and auto_rollback:
                    rollback_performed = self._perform_rollback(document_path, backup_path, validation_report)
            
            execution_time = time.time() - start_time
            
            # 6. 创建增强的执行结果
            enhanced_result = EnhancedExecutionResult(
                success=base_result.success and (validation_report.is_valid if validation_report else True),
                total_tasks=base_result.total_tasks,
                completed_tasks=base_result.completed_tasks,
                failed_tasks=base_result.failed_tasks,
                task_results=base_result.task_results,
                execution_time=execution_time,
                error_summary=self._create_enhanced_error_summary(base_result, validation_report, rollback_performed),
                before_snapshot=before_snapshot,
                after_snapshot=after_snapshot,
                validation_report=validation_report,
                rollback_performed=rollback_performed,
                backup_path=backup_path
            )
            
            logger.info(f"增强执行完成: 成功={enhanced_result.success}, 回滚={rollback_performed}")
            return enhanced_result
            
        except Exception as e:
            execution_time = time.time() - start_time
            logger.error(f"增强执行失败: {e}")
            
            # 发生异常时也尝试回滚
            if backup_path and auto_rollback:
                try:
                    self._restore_from_backup(document_path, backup_path)
                    rollback_performed = True
                    logger.info("异常情况下回滚成功")
                except Exception as rollback_error:
                    logger.error(f"异常情况下回滚失败: {rollback_error}")
            
            return EnhancedExecutionResult(
                success=False,
                total_tasks=len(tasks),
                completed_tasks=0,
                failed_tasks=len(tasks),
                task_results=[],
                execution_time=execution_time,
                error_summary=f"增强执行失败: {e}",
                rollback_performed=rollback_performed,
                backup_path=backup_path
            )
    
    def dry_run_with_validation(self,
                              tasks: List[Task],
                              document_path: str,
                              comments: List[Comment] = None) -> EnhancedExecutionResult:
        """
        试运行并进行预验证
        
        Args:
            tasks: 任务列表
            document_path: 文档路径
            comments: 批注列表
            
        Returns:
            增强的执行结果
        """
        logger.info("开始试运行预验证")
        
        # 执行试运行
        result = self.execute_tasks_with_protection(
            tasks=tasks,
            document_path=document_path,
            comments=comments,
            mode=ExecutionMode.DRY_RUN,
            auto_rollback=False
        )
        
        # 添加预验证信息
        if result.task_results:
            dry_run_warnings = []
            for task_result in result.task_results:
                if not task_result.success:
                    dry_run_warnings.append(f"任务 {task_result.task_id} 可能失败: {task_result.message}")
            
            if dry_run_warnings:
                result.error_summary = f"试运行发现潜在问题: {'; '.join(dry_run_warnings)}"
        
        return result
    
    def _create_document_snapshot(self, document_path: str, comments: List[Comment]) -> DocumentSnapshot:
        """创建文档快照"""
        try:
            with self.doc_loader.open_document(document_path, create_backup=False) as (word_app, document):
                structure = self.doc_inspector.extract_structure(document)
                snapshot = DocumentSnapshot(
                    timestamp=datetime.now(),
                    document_path=document_path,
                    structure=structure,
                    comments=comments or [],
                    checksum=""  # 简化实现，不计算校验和
                )
                return snapshot
        except Exception as e:
            logger.error(f"创建文档快照失败: {e}")
            raise
    
    def _task_executed_successfully(self, task: Task, task_results: List[TaskResult]) -> bool:
        """检查任务是否执行成功"""
        for result in task_results:
            if result.task_id == task.id:
                return result.success
        return False
    
    def _perform_rollback(self, document_path: str, backup_path: str, validation_report: ValidationReport) -> bool:
        """执行回滚操作"""
        try:
            logger.warning(f"检测到 {validation_report.unauthorized_count} 个未授权变更，开始回滚")
            
            # 记录未授权变更的详细信息
            for change in validation_report.unauthorized_changes:
                logger.warning(f"未授权变更: {change.change_type} - {change.old_value} -> {change.new_value}")
            
            # 从备份恢复
            self._restore_from_backup(document_path, backup_path)
            
            logger.info("回滚操作完成")
            return True
            
        except Exception as e:
            logger.error(f"回滚操作失败: {e}")
            return False
    
    def _restore_from_backup(self, document_path: str, backup_path: str):
        """从备份恢复文档"""
        import shutil
        
        try:
            shutil.copy2(backup_path, document_path)
            logger.info(f"文档已从备份恢复: {backup_path} -> {document_path}")
        except Exception as e:
            raise Exception(f"从备份恢复失败: {e}")
    
    def _create_enhanced_error_summary(self, 
                                     base_result: ExecutionResult,
                                     validation_report: Optional[ValidationReport],
                                     rollback_performed: bool) -> Optional[str]:
        """创建增强的错误摘要"""
        error_parts = []
        
        # 基础执行错误
        if base_result.error_summary:
            error_parts.append(f"执行错误: {base_result.error_summary}")
        
        # 格式验证错误
        if validation_report and not validation_report.is_valid:
            error_parts.append(f"格式验证失败: {validation_report.unauthorized_count} 个未授权变更")
        
        # 回滚信息
        if rollback_performed:
            error_parts.append("已执行自动回滚")
        
        return "; ".join(error_parts) if error_parts else None
    
    def get_protection_status(self) -> dict:
        """获取格式保护状态"""
        return {
            "four_layer_protection": {
                "layer_1": "提示词硬约束 - 已启用",
                "layer_2": "规划期过滤 - 已启用",
                "layer_3": "执行期拦截 - 已启用",
                "layer_4": f"事后校验回滚 - {'已启用' if self.enable_validation else '已禁用'}"
            },
            "validation_enabled": self.enable_validation,
            "auto_backup": True,
            "rollback_capability": True
        }


# 便捷函数
def execute_with_full_protection(tasks: List[Task],
                                document_path: str,
                                comments: List[Comment] = None,
                                visible: bool = False,
                                dry_run: bool = False,
                                auto_rollback: bool = True) -> EnhancedExecutionResult:
    """
    便捷函数：使用完整格式保护执行任务
    
    Args:
        tasks: 任务列表
        document_path: 文档路径
        comments: 批注列表
        visible: 是否显示 Word 窗口
        dry_run: 是否试运行
        auto_rollback: 是否自动回滚
        
    Returns:
        增强的执行结果
    """
    executor = EnhancedWordExecutor(visible=visible, enable_validation=True)
    
    if dry_run:
        return executor.dry_run_with_validation(tasks, document_path, comments)
    else:
        return executor.execute_tasks_with_protection(
            tasks=tasks,
            document_path=document_path,
            comments=comments,
            auto_rollback=auto_rollback
        )