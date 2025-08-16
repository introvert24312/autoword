"""
AutoWord Enhanced Executor
增强执行器 - 主要入口点，集成所有核心功能
"""

import os
import logging
import time
from typing import Optional
from enum import Enum
from datetime import datetime

from .doc_loader import WordSession
from .doc_inspector import DocumentInspector
from .prompt_builder import PromptBuilder
from .planner import TaskPlanner
from .word_executor import WordExecutor
from .llm_client import LLMClient, ModelType
from .models import ExecutionResult, TaskResult
from .exceptions import AutoWordError, DocumentError, LLMError


logger = logging.getLogger(__name__)


class WorkflowMode(Enum):
    """工作流程模式"""
    NORMAL = "normal"      # 正常模式
    DRY_RUN = "dry_run"    # 试运行模式
    SAFE = "safe"          # 安全模式


class EnhancedExecutor:
    """增强执行器 - AutoWord的主要入口点"""
    
    def __init__(self, 
                 llm_model: ModelType = ModelType.GPT5,
                 visible: bool = False,
                 max_retries: int = 3):
        """
        初始化增强执行器
        
        Args:
            llm_model: 使用的LLM模型
            visible: 是否显示Word窗口
            max_retries: 最大重试次数
        """
        self.llm_model = llm_model
        self.visible = visible
        self.max_retries = max_retries
        
        # 初始化组件
        self.llm_client = LLMClient()
        self.doc_inspector = DocumentInspector()
        self.prompt_builder = PromptBuilder()
        self.task_planner = TaskPlanner()
        self.word_executor = WordExecutor()
    
    def execute_workflow(self, 
                        document_path: str,
                        mode: WorkflowMode = WorkflowMode.NORMAL) -> ExecutionResult:
        """
        执行完整的文档处理工作流程
        
        Args:
            document_path: 文档路径
            mode: 执行模式
            
        Returns:
            ExecutionResult: 执行结果
        """
        start_time = time.time()
        
        try:
            logger.info(f"开始处理文档: {document_path}")
            logger.info(f"执行模式: {mode.value}")
            
            # 1. 文档加载和检查
            logger.info("步骤1: 文档加载和检查")
            document_info = self._load_and_inspect_document(document_path)
            
            if not document_info.comments:
                logger.warning("文档中没有找到批注")
                return ExecutionResult(
                    success=True,
                    total_tasks=0,
                    completed_tasks=0,
                    failed_tasks=0,
                    task_results=[],
                    execution_time=time.time() - start_time,
                    error_summary="文档中没有批注，无需处理"
                )
            
            logger.info(f"找到 {len(document_info.comments)} 个批注")
            
            # 2. 任务规划
            logger.info("步骤2: 任务规划")
            tasks = self._plan_tasks(document_info)
            
            if not tasks:
                logger.warning("没有生成任何任务")
                return ExecutionResult(
                    success=True,
                    total_tasks=0,
                    completed_tasks=0,
                    failed_tasks=0,
                    task_results=[],
                    execution_time=time.time() - start_time,
                    error_summary="没有生成任何任务"
                )
            
            logger.info(f"生成了 {len(tasks)} 个任务")
            
            # 3. 任务执行
            logger.info("步骤3: 任务执行")
            execution_result = self._execute_tasks(document_path, tasks, mode)
            
            execution_result.execution_time = time.time() - start_time
            logger.info(f"工作流程完成，耗时 {execution_result.execution_time:.2f}s")
            
            return execution_result
            
        except Exception as e:
            logger.error(f"工作流程执行失败: {str(e)}")
            return ExecutionResult(
                success=False,
                total_tasks=0,
                completed_tasks=0,
                failed_tasks=0,
                task_results=[],
                execution_time=time.time() - start_time,
                error_summary=f"工作流程执行失败: {str(e)}"
            )
    
    def _load_and_inspect_document(self, document_path: str):
        """加载和检查文档"""
        try:
            with WordSession(visible=self.visible) as word_app:
                document = word_app.Documents.Open(document_path)
                try:
                    return self.doc_inspector.get_document_info(document)
                finally:
                    document.Close(SaveChanges=False)
        except Exception as e:
            raise DocumentError(f"文档加载失败: {str(e)}")
    
    def _plan_tasks(self, document_info):
        """规划任务"""
        try:
            # 构建提示词
            context = self.prompt_builder.build_context_from_document(document_info)
            user_prompt = self.prompt_builder.build_user_prompt(context)
            
            # 生成任务计划
            planning_result = self.task_planner.generate_plan(
                document_structure=context.document_structure,
                comments=document_info.comments,
                model=self.llm_model
            )
            
            return planning_result.tasks
            
        except Exception as e:
            raise LLMError(f"任务规划失败: {str(e)}")
    
    def _execute_tasks(self, document_path: str, tasks, mode: WorkflowMode):
        """执行任务"""
        try:
            if mode == WorkflowMode.DRY_RUN:
                return self.word_executor.dry_run_tasks(tasks, document_path)
            elif mode == WorkflowMode.SAFE:
                return self.word_executor.execute_tasks(
                    tasks, document_path, 
                    create_backup=True
                )
            else:  # NORMAL
                return self.word_executor.execute_tasks(
                    tasks, document_path,
                    create_backup=False
                )
        except Exception as e:
            raise AutoWordError(f"任务执行失败: {str(e)}")