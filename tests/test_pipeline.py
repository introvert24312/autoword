"""
Test AutoWord Pipeline
测试管道编排器
"""

import pytest
from unittest.mock import Mock, patch, MagicMock
from datetime import datetime

from autoword.core.pipeline import (
    DocumentProcessor, PipelineConfig, PipelineStage, 
    PipelineProgress, PipelineResult, process_document_simple
)
from autoword.core.models import (
    Comment, DocumentStructure, TaskPlan, Task, TaskType, 
    ExecutionResult, TaskResult, ValidationResult
)
from autoword.core.llm_client import ModelType
from autoword.core.word_executor import ExecutionMode
from autoword.core.exceptions import DocumentError, LLMError, TaskExecutionError


class TestPipelineConfig:
    """测试管道配置"""
    
    def test_default_config(self):
        """测试默认配置"""
        config = PipelineConfig()
        
        assert config.model == ModelType.GPT5
        assert config.execution_mode == ExecutionMode.NORMAL
        assert config.create_backup is True
        assert config.enable_validation is True
        assert config.export_results is True
        assert config.output_dir == "output"
        assert config.visible_word is False
        assert config.max_retries == 3
        assert config.chunk_processing is True
    
    def test_custom_config(self):
        """测试自定义配置"""
        config = PipelineConfig(
            model=ModelType.CLAUDE37,
            execution_mode=ExecutionMode.DRY_RUN,
            create_backup=False,
            output_dir="custom_output"
        )
        
        assert config.model == ModelType.CLAUDE37
        assert config.execution_mode == ExecutionMode.DRY_RUN
        assert config.create_backup is False
        assert config.output_dir == "custom_output"


class TestPipelineProgress:
    """测试管道进度"""
    
    def test_progress_creation(self):
        """测试进度创建"""
        progress = PipelineProgress(
            stage=PipelineStage.LOADING,
            progress=0.5,
            message="加载中...",
            timestamp=datetime.now()
        )
        
        assert progress.stage == PipelineStage.LOADING
        assert progress.progress == 0.5
        assert progress.message == "加载中..."
        assert isinstance(progress.timestamp, datetime)


class TestDocumentProcessor:
    """测试文档处理器"""
    
    def setup_method(self):
        """测试前设置"""
        self.config = PipelineConfig(
            enable_validation=False,
            export_results=False
        )
        self.processor = DocumentProcessor(self.config)
    
    def teardown_method(self):
        """测试后清理"""
        self.processor.close()
    
    def test_add_progress_callback(self):
        """测试添加进度回调"""
        callback = Mock()
        
        self.processor.add_progress_callback(callback)
        
        assert callback in self.processor.progress_callbacks
    
    def test_report_progress(self):
        """测试报告进度"""
        callback = Mock()
        self.processor.add_progress_callback(callback)
        
        self.processor._report_progress(PipelineStage.LOADING, 0.5, "测试消息")
        
        callback.assert_called_once()
        progress = callback.call_args[0][0]
        assert progress.stage == PipelineStage.LOADING
        assert progress.progress == 0.5
        assert progress.message == "测试消息"
    
    @patch.object(DocumentProcessor, '_load_document')
    @patch.object(DocumentProcessor, '_inspect_document')
    @patch.object(DocumentProcessor, '_plan_tasks')
    @patch.object(DocumentProcessor, '_execute_tasks')
    def test_process_document_success(self, mock_execute, mock_plan, mock_inspect, mock_load):
        """测试成功处理文档"""
        # 设置模拟
        mock_document = Mock()
        mock_word_session = Mock()
        mock_load.return_value = (mock_document, mock_word_session)
        
        mock_comments = [Mock()]
        mock_structure = Mock()
        mock_inspect.return_value = (mock_comments, mock_structure)
        
        mock_planning_result = Mock()
        mock_planning_result.success = True
        mock_task_plan = Mock()
        mock_task_plan.tasks = [Mock()]
        mock_planning_result.task_plan = mock_task_plan
        mock_plan.return_value = mock_planning_result
        
        mock_execution_result = Mock()
        mock_execution_result.completed_tasks = 1
        mock_execution_result.total_tasks = 1
        mock_execute.return_value = mock_execution_result
        
        # 执行测试
        result = self.processor.process_document("test.docx")
        
        # 验证结果
        assert result.success is True
        assert result.document_path == "test.docx"
        assert PipelineStage.LOADING in result.stages_completed
        assert PipelineStage.INSPECTION in result.stages_completed
        assert PipelineStage.PLANNING in result.stages_completed
        assert PipelineStage.EXECUTION in result.stages_completed
        
        # 验证调用
        mock_load.assert_called_once_with("test.docx")
        mock_inspect.assert_called_once_with(mock_document)
        mock_plan.assert_called_once()
        mock_execute.assert_called_once()
        
        # 验证文档关闭
        mock_document.Close.assert_called_once()
    
    @patch.object(DocumentProcessor, '_load_document')
    def test_process_document_load_failure(self, mock_load):
        """测试文档加载失败"""
        mock_load.side_effect = DocumentError("文档加载失败")
        
        result = self.processor.process_document("test.docx")
        
        assert result.success is False
        assert "文档加载失败" in result.error_message
        assert len(result.stages_completed) == 0
    
    @patch.object(DocumentProcessor, '_load_document')
    @patch.object(DocumentProcessor, '_inspect_document')
    @patch.object(DocumentProcessor, '_plan_tasks')
    def test_process_document_planning_failure(self, mock_plan, mock_inspect, mock_load):
        """测试任务规划失败"""
        # 设置模拟
        mock_document = Mock()
        mock_word_session = Mock()
        mock_load.return_value = (mock_document, mock_word_session)
        
        mock_comments = [Mock()]
        mock_structure = Mock()
        mock_inspect.return_value = (mock_comments, mock_structure)
        
        mock_planning_result = Mock()
        mock_planning_result.success = False
        mock_planning_result.error_message = "规划失败"
        mock_plan.return_value = mock_planning_result
        
        # 执行测试
        result = self.processor.process_document("test.docx")
        
        # 验证结果
        assert result.success is False
        assert "任务规划失败" in result.error_message
        assert PipelineStage.LOADING in result.stages_completed
        assert PipelineStage.INSPECTION in result.stages_completed
        assert PipelineStage.PLANNING not in result.stages_completed
    
    def test_dry_run_document(self):
        """测试试运行文档"""
        with patch.object(self.processor, 'process_document') as mock_process:
            mock_result = Mock()
            mock_process.return_value = mock_result
            
            result = self.processor.dry_run_document("test.docx")
            
            assert result == mock_result
            mock_process.assert_called_once_with("test.docx")
            
            # 验证配置被临时修改
            assert self.processor.config.execution_mode == ExecutionMode.NORMAL  # 已恢复
    
    @patch('autoword.core.pipeline.WordSession')
    @patch('autoword.core.pipeline.DocLoader')
    def test_load_document_success(self, mock_doc_loader_class, mock_word_session_class):
        """测试成功加载文档"""
        # 设置模拟
        mock_word_session = Mock()
        mock_word_app = Mock()
        mock_word_session.__enter__.return_value = mock_word_app
        mock_word_session_class.return_value = mock_word_session
        
        mock_doc_loader = Mock()
        mock_document = Mock()
        mock_doc_loader.load_document.return_value = mock_document
        mock_doc_loader_class.return_value = mock_doc_loader
        
        # 执行测试
        document, word_session = self.processor._load_document("test.docx")
        
        # 验证结果
        assert document == mock_document
        assert word_session == mock_word_session
        
        # 验证调用
        mock_word_session_class.assert_called_once_with(visible=False)
        mock_doc_loader.load_document.assert_called_once_with(mock_word_app, "test.docx")
    
    @patch('autoword.core.pipeline.WordSession')
    def test_load_document_failure(self, mock_word_session_class):
        """测试文档加载失败"""
        mock_word_session_class.side_effect = Exception("COM 初始化失败")
        
        with pytest.raises(DocumentError, match="文档加载失败"):
            self.processor._load_document("test.docx")
    
    def test_inspect_document_success(self):
        """测试成功检查文档"""
        mock_document = Mock()
        
        # 设置模拟
        mock_comments = [Mock()]
        mock_structure = Mock()
        self.processor.doc_inspector.extract_comments.return_value = mock_comments
        self.processor.doc_inspector.extract_structure.return_value = mock_structure
        
        # 执行测试
        comments, structure = self.processor._inspect_document(mock_document)
        
        # 验证结果
        assert comments == mock_comments
        assert structure == mock_structure
        
        # 验证调用
        self.processor.doc_inspector.extract_comments.assert_called_once_with(mock_document)
        self.processor.doc_inspector.extract_structure.assert_called_once_with(mock_document)
    
    def test_inspect_document_failure(self):
        """测试文档检查失败"""
        mock_document = Mock()
        self.processor.doc_inspector.extract_comments.side_effect = Exception("提取批注失败")
        
        with pytest.raises(DocumentError, match="文档检查失败"):
            self.processor._inspect_document(mock_document)
    
    def test_plan_tasks_success(self):
        """测试成功规划任务"""
        mock_structure = Mock()
        mock_comments = [Mock()]
        
        mock_planning_result = Mock()
        self.processor.task_planner.generate_plan.return_value = mock_planning_result
        
        result = self.processor._plan_tasks(mock_structure, mock_comments, "test.docx")
        
        assert result == mock_planning_result
        self.processor.task_planner.generate_plan.assert_called_once_with(
            mock_structure, mock_comments, "test.docx", self.config.model
        )
    
    def test_plan_tasks_failure(self):
        """测试任务规划失败"""
        mock_structure = Mock()
        mock_comments = [Mock()]
        
        self.processor.task_planner.generate_plan.side_effect = Exception("LLM 调用失败")
        
        with pytest.raises(LLMError, match="任务规划失败"):
            self.processor._plan_tasks(mock_structure, mock_comments, "test.docx")
    
    def test_execute_tasks_success(self):
        """测试成功执行任务"""
        mock_task_plan = Mock()
        mock_task_plan.tasks = [Mock()]
        mock_comments = [Mock()]
        
        mock_execution_result = Mock()
        self.processor.word_executor.execute_tasks.return_value = mock_execution_result
        
        result = self.processor._execute_tasks(mock_task_plan, "test.docx", mock_comments)
        
        assert result == mock_execution_result
        self.processor.word_executor.execute_tasks.assert_called_once_with(
            tasks=mock_task_plan.tasks,
            document_path="test.docx",
            comments=mock_comments,
            mode=self.config.execution_mode,
            create_backup=self.config.create_backup
        )
    
    def test_execute_tasks_failure(self):
        """测试任务执行失败"""
        mock_task_plan = Mock()
        mock_task_plan.tasks = [Mock()]
        mock_comments = [Mock()]
        
        self.processor.word_executor.execute_tasks.side_effect = Exception("执行失败")
        
        with pytest.raises(TaskExecutionError, match="任务执行失败"):
            self.processor._execute_tasks(mock_task_plan, "test.docx", mock_comments)


class TestConvenienceFunction:
    """测试便捷函数"""
    
    @patch('autoword.core.pipeline.DocumentProcessor')
    def test_process_document_simple(self, mock_processor_class):
        """测试简单处理文档便捷函数"""
        mock_processor = Mock()
        mock_result = Mock()
        mock_processor.process_document.return_value = mock_result
        mock_processor_class.return_value = mock_processor
        
        result = process_document_simple(
            "test.docx",
            model=ModelType.CLAUDE37,
            dry_run=True,
            output_dir="custom_output"
        )
        
        assert result == mock_result
        
        # 验证处理器创建
        mock_processor_class.assert_called_once()
        config = mock_processor_class.call_args[0][0]
        assert config.model == ModelType.CLAUDE37
        assert config.execution_mode == ExecutionMode.DRY_RUN
        assert config.output_dir == "custom_output"
        
        # 验证调用
        mock_processor.process_document.assert_called_once_with("test.docx")
        mock_processor.close.assert_called_once()


if __name__ == "__main__":
    pytest.main([__file__])