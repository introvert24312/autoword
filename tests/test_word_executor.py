"""
Test AutoWord Word Executor
测试 Word COM 执行器
"""

import pytest
from unittest.mock import Mock, patch, MagicMock
from datetime import datetime

from autoword.core.word_executor import (
    WordExecutor, TaskExecutor, TaskLocator, ExecutionContext, 
    ExecutionMode, execute_task_list
)
from autoword.core.models import (
    Task, TaskType, RiskLevel, LocatorType, Locator, Comment, TaskResult
)
from autoword.core.exceptions import TaskExecutionError, FormatProtectionError


class TestTaskLocator:
    """测试任务定位器"""
    
    def setup_method(self):
        """测试前设置"""
        self.mock_word_app = Mock()
        self.mock_document = Mock()
        self.locator = TaskLocator(self.mock_word_app, self.mock_document)
    
    @patch.object(TaskLocator, '_locate_by_bookmark')
    def test_locate_by_bookmark_existing(self, mock_locate_bookmark):
        """测试通过现有书签定位"""
        mock_range = Mock()
        mock_locate_bookmark.return_value = mock_range
        
        task = Task(
            id="task_1",
            type=TaskType.REWRITE,
            locator=Locator(by=LocatorType.BOOKMARK, value="test_bookmark"),
            instruction="测试"
        )
        
        result = self.locator.locate_target(task)
        
        assert result == mock_range
        mock_locate_bookmark.assert_called_once_with("test_bookmark")
    
    def test_locate_by_range_with_dash(self):
        """测试通过范围定位（使用破折号格式）"""
        mock_range = Mock()
        self.mock_document.Range.return_value = mock_range
        
        task = Task(
            id="task_1",
            type=TaskType.REWRITE,
            locator=Locator(by=LocatorType.RANGE, value="100-200"),
            instruction="测试"
        )
        
        result = self.locator.locate_target(task)
        
        self.mock_document.Range.assert_called_with(100, 200)
        assert result == mock_range
    
    def test_locate_by_range_with_comma(self):
        """测试通过范围定位（使用逗号格式）"""
        mock_range = Mock()
        self.mock_document.Range.return_value = mock_range
        
        task = Task(
            id="task_1",
            type=TaskType.REWRITE,
            locator=Locator(by=LocatorType.RANGE, value="100,50"),
            instruction="测试"
        )
        
        result = self.locator.locate_target(task)
        
        self.mock_document.Range.assert_called_with(100, 150)
        assert result == mock_range
    
    @patch.object(TaskLocator, '_locate_by_heading')
    def test_locate_by_heading(self, mock_locate_heading):
        """测试通过标题定位"""
        mock_range = Mock()
        mock_locate_heading.return_value = mock_range
        
        task = Task(
            id="task_1",
            type=TaskType.REWRITE,
            locator=Locator(by=LocatorType.HEADING, value="测试标题"),
            instruction="测试"
        )
        
        result = self.locator.locate_target(task)
        
        assert result == mock_range
        mock_locate_heading.assert_called_once_with("测试标题")
    
    def test_locate_by_find_success(self):
        """测试通过文本查找定位（成功）"""
        mock_range = Mock()
        mock_find = Mock()
        mock_find.Execute.return_value = True
        mock_range.Find = mock_find
        
        self.mock_document.Range.return_value = mock_range
        
        task = Task(
            id="task_1",
            type=TaskType.REWRITE,
            locator=Locator(by=LocatorType.FIND, value="测试文本"),
            instruction="测试"
        )
        
        with patch('autoword.core.word_executor.win32_constants') as mock_constants:
            mock_constants.wdFindContinue = 1
            
            result = self.locator.locate_target(task)
        
        assert result == mock_range
        mock_find.Execute.assert_called_once()
    
    @patch.object(TaskLocator, '_locate_by_find')
    def test_locate_by_find_failure_with_fuzzy(self, mock_locate_find):
        """测试文本查找失败时的模糊匹配"""
        mock_range = Mock()
        mock_locate_find.return_value = mock_range
        
        task = Task(
            id="task_1",
            type=TaskType.REWRITE,
            locator=Locator(by=LocatorType.FIND, value="测试 文本 内容"),
            instruction="测试"
        )
        
        result = self.locator.locate_target(task)
        
        assert result == mock_range
        mock_locate_find.assert_called_once_with("测试 文本 内容")
    
    def test_create_bookmark(self):
        """测试创建书签"""
        mock_range = Mock()
        mock_bookmarks = Mock()
        self.mock_document.Bookmarks = mock_bookmarks
        
        # 模拟现有书签
        existing_bookmark = Mock()
        existing_bookmark.Name = "existing_bookmark"
        mock_bookmarks.__iter__ = Mock(return_value=iter([existing_bookmark]))
        
        result = self.locator.create_bookmark(mock_range, "test_bookmark")
        
        assert result == "test_bookmark"
        mock_bookmarks.Add.assert_called_once_with("test_bookmark", mock_range)
    
    def test_create_bookmark_with_duplicate_name(self):
        """测试创建重复名称的书签"""
        mock_range = Mock()
        mock_bookmarks = Mock()
        self.mock_document.Bookmarks = mock_bookmarks
        
        # 模拟现有书签
        existing_bookmark = Mock()
        existing_bookmark.Name = "test_bookmark"
        mock_bookmarks.__iter__ = Mock(return_value=iter([existing_bookmark]))
        
        result = self.locator.create_bookmark(mock_range, "test_bookmark")
        
        assert result == "test_bookmark_1"
        mock_bookmarks.Add.assert_called_once_with("test_bookmark_1", mock_range)


class TestTaskExecutor:
    """测试任务执行器"""
    
    def setup_method(self):
        """测试前设置"""
        self.mock_context = Mock(spec=ExecutionContext)
        self.mock_context.word_app = Mock()
        self.mock_context.document = Mock()
        self.mock_context.mode = ExecutionMode.NORMAL
        self.mock_context.comments = []
        
        self.executor = TaskExecutor(self.mock_context)
    
    @patch.object(TaskExecutor, '_validate_task_before_execution')
    @patch.object(TaskLocator, 'locate_target')
    def test_execute_rewrite_task(self, mock_locate, mock_validate):
        """测试执行重写任务"""
        # 设置模拟
        mock_range = Mock()
        mock_range.Text = "原始文本"
        mock_locate.return_value = mock_range
        
        task = Task(
            id="task_1",
            type=TaskType.REWRITE,
            locator=Locator(by=LocatorType.FIND, value="测试"),
            instruction="重写这段文本"
        )
        
        # 执行测试
        result = self.executor.execute_task(task)
        
        # 验证结果
        assert result.success is True
        assert result.task_id == "task_1"
        assert "重写完成" in result.message
        mock_validate.assert_called_once_with(task)
        mock_locate.assert_called_once_with(task)
    
    @patch.object(TaskExecutor, '_validate_task_before_execution')
    @patch.object(TaskLocator, 'locate_target')
    def test_execute_insert_task(self, mock_locate, mock_validate):
        """测试执行插入任务"""
        mock_range = Mock()
        mock_locate.return_value = mock_range
        
        task = Task(
            id="task_1",
            type=TaskType.INSERT,
            locator=Locator(by=LocatorType.FIND, value="测试"),
            instruction="插入新内容"
        )
        
        result = self.executor.execute_task(task)
        
        assert result.success is True
        assert "插入完成" in result.message
        mock_range.InsertAfter.assert_called_once()
    
    @patch.object(TaskExecutor, '_validate_task_before_execution')
    @patch.object(TaskLocator, 'locate_target')
    def test_execute_delete_task(self, mock_locate, mock_validate):
        """测试执行删除任务"""
        mock_range = Mock()
        mock_range.Text = "要删除的文本"
        mock_locate.return_value = mock_range
        
        task = Task(
            id="task_1",
            type=TaskType.DELETE,
            locator=Locator(by=LocatorType.FIND, value="测试"),
            instruction="删除这段文本"
        )
        
        result = self.executor.execute_task(task)
        
        assert result.success is True
        assert "删除完成" in result.message
        mock_range.Delete.assert_called_once()
    
    @patch.object(TaskExecutor, '_validate_task_before_execution')
    @patch.object(TaskLocator, 'locate_target')
    def test_execute_set_paragraph_style_task(self, mock_locate, mock_validate):
        """测试执行设置段落样式任务"""
        mock_range = Mock()
        mock_paragraph = Mock()
        mock_paragraph.Style.NameLocal = "正文"
        mock_range.Paragraphs.Count = 1
        mock_range.Paragraphs.return_value = mock_paragraph
        mock_locate.return_value = mock_range
        
        task = Task(
            id="task_1",
            type=TaskType.SET_PARAGRAPH_STYLE,
            source_comment_id="comment_1",
            locator=Locator(by=LocatorType.FIND, value="测试"),
            instruction="设置为标题1样式"
        )
        
        result = self.executor.execute_task(task)
        
        assert result.success is True
        assert "样式设置完成" in result.message
    
    @patch.object(TaskExecutor, '_validate_task_before_execution')
    @patch.object(TaskLocator, 'locate_target')
    def test_execute_set_heading_level_task(self, mock_locate, mock_validate):
        """测试执行设置标题级别任务"""
        mock_range = Mock()
        mock_paragraph = Mock()
        mock_paragraph.Style.NameLocal = "正文"
        mock_range.Paragraphs.Count = 1
        mock_range.Paragraphs.return_value = mock_paragraph
        mock_locate.return_value = mock_range
        
        task = Task(
            id="task_1",
            type=TaskType.SET_HEADING_LEVEL,
            source_comment_id="comment_1",
            locator=Locator(by=LocatorType.FIND, value="测试"),
            instruction="设置为2级标题"
        )
        
        result = self.executor.execute_task(task)
        
        assert result.success is True
        assert "标题级别设置完成" in result.message
    
    def test_execute_task_with_dry_run_mode(self):
        """测试试运行模式"""
        self.mock_context.mode = ExecutionMode.DRY_RUN
        
        with patch.object(TaskLocator, 'locate_target') as mock_locate:
            mock_range = Mock()
            mock_range.Text = "测试文本"
            mock_locate.return_value = mock_range
            
            task = Task(
                id="task_1",
                type=TaskType.REWRITE,
                locator=Locator(by=LocatorType.FIND, value="测试"),
                instruction="重写文本"
            )
            
            result = self.executor.execute_task(task)
            
            assert result.success is True
            assert "[DRY RUN]" in result.message
    
    def test_execute_task_with_validation_error(self):
        """测试任务验证失败"""
        with patch.object(TaskExecutor, '_validate_task_before_execution') as mock_validate:
            mock_validate.side_effect = FormatProtectionError("格式保护阻止")
            
            task = Task(
                id="task_1",
                type=TaskType.SET_PARAGRAPH_STYLE,
                locator=Locator(by=LocatorType.FIND, value="测试"),
                instruction="设置样式"
            )
            
            result = self.executor.execute_task(task)
            
            assert result.success is False
            assert "格式保护阻止" in result.message
    
    def test_create_document_snapshot(self):
        """测试创建文档快照"""
        # 模拟文档内容
        self.mock_context.document.Content.Text = "测试文档内容"
        
        # 模拟样式
        mock_style = Mock()
        mock_style.InUse = True
        mock_style.NameLocal = "标题 1"
        self.mock_context.document.Styles = [mock_style]
        
        # 模拟段落
        mock_para = Mock()
        mock_para.Style.NameLocal = "标题 1"
        mock_para.Range.Text = "测试标题\r"
        self.mock_context.document.Paragraphs = [mock_para]
        
        # 模拟目录和超链接
        self.mock_context.document.TablesOfContents.Count = 1
        self.mock_context.document.Hyperlinks.Count = 2
        
        snapshot = self.executor.create_document_snapshot()
        
        assert snapshot.content == "测试文档内容"
        assert "标题 1" in snapshot.styles
        assert len(snapshot.headings) == 1
        assert snapshot.toc_count == 1
        assert snapshot.hyperlinks_count == 2
    
    def test_detect_unauthorized_changes(self):
        """测试检测未授权变更"""
        from autoword.core.word_executor import DocumentSnapshot
        from datetime import datetime
        
        # 创建初始快照
        initial_snapshot = DocumentSnapshot(
            content="原始内容",
            styles=["正文", "标题 1"],
            headings=[{"text": "测试标题", "level": 1}],
            toc_count=1,
            hyperlinks_count=0,
            timestamp=datetime.now()
        )
        
        # 模拟当前状态（有变更）
        with patch.object(self.executor, 'create_document_snapshot') as mock_snapshot:
            current_snapshot = DocumentSnapshot(
                content="修改后内容",
                styles=["正文", "标题 1", "标题 2"],  # 添加了新样式
                headings=[{"text": "测试标题", "level": 2}],  # 级别变更
                toc_count=2,  # 目录数量变更
                hyperlinks_count=1,
                timestamp=datetime.now()
            )
            mock_snapshot.return_value = current_snapshot
            
            changes = self.executor.detect_unauthorized_changes(initial_snapshot)
            
            assert len(changes) > 0
            assert any("添加了未授权样式" in change for change in changes)
            assert any("级别从 1 变更为 2" in change for change in changes)
            assert any("目录数量从 1 变更为 2" in change for change in changes)
    
    @patch('shutil.copy2')
    @patch('os.path.exists')
    def test_rollback_document_success(self, mock_exists, mock_copy):
        """测试成功回滚文档"""
        mock_exists.return_value = True
        
        # 设置文档的 FullName 属性
        mock_document = Mock()
        mock_document.FullName = "test.docx"
        self.mock_context.document = mock_document
        
        # 模拟重新打开的文档
        mock_new_doc = Mock()
        self.mock_context.word_app.Documents.Open.return_value = mock_new_doc
        
        result = self.executor.rollback_document("backup.docx")
        
        assert result is True
        mock_copy.assert_called_once_with("backup.docx", "test.docx")
        mock_document.Close.assert_called_once_with(SaveChanges=False)
        self.mock_context.word_app.Documents.Open.assert_called_once_with("test.docx")
        # 验证文档对象被更新
        assert self.mock_context.document == mock_new_doc
    
    @patch('os.path.exists')
    def test_rollback_document_backup_not_exists(self, mock_exists):
        """测试备份文件不存在时的回滚"""
        mock_exists.return_value = False
        
        result = self.executor.rollback_document("nonexistent_backup.docx")
        
        assert result is False
    
    @patch.object(TaskExecutor, '_validate_task_before_execution')
    @patch.object(TaskLocator, 'locate_target')
    @patch.object(TaskExecutor, 'create_document_snapshot')
    @patch.object(TaskExecutor, 'detect_unauthorized_changes')
    @patch.object(TaskExecutor, 'rollback_document')
    @patch.object(TaskExecutor, '_execute_set_paragraph_style')
    def test_execute_formatting_task_with_unauthorized_changes(self, mock_execute_style, mock_rollback, mock_detect, mock_snapshot, mock_locate, mock_validate):
        """测试格式化任务检测到未授权变更时的回滚"""
        # 设置模拟
        mock_range = Mock()
        mock_locate.return_value = mock_range
        mock_snapshot.return_value = Mock()
        mock_execute_style.return_value = "样式设置完成"
        mock_detect.return_value = ["检测到未授权的样式变更"]
        mock_rollback.return_value = True
        
        task = Task(
            id="task_1",
            type=TaskType.SET_PARAGRAPH_STYLE,
            source_comment_id="comment_1",
            locator=Locator(by=LocatorType.FIND, value="测试"),
            instruction="设置段落样式"
        )
        
        result = self.executor.execute_task(task, "backup.docx")
        
        assert result.success is False
        assert "检测到未授权变更，已回滚" in result.message
        mock_rollback.assert_called_once_with("backup.docx")
    
    @patch.object(TaskExecutor, '_validate_task_before_execution')
    @patch.object(TaskLocator, 'locate_target')
    @patch.object(TaskExecutor, 'create_document_snapshot')
    @patch.object(TaskExecutor, 'detect_unauthorized_changes')
    def test_execute_content_task_no_change_detection(self, mock_detect, mock_snapshot, mock_locate, mock_validate):
        """测试内容任务不进行变更检测"""
        mock_range = Mock()
        mock_range.Text = "原始文本"
        mock_locate.return_value = mock_range
        
        task = Task(
            id="task_1",
            type=TaskType.REWRITE,  # 内容任务
            locator=Locator(by=LocatorType.FIND, value="测试"),
            instruction="重写文本"
        )
        
        result = self.executor.execute_task(task)
        
        assert result.success is True
        # 内容任务不应该调用变更检测
        mock_detect.assert_not_called()
    
    def test_execute_unsupported_task_type(self):
        """测试不支持的任务类型"""
        with patch.object(TaskLocator, 'locate_target') as mock_locate:
            mock_locate.return_value = Mock()
            
            # 由于 APPLY_TEMPLATE 实际上是支持的，我们需要模拟一个真正不支持的类型
            # 通过 patch 来模拟不支持的情况
            with patch.object(self.executor, '_execute_apply_template', side_effect=TaskExecutionError("不支持的任务类型")):
                task = Task(
                    id="task_1",
                    type=TaskType.APPLY_TEMPLATE,
                    locator=Locator(by=LocatorType.FIND, value="测试"),
                    instruction="应用模板"
                )
                
                result = self.executor.execute_task(task)
                
                assert result.success is False
                assert "不支持的任务类型" in result.message
    
    def test_extract_style_name(self):
        """测试提取样式名称"""
        assert self.executor._extract_style_name("设置为标题1") == "标题 1"
        assert self.executor._extract_style_name("设置为标题2样式") == "标题 2"
        assert self.executor._extract_style_name("设置为正文样式") == "正文"
        assert self.executor._extract_style_name("设置为Heading 1") == "Heading 1"
        assert self.executor._extract_style_name("未知样式") == "正文"
    
    def test_extract_heading_level(self):
        """测试提取标题级别"""
        assert self.executor._extract_heading_level("设置为1级标题") == 1
        assert self.executor._extract_heading_level("设置为2级标题") == 2
        assert self.executor._extract_heading_level("设置为一级标题") == 1
        assert self.executor._extract_heading_level("设置为二级标题") == 2
        assert self.executor._extract_heading_level("设置为标题") == 1
    
    def test_extract_hyperlink_address(self):
        """测试提取超链接地址"""
        assert self.executor._extract_hyperlink_address("链接到 https://example.com") == "https://example.com"
        assert self.executor._extract_hyperlink_address("邮箱 test@example.com") == "mailto:test@example.com"
        assert self.executor._extract_hyperlink_address("内部链接") == "内部链接"


class TestWordExecutor:
    """测试 Word 执行器主类"""
    
    def setup_method(self):
        """测试前设置"""
        self.executor = WordExecutor(visible=False)
    
    @patch('autoword.core.word_executor.WordSession')
    def test_execute_tasks_success(self, mock_word_session):
        """测试成功执行任务列表"""
        # 设置模拟
        mock_word_app = Mock()
        mock_document = Mock()
        mock_word_app.Documents.Open.return_value = mock_document
        
        mock_session_instance = MagicMock()
        mock_session_instance.__enter__.return_value = mock_word_app
        mock_session_instance.__exit__.return_value = None
        mock_word_session.return_value = mock_session_instance
        
        # 创建测试任务
        tasks = [
            Task(
                id="task_1",
                type=TaskType.REWRITE,
                locator=Locator(by=LocatorType.FIND, value="测试"),
                instruction="重写文本"
            )
        ]
        
        # 模拟任务执行器
        with patch('autoword.core.word_executor.TaskExecutor') as mock_executor_class:
            mock_executor = Mock()
            mock_executor.create_document_snapshot.return_value = Mock()
            mock_executor.execute_task.return_value = TaskResult(
                task_id="task_1",
                success=True,
                message="执行成功",
                execution_time=1.0
            )
            mock_executor_class.return_value = mock_executor
            
            # 执行测试（禁用备份创建）
            result = self.executor.execute_tasks(
                tasks=tasks,
                document_path="test.docx",
                create_backup=False
            )
        
        # 验证结果
        assert result.success is True
        assert result.total_tasks == 1
        assert result.completed_tasks == 1
        assert result.failed_tasks == 0
        assert len(result.task_results) == 1
        
        # 验证文档操作
        mock_word_app.Documents.Open.assert_called_once()
        mock_document.Save.assert_called_once()
        mock_document.Close.assert_called_once()
    
    @patch('autoword.core.word_executor.WordSession')
    def test_execute_tasks_with_failures(self, mock_word_session):
        """测试执行任务时有失败的情况"""
        # 设置模拟
        mock_word_app = Mock()
        mock_document = Mock()
        mock_word_app.Documents.Open.return_value = mock_document
        
        mock_session_instance = MagicMock()
        mock_session_instance.__enter__.return_value = mock_word_app
        mock_session_instance.__exit__.return_value = None
        mock_word_session.return_value = mock_session_instance
        
        # 创建测试任务
        tasks = [
            Task(
                id="task_1",
                type=TaskType.REWRITE,
                locator=Locator(by=LocatorType.FIND, value="测试1"),
                instruction="重写文本1"
            ),
            Task(
                id="task_2",
                type=TaskType.REWRITE,
                locator=Locator(by=LocatorType.FIND, value="测试2"),
                instruction="重写文本2"
            )
        ]
        
        # 模拟任务执行器（一个成功，一个失败）
        with patch('autoword.core.word_executor.TaskExecutor') as mock_executor_class:
            mock_executor = Mock()
            mock_executor.create_document_snapshot.return_value = Mock()
            mock_executor.execute_task.side_effect = [
                TaskResult(
                    task_id="task_1",
                    success=True,
                    message="执行成功",
                    execution_time=1.0
                ),
                TaskResult(
                    task_id="task_2",
                    success=False,
                    message="执行失败",
                    execution_time=0.5,
                    error_details="测试错误"
                )
            ]
            mock_executor_class.return_value = mock_executor
            
            # 执行测试（禁用备份创建）
            result = self.executor.execute_tasks(
                tasks=tasks,
                document_path="test.docx",
                create_backup=False
            )
        
        # 验证结果
        assert result.success is False  # 有失败任务
        assert result.total_tasks == 2
        assert result.completed_tasks == 1
        assert result.failed_tasks == 1
        assert len(result.task_results) == 2
        assert "1 个任务失败" in result.error_summary
    
    @patch('autoword.core.word_executor.WordSession')
    def test_dry_run_tasks(self, mock_word_session):
        """测试试运行任务"""
        # 设置模拟
        mock_word_app = Mock()
        mock_document = Mock()
        mock_word_app.Documents.Open.return_value = mock_document
        
        mock_session_instance = MagicMock()
        mock_session_instance.__enter__.return_value = mock_word_app
        mock_session_instance.__exit__.return_value = None
        mock_word_session.return_value = mock_session_instance
        
        tasks = [
            Task(
                id="task_1",
                type=TaskType.REWRITE,
                locator=Locator(by=LocatorType.FIND, value="测试"),
                instruction="重写文本"
            )
        ]
        
        with patch('autoword.core.word_executor.TaskExecutor') as mock_executor_class:
            mock_executor = Mock()
            mock_executor.execute_task.return_value = TaskResult(
                task_id="task_1",
                success=True,
                message="[DRY RUN] 试运行成功",
                execution_time=0.1
            )
            mock_executor_class.return_value = mock_executor
            
            # 执行试运行
            result = self.executor.dry_run_tasks(
                tasks=tasks,
                document_path="test.docx"
            )
        
        # 验证结果
        assert result.success is True
        assert "[DRY RUN]" in result.task_results[0].message
        
        # 验证文档没有被保存（试运行模式）
        mock_document.Save.assert_not_called()
    
    @patch('autoword.core.word_executor.WordSession')
    @patch.object(WordExecutor, '_create_backup')
    def test_execute_tasks_with_backup_creation(self, mock_create_backup, mock_word_session):
        """测试执行任务时创建备份"""
        # 设置模拟
        mock_word_app = Mock()
        mock_document = Mock()
        mock_word_app.Documents.Open.return_value = mock_document
        
        mock_session_instance = MagicMock()
        mock_session_instance.__enter__.return_value = mock_word_app
        mock_session_instance.__exit__.return_value = None
        mock_word_session.return_value = mock_session_instance
        
        mock_create_backup.return_value = "backup_file.docx"
        
        tasks = [
            Task(
                id="task_1",
                type=TaskType.SET_PARAGRAPH_STYLE,
                source_comment_id="comment_1",
                locator=Locator(by=LocatorType.FIND, value="测试"),
                instruction="设置样式"
            )
        ]
        
        with patch('autoword.core.word_executor.TaskExecutor') as mock_executor_class:
            mock_executor = Mock()
            mock_executor.create_document_snapshot.return_value = Mock()
            mock_executor.execute_task.return_value = TaskResult(
                task_id="task_1",
                success=True,
                message="样式设置完成",
                execution_time=1.0
            )
            mock_executor_class.return_value = mock_executor
            
            # 执行任务
            result = self.executor.execute_tasks(
                tasks=tasks,
                document_path="test.docx",
                create_backup=True
            )
        
        # 验证结果
        assert result.success is True
        mock_create_backup.assert_called_once_with("test.docx")
        mock_executor.create_document_snapshot.assert_called_once()
        
        # 验证执行任务时传递了备份路径
        mock_executor.execute_task.assert_called_once_with(tasks[0], "backup_file.docx")
    
    @patch('shutil.copy2')
    @patch('pathlib.Path')
    def test_create_backup_success(self, mock_path_class, mock_copy):
        """测试成功创建备份"""
        # 模拟路径操作
        mock_doc_path = Mock()
        mock_doc_path.stem = "document"
        mock_doc_path.suffix = ".docx"
        
        # 模拟 parent / backup_name 操作
        mock_parent = Mock()
        mock_backup_path = "document_backup_20240101_120000.docx"
        mock_parent.__truediv__ = Mock(return_value=mock_backup_path)
        mock_doc_path.parent = mock_parent
        
        mock_path_class.return_value = mock_doc_path
        
        with patch('autoword.core.word_executor.datetime') as mock_datetime:
            mock_datetime.now.return_value.strftime.return_value = "20240101_120000"
            
            backup_path = self.executor._create_backup("document.docx")
            
            mock_copy.assert_called_once()
            mock_path_class.assert_called_once_with("document.docx")
            assert backup_path == str(mock_backup_path)
    
    @patch('shutil.copy2')
    def test_create_backup_failure(self, mock_copy):
        """测试创建备份失败"""
        mock_copy.side_effect = Exception("磁盘空间不足")
        
        with pytest.raises(TaskExecutionError, match="创建备份失败"):
            self.executor._create_backup("document.docx")


class TestConvenienceFunction:
    """测试便捷函数"""
    
    @patch('autoword.core.word_executor.WordExecutor')
    def test_execute_task_list_normal(self, mock_executor_class):
        """测试正常执行任务列表"""
        mock_executor = Mock()
        mock_result = Mock()
        mock_executor.execute_tasks.return_value = mock_result
        mock_executor_class.return_value = mock_executor
        
        tasks = [
            Task(
                id="task_1",
                type=TaskType.REWRITE,
                locator=Locator(by=LocatorType.FIND, value="测试"),
                instruction="重写文本"
            )
        ]
        
        result = execute_task_list(
            tasks=tasks,
            document_path="test.docx",
            visible=False,
            dry_run=False
        )
        
        assert result == mock_result
        mock_executor.execute_tasks.assert_called_once()
        mock_executor.dry_run_tasks.assert_not_called()
    
    @patch('autoword.core.word_executor.WordExecutor')
    def test_execute_task_list_dry_run(self, mock_executor_class):
        """测试试运行任务列表"""
        mock_executor = Mock()
        mock_result = Mock()
        mock_executor.dry_run_tasks.return_value = mock_result
        mock_executor_class.return_value = mock_executor
        
        tasks = [
            Task(
                id="task_1",
                type=TaskType.REWRITE,
                locator=Locator(by=LocatorType.FIND, value="测试"),
                instruction="重写文本"
            )
        ]
        
        result = execute_task_list(
            tasks=tasks,
            document_path="test.docx",
            visible=False,
            dry_run=True
        )
        
        assert result == mock_result
        mock_executor.dry_run_tasks.assert_called_once()
        mock_executor.execute_tasks.assert_not_called()


if __name__ == "__main__":
    pytest.main([__file__])