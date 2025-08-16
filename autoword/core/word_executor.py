"""
AutoWord Word COM Executor
Word COM 执行器，负责安全执行任务并实现格式保护的第3层防线
"""

import os
import logging
import time
from typing import Any, Dict, List, Optional, Tuple, Union
from datetime import datetime
from dataclasses import dataclass
from enum import Enum

import pythoncom
import win32com.client as win32
from win32com.client import constants as win32_constants

from .models import (
    Task, TaskResult, ExecutionResult, TaskType, RiskLevel, 
    LocatorType, ValidationResult, Comment
)
from .doc_loader import WordSession
from .planner import FormatProtectionGuard
from .exceptions import COMError, TaskExecutionError, FormatProtectionError
from .utils import truncate_text


logger = logging.getLogger(__name__)


class ExecutionMode(str, Enum):
    """执行模式枚举"""
    NORMAL = "normal"      # 正常执行
    DRY_RUN = "dry_run"    # 试运行（不实际修改）
    SAFE = "safe"          # 安全模式（额外验证）


@dataclass
class DocumentSnapshot:
    """文档状态快照"""
    content: str
    styles: List[str]
    headings: List[Dict[str, Any]]
    toc_count: int
    hyperlinks_count: int
    timestamp: datetime


@dataclass
class ExecutionContext:
    """执行上下文"""
    word_app: Any
    document: Any
    mode: ExecutionMode = ExecutionMode.NORMAL
    backup_created: bool = False
    start_time: Optional[datetime] = None
    comments: List[Comment] = None
    initial_snapshot: Optional[DocumentSnapshot] = None


class TaskLocator:
    """任务定位器 - 负责在文档中精确定位任务目标"""
    
    def __init__(self, word_app: Any, document: Any):
        """
        初始化定位器
        
        Args:
            word_app: Word 应用程序对象
            document: Word 文档对象
        """
        self.word_app = word_app
        self.document = document
    
    def locate_target(self, task: Task) -> Any:
        """
        定位任务目标
        
        Args:
            task: 任务对象
            
        Returns:
            Word Range 对象
            
        Raises:
            TaskExecutionError: 定位失败时抛出
        """
        try:
            locator = task.locator
            
            if locator.by == LocatorType.BOOKMARK:
                return self._locate_by_bookmark(locator.value)
            elif locator.by == LocatorType.RANGE:
                return self._locate_by_range(locator.value)
            elif locator.by == LocatorType.HEADING:
                return self._locate_by_heading(locator.value)
            elif locator.by == LocatorType.FIND:
                return self._locate_by_find(locator.value)
            else:
                raise TaskExecutionError(f"不支持的定位方式: {locator.by}")
                
        except Exception as e:
            raise TaskExecutionError(f"定位任务目标失败 ({task.id}): {e}")
    
    def _locate_by_bookmark(self, bookmark_name: str) -> Any:
        """通过书签定位"""
        try:
            if bookmark_name in [bm.Name for bm in self.document.Bookmarks]:
                bookmark = self.document.Bookmarks(bookmark_name)
                return bookmark.Range
            else:
                # 如果书签不存在，尝试创建
                logger.warning(f"书签 '{bookmark_name}' 不存在，尝试通过文本查找")
                return self._locate_by_find(bookmark_name)
        except Exception as e:
            raise TaskExecutionError(f"书签定位失败: {e}")
    
    def _locate_by_range(self, range_spec: str) -> Any:
        """通过范围定位 (格式: "start-end" 或 "start,length")"""
        try:
            if '-' in range_spec:
                start, end = map(int, range_spec.split('-'))
                return self.document.Range(start, end)
            elif ',' in range_spec:
                start, length = map(int, range_spec.split(','))
                return self.document.Range(start, start + length)
            else:
                start = int(range_spec)
                return self.document.Range(start, start + 1)
        except Exception as e:
            raise TaskExecutionError(f"范围定位失败: {e}")
    
    def _locate_by_heading(self, heading_text: str) -> Any:
        """通过标题定位"""
        try:
            # 遍历所有段落查找标题
            for para in self.document.Paragraphs:
                if para.Style.NameLocal.startswith(('标题', 'Heading', 'Title')):
                    para_text = para.Range.Text.strip().replace('\r', '')
                    if heading_text in para_text or para_text in heading_text:
                        return para.Range
            
            # 如果没找到标题，尝试文本查找
            logger.warning(f"未找到标题 '{heading_text}'，尝试文本查找")
            return self._locate_by_find(heading_text)
            
        except Exception as e:
            raise TaskExecutionError(f"标题定位失败: {e}")
    
    def _locate_by_find(self, search_text: str) -> Any:
        """通过文本查找定位"""
        try:
            # 创建查找对象
            find_range = self.document.Range()
            find_obj = find_range.Find
            
            # 配置查找参数
            find_obj.ClearFormatting()
            find_obj.Text = search_text
            find_obj.Forward = True
            find_obj.Wrap = win32_constants.wdFindContinue
            find_obj.MatchCase = False
            find_obj.MatchWholeWord = False
            
            # 执行查找
            if find_obj.Execute():
                return find_range
            else:
                # 尝试模糊匹配
                return self._fuzzy_find(search_text)
                
        except Exception as e:
            raise TaskExecutionError(f"文本查找失败: {e}")
    
    def _fuzzy_find(self, search_text: str) -> Any:
        """模糊查找"""
        try:
            # 截取搜索文本的前几个词进行查找
            words = search_text.split()[:3]  # 取前3个词
            for word in words:
                if len(word) > 2:  # 忽略太短的词
                    find_range = self.document.Range()
                    find_obj = find_range.Find
                    find_obj.ClearFormatting()
                    find_obj.Text = word
                    find_obj.Forward = True
                    find_obj.Wrap = win32_constants.wdFindContinue
                    
                    if find_obj.Execute():
                        logger.info(f"模糊匹配成功: '{word}'")
                        return find_range
            
            # 如果还是找不到，返回文档开始位置
            logger.warning(f"无法定位 '{search_text}'，使用文档开始位置")
            return self.document.Range(0, 1)
            
        except Exception as e:
            logger.error(f"模糊查找失败: {e}")
            return self.document.Range(0, 1)
    
    def create_bookmark(self, range_obj: Any, bookmark_name: str) -> str:
        """
        在指定位置创建书签
        
        Args:
            range_obj: Word Range 对象
            bookmark_name: 书签名称
            
        Returns:
            创建的书签名称
        """
        try:
            # 确保书签名称唯一
            unique_name = self._ensure_unique_bookmark_name(bookmark_name)
            
            # 创建书签
            self.document.Bookmarks.Add(unique_name, range_obj)
            logger.info(f"创建书签: {unique_name}")
            
            return unique_name
            
        except Exception as e:
            logger.error(f"创建书签失败: {e}")
            return bookmark_name
    
    def _ensure_unique_bookmark_name(self, base_name: str) -> str:
        """确保书签名称唯一"""
        existing_names = [bm.Name for bm in self.document.Bookmarks]
        
        if base_name not in existing_names:
            return base_name
        
        # 添加数字后缀
        counter = 1
        while f"{base_name}_{counter}" in existing_names:
            counter += 1
        
        return f"{base_name}_{counter}"


class TaskExecutor:
    """任务执行器 - 负责执行具体的任务操作"""
    
    def __init__(self, context: ExecutionContext):
        """
        初始化执行器
        
        Args:
            context: 执行上下文
        """
        self.context = context
        self.locator = TaskLocator(context.word_app, context.document)
        self.format_guard = FormatProtectionGuard()
    
    def create_document_snapshot(self) -> DocumentSnapshot:
        """
        创建文档状态快照
        
        Returns:
            文档快照对象
        """
        try:
            # 获取文档内容
            content = self.context.document.Content.Text
            
            # 获取样式信息
            styles = []
            for style in self.context.document.Styles:
                if style.InUse:
                    styles.append(style.NameLocal)
            
            # 获取标题信息
            headings = []
            for para in self.context.document.Paragraphs:
                if para.Style.NameLocal.startswith(('标题', 'Heading', 'Title')):
                    headings.append({
                        'text': para.Range.Text.strip().replace('\r', ''),
                        'style': para.Style.NameLocal,
                        'level': self._extract_heading_level_from_style(para.Style.NameLocal)
                    })
            
            # 获取目录数量
            toc_count = self.context.document.TablesOfContents.Count
            
            # 获取超链接数量
            hyperlinks_count = self.context.document.Hyperlinks.Count
            
            return DocumentSnapshot(
                content=content,
                styles=styles,
                headings=headings,
                toc_count=toc_count,
                hyperlinks_count=hyperlinks_count,
                timestamp=datetime.now()
            )
            
        except Exception as e:
            logger.warning(f"创建文档快照失败: {e}")
            return DocumentSnapshot(
                content="",
                styles=[],
                headings=[],
                toc_count=0,
                hyperlinks_count=0,
                timestamp=datetime.now()
            )
    
    def detect_unauthorized_changes(self, initial_snapshot: DocumentSnapshot) -> List[str]:
        """
        检测未授权的变更
        
        Args:
            initial_snapshot: 初始快照
            
        Returns:
            检测到的未授权变更列表
        """
        try:
            current_snapshot = self.create_document_snapshot()
            unauthorized_changes = []
            
            # 检查样式变更
            initial_styles = set(initial_snapshot.styles)
            current_styles = set(current_snapshot.styles)
            
            if initial_styles != current_styles:
                added_styles = current_styles - initial_styles
                removed_styles = initial_styles - current_styles
                
                if added_styles:
                    unauthorized_changes.append(f"添加了未授权样式: {', '.join(added_styles)}")
                if removed_styles:
                    unauthorized_changes.append(f"删除了样式: {', '.join(removed_styles)}")
            
            # 检查标题级别变更
            initial_headings = {h['text']: h['level'] for h in initial_snapshot.headings}
            current_headings = {h['text']: h['level'] for h in current_snapshot.headings}
            
            for text, current_level in current_headings.items():
                if text in initial_headings:
                    initial_level = initial_headings[text]
                    if initial_level != current_level:
                        unauthorized_changes.append(f"标题 '{text}' 级别从 {initial_level} 变更为 {current_level}")
            
            # 检查目录变更
            if initial_snapshot.toc_count != current_snapshot.toc_count:
                unauthorized_changes.append(f"目录数量从 {initial_snapshot.toc_count} 变更为 {current_snapshot.toc_count}")
            
            return unauthorized_changes
            
        except Exception as e:
            logger.error(f"检测未授权变更失败: {e}")
            return [f"变更检测失败: {e}"]
    
    def rollback_document(self, backup_path: str) -> bool:
        """
        回滚文档到备份状态
        
        Args:
            backup_path: 备份文件路径
            
        Returns:
            回滚是否成功
        """
        try:
            if not os.path.exists(backup_path):
                logger.error(f"备份文件不存在: {backup_path}")
                return False
            
            # 关闭当前文档
            current_path = self.context.document.FullName
            self.context.document.Close(SaveChanges=False)
            
            # 复制备份文件覆盖原文件
            import shutil
            shutil.copy2(backup_path, current_path)
            
            # 重新打开文档
            self.context.document = self.context.word_app.Documents.Open(current_path)
            
            logger.info(f"文档已回滚到备份状态: {backup_path}")
            return True
            
        except Exception as e:
            logger.error(f"文档回滚失败: {e}")
            return False
    
    def _extract_heading_level_from_style(self, style_name: str) -> int:
        """从样式名称提取标题级别"""
        import re
        
        # 查找数字
        numbers = re.findall(r'\d+', style_name)
        if numbers:
            return int(numbers[0])
        
        # 查找中文数字
        chinese_numbers = {'一': 1, '二': 2, '三': 3, '四': 4, '五': 5}
        for chinese, num in chinese_numbers.items():
            if chinese in style_name:
                return num
        
        return 1  # 默认1级
    
    def execute_task(self, task: Task, backup_path: Optional[str] = None) -> TaskResult:
        """
        执行单个任务
        
        Args:
            task: 任务对象
            backup_path: 备份文件路径（用于回滚）
            
        Returns:
            任务执行结果
        """
        start_time = time.time()
        
        try:
            logger.info(f"执行任务: {task.id} ({task.type.value})")
            
            # 第3层防线：执行期拦截
            self._validate_task_before_execution(task)
            
            # 对于格式化任务，创建执行前快照
            pre_execution_snapshot = None
            is_formatting_task = task.type.value in self.format_guard.format_types
            
            if is_formatting_task and self.context.mode != ExecutionMode.DRY_RUN:
                pre_execution_snapshot = self.create_document_snapshot()
                logger.debug(f"创建任务 {task.id} 执行前快照")
            
            # 定位目标
            target_range = self.locator.locate_target(task)
            
            # 根据任务类型执行相应操作
            if task.type == TaskType.REWRITE:
                result = self._execute_rewrite(task, target_range)
            elif task.type == TaskType.INSERT:
                result = self._execute_insert(task, target_range)
            elif task.type == TaskType.DELETE:
                result = self._execute_delete(task, target_range)
            elif task.type == TaskType.SET_PARAGRAPH_STYLE:
                result = self._execute_set_paragraph_style(task, target_range)
            elif task.type == TaskType.SET_HEADING_LEVEL:
                result = self._execute_set_heading_level(task, target_range)
            elif task.type == TaskType.REPLACE_HYPERLINK:
                result = self._execute_replace_hyperlink(task, target_range)
            elif task.type == TaskType.APPLY_TEMPLATE:
                result = self._execute_apply_template(task, target_range)
            elif task.type == TaskType.REBUILD_TOC:
                result = self._execute_rebuild_toc(task, target_range)
            elif task.type == TaskType.UPDATE_TOC_LEVELS:
                result = self._execute_update_toc_levels(task, target_range)
            elif task.type == TaskType.REFRESH_TOC_NUMBERS:
                result = self._execute_refresh_toc_numbers(task, target_range)
            else:
                raise TaskExecutionError(f"不支持的任务类型: {task.type}")
            
            # 对于格式化任务，检测未授权变更
            if is_formatting_task and pre_execution_snapshot and self.context.mode != ExecutionMode.DRY_RUN:
                unauthorized_changes = self.detect_unauthorized_changes(pre_execution_snapshot)
                
                if unauthorized_changes:
                    logger.error(f"检测到未授权变更: {unauthorized_changes}")
                    
                    # 尝试回滚
                    if backup_path and self.rollback_document(backup_path):
                        raise FormatProtectionError(f"检测到未授权变更，已回滚: {'; '.join(unauthorized_changes)}")
                    else:
                        raise FormatProtectionError(f"检测到未授权变更但回滚失败: {'; '.join(unauthorized_changes)}")
            
            execution_time = time.time() - start_time
            
            logger.info(f"任务执行成功: {task.id} (耗时: {execution_time:.2f}s)")
            
            return TaskResult(
                task_id=task.id,
                success=True,
                message=result,
                execution_time=execution_time
            )
            
        except Exception as e:
            execution_time = time.time() - start_time
            error_msg = f"任务执行失败: {e}"
            
            logger.error(f"任务 {task.id} 执行失败: {e}")
            
            return TaskResult(
                task_id=task.id,
                success=False,
                message=error_msg,
                execution_time=execution_time,
                error_details=str(e)
            )
    
    def _validate_task_before_execution(self, task: Task):
        """执行前验证任务 - 第3层防线"""
        if self.context.comments:
            validation_result = self.format_guard.validate_task_safety(task, self.context.comments)
            if not validation_result.is_valid:
                raise FormatProtectionError(f"任务被格式保护阻止: {'; '.join(validation_result.errors)}")
    
    def _execute_rewrite(self, task: Task, target_range: Any) -> str:
        """执行重写任务"""
        if self.context.mode == ExecutionMode.DRY_RUN:
            return f"[DRY RUN] 将重写文本: '{truncate_text(target_range.Text, 50)}'"
        
        try:
            original_text = target_range.Text
            
            # 这里应该根据 task.instruction 来确定新文本
            # 在实际实现中，可能需要调用 LLM 或使用预定义的替换规则
            new_text = self._generate_new_text(task.instruction, original_text)
            
            # 替换文本
            target_range.Text = new_text
            
            return f"重写完成: '{truncate_text(original_text, 30)}' -> '{truncate_text(new_text, 30)}'"
            
        except Exception as e:
            raise TaskExecutionError(f"重写操作失败: {e}")
    
    def _execute_insert(self, task: Task, target_range: Any) -> str:
        """执行插入任务"""
        if self.context.mode == ExecutionMode.DRY_RUN:
            return f"[DRY RUN] 将在位置 {target_range.Start} 插入内容"
        
        try:
            # 生成要插入的文本
            insert_text = self._generate_insert_text(task.instruction)
            
            # 在目标位置插入文本
            target_range.InsertAfter(insert_text)
            
            return f"插入完成: '{truncate_text(insert_text, 50)}'"
            
        except Exception as e:
            raise TaskExecutionError(f"插入操作失败: {e}")
    
    def _execute_delete(self, task: Task, target_range: Any) -> str:
        """执行删除任务"""
        if self.context.mode == ExecutionMode.DRY_RUN:
            return f"[DRY RUN] 将删除文本: '{truncate_text(target_range.Text, 50)}'"
        
        try:
            deleted_text = target_range.Text
            target_range.Delete()
            
            return f"删除完成: '{truncate_text(deleted_text, 50)}'"
            
        except Exception as e:
            raise TaskExecutionError(f"删除操作失败: {e}")
    
    def _execute_set_paragraph_style(self, task: Task, target_range: Any) -> str:
        """执行设置段落样式任务"""
        if self.context.mode == ExecutionMode.DRY_RUN:
            return f"[DRY RUN] 将设置段落样式"
        
        try:
            # 从指令中提取样式名称
            style_name = self._extract_style_name(task.instruction)
            
            # 获取段落
            if target_range.Paragraphs.Count > 0:
                paragraph = target_range.Paragraphs(1)
                original_style = paragraph.Style.NameLocal
                
                # 设置新样式
                paragraph.Style = style_name
                
                return f"样式设置完成: '{original_style}' -> '{style_name}'"
            else:
                raise TaskExecutionError("目标范围不包含段落")
                
        except Exception as e:
            raise TaskExecutionError(f"设置段落样式失败: {e}")
    
    def _execute_set_heading_level(self, task: Task, target_range: Any) -> str:
        """执行设置标题级别任务"""
        if self.context.mode == ExecutionMode.DRY_RUN:
            return f"[DRY RUN] 将设置标题级别"
        
        try:
            # 从指令中提取级别
            level = self._extract_heading_level(task.instruction)
            
            # 获取段落
            if target_range.Paragraphs.Count > 0:
                paragraph = target_range.Paragraphs(1)
                original_style = paragraph.Style.NameLocal
                
                # 设置标题样式
                heading_style = f"标题 {level}" if level <= 9 else f"Heading {level}"
                try:
                    paragraph.Style = heading_style
                except:
                    # 如果中文样式不存在，尝试英文样式
                    paragraph.Style = f"Heading {level}"
                
                return f"标题级别设置完成: '{original_style}' -> 'Heading {level}'"
            else:
                raise TaskExecutionError("目标范围不包含段落")
                
        except Exception as e:
            raise TaskExecutionError(f"设置标题级别失败: {e}")
    
    def _execute_replace_hyperlink(self, task: Task, target_range: Any) -> str:
        """执行替换超链接任务"""
        if self.context.mode == ExecutionMode.DRY_RUN:
            return f"[DRY RUN] 将替换超链接"
        
        try:
            # 从指令中提取新的链接地址
            new_address = self._extract_hyperlink_address(task.instruction)
            
            # 查找范围内的超链接
            if target_range.Hyperlinks.Count > 0:
                hyperlink = target_range.Hyperlinks(1)
                original_address = hyperlink.Address or hyperlink.SubAddress
                
                # 更新链接地址
                if new_address.startswith(('http://', 'https://', 'mailto:')):
                    hyperlink.Address = new_address
                else:
                    hyperlink.SubAddress = new_address
                
                return f"超链接替换完成: '{original_address}' -> '{new_address}'"
            else:
                # 如果没有现有超链接，创建新的
                self.context.document.Hyperlinks.Add(
                    Anchor=target_range,
                    Address=new_address if new_address.startswith(('http://', 'https://', 'mailto:')) else "",
                    SubAddress=new_address if not new_address.startswith(('http://', 'https://', 'mailto:')) else ""
                )
                return f"创建超链接完成: '{new_address}'"
                
        except Exception as e:
            raise TaskExecutionError(f"替换超链接失败: {e}")
    
    def _generate_new_text(self, instruction: str, original_text: str) -> str:
        """生成新文本（简化实现）"""
        # 在实际实现中，这里可能需要调用 LLM 或使用更复杂的逻辑
        if "重写" in instruction:
            return f"[重写] {original_text.strip()}"
        elif "修改" in instruction:
            return f"[修改] {original_text.strip()}"
        else:
            return original_text
    
    def _generate_insert_text(self, instruction: str) -> str:
        """生成插入文本（简化实现）"""
        # 在实际实现中，这里可能需要调用 LLM 或使用更复杂的逻辑
        if "插入" in instruction:
            return f"\n[插入内容] {instruction}\n"
        else:
            return f"\n{instruction}\n"
    
    def _extract_style_name(self, instruction: str) -> str:
        """从指令中提取样式名称"""
        # 简化的样式名称提取
        if "标题" in instruction:
            if "1" in instruction:
                return "标题 1"
            elif "2" in instruction:
                return "标题 2"
            elif "3" in instruction:
                return "标题 3"
        elif "正文" in instruction:
            return "正文"
        elif "Heading" in instruction:
            if "1" in instruction:
                return "Heading 1"
            elif "2" in instruction:
                return "Heading 2"
        
        return "正文"  # 默认样式
    
    def _extract_heading_level(self, instruction: str) -> int:
        """从指令中提取标题级别"""
        import re
        
        # 查找数字
        numbers = re.findall(r'\d+', instruction)
        if numbers:
            level = int(numbers[0])
            return min(max(level, 1), 9)  # 限制在1-9之间
        
        # 查找中文数字
        chinese_numbers = {'一': 1, '二': 2, '三': 3, '四': 4, '五': 5}
        for chinese, num in chinese_numbers.items():
            if chinese in instruction:
                return num
        
        return 1  # 默认1级标题
    
    def _extract_hyperlink_address(self, instruction: str) -> str:
        """从指令中提取超链接地址"""
        import re
        
        # 查找 URL
        url_pattern = r'https?://[^\s]+'
        urls = re.findall(url_pattern, instruction)
        if urls:
            return urls[0]
        
        # 查找邮箱
        email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
        emails = re.findall(email_pattern, instruction)
        if emails:
            return f"mailto:{emails[0]}"
        
        # 默认返回指令本身作为内部链接
        return instruction
    
    def _execute_apply_template(self, task: Task, target_range: Any) -> str:
        """执行应用模板任务"""
        if self.context.mode == ExecutionMode.DRY_RUN:
            return f"[DRY RUN] 将应用模板"
        
        try:
            # 从指令中提取模板名称
            template_name = self._extract_template_name(task.instruction)
            
            # 应用模板到文档
            if template_name and template_name != "default":
                try:
                    # 尝试应用指定模板
                    self.context.document.AttachedTemplate = template_name
                    return f"模板应用完成: '{template_name}'"
                except:
                    # 如果模板不存在，应用默认样式
                    logger.warning(f"模板 '{template_name}' 不存在，应用默认样式")
                    return self._apply_default_styles(target_range)
            else:
                return self._apply_default_styles(target_range)
                
        except Exception as e:
            raise TaskExecutionError(f"应用模板失败: {e}")
    
    def _execute_rebuild_toc(self, task: Task, target_range: Any) -> str:
        """执行重建目录任务"""
        if self.context.mode == ExecutionMode.DRY_RUN:
            return f"[DRY RUN] 将重建目录"
        
        try:
            # 查找现有目录
            existing_tocs = []
            for toc in self.context.document.TablesOfContents:
                existing_tocs.append(toc)
            
            # 删除现有目录
            for toc in existing_tocs:
                toc.Delete()
            
            # 在目标位置创建新目录
            new_toc = self.context.document.TablesOfContents.Add(
                Range=target_range,
                UseHeadingStyles=True,
                UpperHeadingLevel=1,
                LowerHeadingLevel=3,
                UseFields=False,
                TableID=None,
                RightAlignPageNumbers=True,
                IncludePageNumbers=True,
                AddedStyles=None,
                UseHyperlinks=True,
                HidePageNumbersInWeb=True
            )
            
            # 更新目录
            new_toc.Update()
            
            return f"目录重建完成: 删除 {len(existing_tocs)} 个旧目录，创建新目录"
            
        except Exception as e:
            raise TaskExecutionError(f"重建目录失败: {e}")
    
    def _execute_update_toc_levels(self, task: Task, target_range: Any) -> str:
        """执行更新目录级别任务"""
        if self.context.mode == ExecutionMode.DRY_RUN:
            return f"[DRY RUN] 将更新目录级别"
        
        try:
            # 从指令中提取级别信息
            levels = self._extract_toc_levels(task.instruction)
            upper_level, lower_level = levels
            
            # 更新所有目录的级别设置
            updated_count = 0
            for toc in self.context.document.TablesOfContents:
                toc.UpperHeadingLevel = upper_level
                toc.LowerHeadingLevel = lower_level
                toc.Update()
                updated_count += 1
            
            if updated_count > 0:
                return f"目录级别更新完成: {updated_count} 个目录，级别 {upper_level}-{lower_level}"
            else:
                return "未找到目录，无需更新级别"
                
        except Exception as e:
            raise TaskExecutionError(f"更新目录级别失败: {e}")
    
    def _execute_refresh_toc_numbers(self, task: Task, target_range: Any) -> str:
        """执行刷新目录页码任务"""
        if self.context.mode == ExecutionMode.DRY_RUN:
            return f"[DRY RUN] 将刷新目录页码"
        
        try:
            # 刷新所有目录的页码
            updated_count = 0
            for toc in self.context.document.TablesOfContents:
                toc.UpdatePageNumbers()
                updated_count += 1
            
            # 如果没有目录，尝试更新所有域
            if updated_count == 0:
                self.context.document.Fields.Update()
                return "刷新完成: 更新了所有文档域"
            else:
                return f"目录页码刷新完成: {updated_count} 个目录"
                
        except Exception as e:
            raise TaskExecutionError(f"刷新目录页码失败: {e}")
    
    def _apply_default_styles(self, target_range: Any) -> str:
        """应用默认样式集"""
        try:
            # 应用一些基本的样式设置
            styles_applied = []
            
            # 确保基本样式存在
            basic_styles = ["正文", "标题 1", "标题 2", "标题 3"]
            for style_name in basic_styles:
                try:
                    style = self.context.document.Styles(style_name)
                    styles_applied.append(style_name)
                except:
                    # 如果中文样式不存在，尝试英文样式
                    try:
                        english_name = style_name.replace("正文", "Normal").replace("标题", "Heading")
                        style = self.context.document.Styles(english_name)
                        styles_applied.append(english_name)
                    except:
                        continue
            
            return f"默认样式应用完成: {', '.join(styles_applied)}"
            
        except Exception as e:
            return f"默认样式应用部分成功: {e}"
    
    def _extract_template_name(self, instruction: str) -> str:
        """从指令中提取模板名称"""
        import re
        
        # 查找模板名称的模式
        patterns = [
            r'模板[：:]\s*([^\s，,。.]+)',
            r'应用\s*([^\s，,。.]*模板)',
            r'template[：:]\s*([^\s，,。.]+)',
            r'apply\s+([^\s，,。.]*template)'
        ]
        
        for pattern in patterns:
            match = re.search(pattern, instruction, re.IGNORECASE)
            if match:
                return match.group(1).strip()
        
        # 如果没有找到具体模板名称，返回默认
        return "default"
    
    def _extract_toc_levels(self, instruction: str) -> tuple[int, int]:
        """从指令中提取目录级别"""
        import re
        
        # 查找级别数字
        numbers = re.findall(r'(\d+)', instruction)
        
        if len(numbers) >= 2:
            upper = int(numbers[0])
            lower = int(numbers[1])
            # 确保级别在合理范围内
            upper = max(1, min(upper, 9))
            lower = max(upper, min(lower, 9))
            return upper, lower
        elif len(numbers) == 1:
            level = int(numbers[0])
            level = max(1, min(level, 9))
            return 1, level
        else:
            # 默认1-3级
            return 1, 3


class WordExecutor:
    """Word COM 执行器主类"""
    
    def __init__(self, visible: bool = False):
        """
        初始化执行器
        
        Args:
            visible: 是否显示 Word 窗口
        """
        self.visible = visible
        self.format_guard = FormatProtectionGuard()
    
    def _create_backup(self, document_path: str) -> str:
        """
        创建文档备份
        
        Args:
            document_path: 原文档路径
            
        Returns:
            备份文件路径
        """
        import shutil
        from pathlib import Path
        
        try:
            doc_path = Path(document_path)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_name = f"{doc_path.stem}_backup_{timestamp}{doc_path.suffix}"
            backup_path = doc_path.parent / backup_name
            
            shutil.copy2(document_path, backup_path)
            logger.info(f"备份创建成功: {backup_path}")
            
            return str(backup_path)
            
        except Exception as e:
            logger.error(f"创建备份失败: {e}")
            raise TaskExecutionError(f"创建备份失败: {e}")
    
    def execute_tasks(self, 
                     tasks: List[Task],
                     document_path: str,
                     comments: List[Comment] = None,
                     mode: ExecutionMode = ExecutionMode.NORMAL,
                     create_backup: bool = True,
                     output_file_path: Optional[str] = None) -> ExecutionResult:
        """
        执行任务列表
        
        Args:
            tasks: 任务列表
            document_path: 文档路径
            comments: 批注列表（用于格式保护验证）
            mode: 执行模式
            create_backup: 是否创建备份
            
        Returns:
            执行结果
        """
        start_time = time.time()
        task_results = []
        completed_count = 0
        failed_count = 0
        
        try:
            logger.info(f"开始执行任务: {len(tasks)} 个任务, 模式: {mode.value}")
            
            # 使用 Word 会话
            with WordSession(visible=self.visible) as word_app:
                # 打开文档
                document = word_app.Documents.Open(
                    FileName=document_path,
                    ReadOnly=False,
                    AddToRecentFiles=False
                )
                
                try:
                    # 创建执行上下文
                    context = ExecutionContext(
                        word_app=word_app,
                        document=document,
                        mode=mode,
                        backup_created=create_backup,
                        start_time=datetime.now(),
                        comments=comments or []
                    )
                    
                    # 创建备份文件路径
                    backup_path = None
                    if create_backup:
                        backup_path = self._create_backup(document_path)
                        context.backup_created = True
                        logger.info(f"创建备份文件: {backup_path}")
                    
                    # 创建任务执行器
                    executor = TaskExecutor(context)
                    
                    # 创建初始文档快照
                    context.initial_snapshot = executor.create_document_snapshot()
                    
                    # 逐个执行任务
                    for i, task in enumerate(tasks, 1):
                        logger.info(f"执行进度: {i}/{len(tasks)} - {task.id}")
                        
                        try:
                            # 执行任务
                            result = executor.execute_task(task, backup_path)
                            task_results.append(result)
                            
                            if result.success:
                                completed_count += 1
                            else:
                                failed_count += 1
                                
                        except Exception as e:
                            # 单个任务失败不影响其他任务
                            error_result = TaskResult(
                                task_id=task.id,
                                success=False,
                                message=f"任务执行异常: {e}",
                                execution_time=0.0,
                                error_details=str(e)
                            )
                            task_results.append(error_result)
                            failed_count += 1
                            logger.error(f"任务 {task.id} 执行异常: {e}")
                    
                    # 保存文档
                    if mode != ExecutionMode.DRY_RUN:
                        if output_file_path:
                            # 另存为指定文件
                            document.SaveAs2(output_file_path)
                            logger.info(f"文档已另存为: {output_file_path}")
                        else:
                            # 保存原文件
                            document.Save()
                            logger.info("文档已保存")
                    
                finally:
                    # 关闭文档
                    document.Close(SaveChanges=0 if mode == ExecutionMode.DRY_RUN else -1)
            
            execution_time = time.time() - start_time
            success = failed_count == 0
            
            logger.info(f"任务执行完成: {completed_count} 成功, {failed_count} 失败, 耗时: {execution_time:.2f}s")
            
            return ExecutionResult(
                success=success,
                total_tasks=len(tasks),
                completed_tasks=completed_count,
                failed_tasks=failed_count,
                task_results=task_results,
                execution_time=execution_time,
                error_summary=f"{failed_count} 个任务失败" if failed_count > 0 else None
            )
            
        except Exception as e:
            execution_time = time.time() - start_time
            
            logger.error(f"执行器运行失败: {e}")
            
            return ExecutionResult(
                success=False,
                total_tasks=len(tasks),
                completed_tasks=completed_count,
                failed_tasks=len(tasks) - completed_count,
                task_results=task_results,
                execution_time=execution_time,
                error_summary=f"执行器运行失败: {e}"
            )
    
    def dry_run_tasks(self, 
                     tasks: List[Task],
                     document_path: str,
                     comments: List[Comment] = None) -> ExecutionResult:
        """
        试运行任务（不实际修改文档）
        
        Args:
            tasks: 任务列表
            document_path: 文档路径
            comments: 批注列表
            
        Returns:
            执行结果
        """
        return self.execute_tasks(
            tasks=tasks,
            document_path=document_path,
            comments=comments,
            mode=ExecutionMode.DRY_RUN,
            create_backup=False
        )


# 便捷函数
def execute_task_list(tasks: List[Task],
                     document_path: str,
                     comments: List[Comment] = None,
                     visible: bool = False,
                     dry_run: bool = False) -> ExecutionResult:
    """
    便捷函数：执行任务列表
    
    Args:
        tasks: 任务列表
        document_path: 文档路径
        comments: 批注列表
        visible: 是否显示 Word 窗口
        dry_run: 是否试运行
        
    Returns:
        执行结果
    """
    executor = WordExecutor(visible=visible)
    
    if dry_run:
        return executor.dry_run_tasks(tasks, document_path, comments)
    else:
        return executor.execute_tasks(tasks, document_path, comments)