"""
AutoWord Format Validator
格式验证器 - 实现第4层防线：事后校验回滚
"""

import logging
from typing import List, Dict, Any, Optional, Tuple
from dataclasses import dataclass
from datetime import datetime

from .models import (
    Task, Comment, DocumentStructure, DocumentSnapshot, 
    ValidationResult, TaskType
)
from .constants import FORMAT_TYPES
from .exceptions import FormatProtectionError


logger = logging.getLogger(__name__)


@dataclass
class FormatChange:
    """格式变更记录"""
    change_type: str
    element_type: str
    element_id: str
    old_value: Any
    new_value: Any
    authorized: bool
    source_comment_id: Optional[str] = None
    timestamp: datetime = None
    
    def __post_init__(self):
        if self.timestamp is None:
            self.timestamp = datetime.now()


@dataclass
class ValidationReport:
    """验证报告"""
    is_valid: bool
    authorized_changes: List[FormatChange]
    unauthorized_changes: List[FormatChange]
    warnings: List[str]
    recommendations: List[str]
    validation_time: datetime
    
    @property
    def total_changes(self) -> int:
        return len(self.authorized_changes) + len(self.unauthorized_changes)
    
    @property
    def unauthorized_count(self) -> int:
        return len(self.unauthorized_changes)


class DocumentComparator:
    """文档比较器 - 比较执行前后的文档状态"""
    
    def __init__(self):
        """初始化比较器"""
        pass
    
    def compare_snapshots(self, 
                         before: DocumentSnapshot, 
                         after: DocumentSnapshot) -> List[FormatChange]:
        """
        比较两个文档快照
        
        Args:
            before: 执行前快照
            after: 执行后快照
            
        Returns:
            格式变更列表
        """
        changes = []
        
        # 比较标题
        changes.extend(self._compare_headings(before.structure.headings, after.structure.headings))
        
        # 比较样式
        changes.extend(self._compare_styles(before.structure.styles, after.structure.styles))
        
        # 比较目录
        changes.extend(self._compare_toc(before.structure.toc_entries, after.structure.toc_entries))
        
        # 比较超链接
        changes.extend(self._compare_hyperlinks(before.structure.hyperlinks, after.structure.hyperlinks))
        
        return changes
    
    def _compare_headings(self, before_headings, after_headings) -> List[FormatChange]:
        """比较标题变更"""
        changes = []
        
        # 创建标题映射（基于位置范围）
        before_map = {(h.range_start, h.range_end): h for h in before_headings}
        after_map = {(h.range_start, h.range_end): h for h in after_headings}
        
        # 检查级别变更
        for pos, after_heading in after_map.items():
            if pos in before_map:
                before_heading = before_map[pos]
                if before_heading.level != after_heading.level:
                    changes.append(FormatChange(
                        change_type="heading_level_change",
                        element_type="heading",
                        element_id=f"heading_{pos[0]}_{pos[1]}",
                        old_value=before_heading.level,
                        new_value=after_heading.level,
                        authorized=False  # 需要后续验证
                    ))
                
                if before_heading.style != after_heading.style:
                    changes.append(FormatChange(
                        change_type="heading_style_change",
                        element_type="heading",
                        element_id=f"heading_{pos[0]}_{pos[1]}",
                        old_value=before_heading.style,
                        new_value=after_heading.style,
                        authorized=False  # 需要后续验证
                    ))
        
        # 检查新增标题
        for pos, after_heading in after_map.items():
            if pos not in before_map:
                changes.append(FormatChange(
                    change_type="heading_added",
                    element_type="heading",
                    element_id=f"heading_{pos[0]}_{pos[1]}",
                    old_value=None,
                    new_value=after_heading.level,
                    authorized=False  # 需要后续验证
                ))
        
        # 检查删除的标题
        for pos, before_heading in before_map.items():
            if pos not in after_map:
                changes.append(FormatChange(
                    change_type="heading_removed",
                    element_type="heading",
                    element_id=f"heading_{pos[0]}_{pos[1]}",
                    old_value=before_heading.level,
                    new_value=None,
                    authorized=False  # 需要后续验证
                ))
        
        return changes
    
    def _compare_styles(self, before_styles, after_styles) -> List[FormatChange]:
        """比较样式变更"""
        changes = []
        
        before_map = {s.name: s for s in before_styles}
        after_map = {s.name: s for s in after_styles}
        
        # 检查样式使用状态变更
        for name, after_style in after_map.items():
            if name in before_map:
                before_style = before_map[name]
                if before_style.in_use != after_style.in_use:
                    changes.append(FormatChange(
                        change_type="style_usage_change",
                        element_type="style",
                        element_id=name,
                        old_value=before_style.in_use,
                        new_value=after_style.in_use,
                        authorized=False  # 需要后续验证
                    ))
        
        return changes
    
    def _compare_toc(self, before_toc, after_toc) -> List[FormatChange]:
        """比较目录变更"""
        changes = []
        
        # 比较目录条目数量
        if len(before_toc) != len(after_toc):
            changes.append(FormatChange(
                change_type="toc_structure_change",
                element_type="toc",
                element_id="toc_entries",
                old_value=len(before_toc),
                new_value=len(after_toc),
                authorized=False  # 需要后续验证
            ))
        
        # 比较目录级别分布
        before_levels = {}
        after_levels = {}
        
        for entry in before_toc:
            before_levels[entry.level] = before_levels.get(entry.level, 0) + 1
        
        for entry in after_toc:
            after_levels[entry.level] = after_levels.get(entry.level, 0) + 1
        
        if before_levels != after_levels:
            changes.append(FormatChange(
                change_type="toc_levels_change",
                element_type="toc",
                element_id="toc_levels",
                old_value=before_levels,
                new_value=after_levels,
                authorized=False  # 需要后续验证
            ))
        
        return changes
    
    def _compare_hyperlinks(self, before_links, after_links) -> List[FormatChange]:
        """比较超链接变更"""
        changes = []
        
        # 创建链接映射（基于位置）
        before_map = {(h.range_start, h.range_end): h for h in before_links}
        after_map = {(h.range_start, h.range_end): h for h in after_links}
        
        # 检查链接地址变更
        for pos, after_link in after_map.items():
            if pos in before_map:
                before_link = before_map[pos]
                if before_link.address != after_link.address:
                    changes.append(FormatChange(
                        change_type="hyperlink_address_change",
                        element_type="hyperlink",
                        element_id=f"link_{pos[0]}_{pos[1]}",
                        old_value=before_link.address,
                        new_value=after_link.address,
                        authorized=False  # 需要后续验证
                    ))
        
        return changes


class AuthorizationChecker:
    """授权检查器 - 验证格式变更是否有相应的批注授权"""
    
    def __init__(self):
        """初始化授权检查器"""
        pass
    
    def check_authorization(self, 
                          changes: List[FormatChange],
                          executed_tasks: List[Task],
                          comments: List[Comment]) -> List[FormatChange]:
        """
        检查格式变更的授权状态
        
        Args:
            changes: 格式变更列表
            executed_tasks: 已执行的任务列表
            comments: 批注列表
            
        Returns:
            更新授权状态的格式变更列表
        """
        # 创建任务映射
        task_map = {task.id: task for task in executed_tasks}
        comment_map = {comment.id: comment for comment in comments}
        
        # 创建位置到任务的映射
        position_to_tasks = self._create_position_mapping(executed_tasks)
        
        for change in changes:
            # 检查是否有对应的授权任务
            authorized_task = self._find_authorizing_task(change, position_to_tasks, task_map)
            
            if authorized_task:
                change.authorized = True
                change.source_comment_id = authorized_task.source_comment_id
                logger.debug(f"格式变更已授权: {change.change_type} by task {authorized_task.id}")
            else:
                change.authorized = False
                logger.warning(f"未授权的格式变更: {change.change_type}")
        
        return changes
    
    def _create_position_mapping(self, tasks: List[Task]) -> Dict[str, List[Task]]:
        """创建位置到任务的映射"""
        position_map = {}
        
        for task in tasks:
            # 根据定位器类型创建位置键
            if task.locator.by.value == "range":
                key = f"range_{task.locator.value}"
            elif task.locator.by.value == "bookmark":
                key = f"bookmark_{task.locator.value}"
            elif task.locator.by.value == "heading":
                key = f"heading_{task.locator.value}"
            elif task.locator.by.value == "find":
                key = f"find_{task.locator.value}"
            else:
                key = f"unknown_{task.locator.value}"
            
            if key not in position_map:
                position_map[key] = []
            position_map[key].append(task)
        
        return position_map
    
    def _find_authorizing_task(self, 
                             change: FormatChange,
                             position_map: Dict[str, List[Task]],
                             task_map: Dict[str, Task]) -> Optional[Task]:
        """查找授权此变更的任务"""
        # 根据变更类型查找可能的授权任务
        relevant_task_types = self._get_relevant_task_types(change.change_type)
        
        # 在位置映射中查找相关任务
        for tasks in position_map.values():
            for task in tasks:
                if (task.type in relevant_task_types and 
                    task.source_comment_id and  # 必须有批注授权
                    self._is_task_relevant_to_change(task, change)):
                    return task
        
        return None
    
    def _get_relevant_task_types(self, change_type: str) -> List[TaskType]:
        """获取与变更类型相关的任务类型"""
        type_mapping = {
            "heading_level_change": [TaskType.SET_HEADING_LEVEL],
            "heading_style_change": [TaskType.SET_HEADING_LEVEL, TaskType.SET_PARAGRAPH_STYLE],
            "style_usage_change": [TaskType.SET_PARAGRAPH_STYLE, TaskType.APPLY_TEMPLATE],
            "toc_structure_change": [TaskType.REBUILD_TOC, TaskType.UPDATE_TOC_LEVELS],
            "toc_levels_change": [TaskType.UPDATE_TOC_LEVELS],
            "hyperlink_address_change": [TaskType.REPLACE_HYPERLINK],
        }
        
        return type_mapping.get(change_type, [])
    
    def _is_task_relevant_to_change(self, task: Task, change: FormatChange) -> bool:
        """判断任务是否与变更相关"""
        # 这里可以实现更复杂的相关性判断逻辑
        # 目前简化为基于任务类型的判断
        relevant_types = self._get_relevant_task_types(change.change_type)
        return task.type in relevant_types


class FormatValidator:
    """格式验证器主类 - 第4层防线：事后校验回滚"""
    
    def __init__(self):
        """初始化格式验证器"""
        self.comparator = DocumentComparator()
        self.auth_checker = AuthorizationChecker()
    
    def validate_execution_result(self,
                                before_snapshot: DocumentSnapshot,
                                after_snapshot: DocumentSnapshot,
                                executed_tasks: List[Task],
                                comments: List[Comment]) -> ValidationReport:
        """
        验证执行结果
        
        Args:
            before_snapshot: 执行前快照
            after_snapshot: 执行后快照
            executed_tasks: 已执行的任务列表
            comments: 批注列表
            
        Returns:
            验证报告
        """
        start_time = datetime.now()
        
        try:
            logger.info("开始格式验证 - 第4层防线：事后校验")
            
            # 1. 比较文档快照，识别所有变更
            all_changes = self.comparator.compare_snapshots(before_snapshot, after_snapshot)
            logger.info(f"检测到 {len(all_changes)} 个格式变更")
            
            # 2. 检查变更的授权状态
            authorized_changes = self.auth_checker.check_authorization(
                all_changes, executed_tasks, comments
            )
            
            # 3. 分类变更
            authorized = [c for c in authorized_changes if c.authorized]
            unauthorized = [c for c in authorized_changes if not c.authorized]
            
            # 4. 生成警告和建议
            warnings = self._generate_warnings(unauthorized)
            recommendations = self._generate_recommendations(unauthorized)
            
            # 5. 判断整体验证结果
            is_valid = len(unauthorized) == 0
            
            logger.info(f"格式验证完成: {len(authorized)} 个授权变更, {len(unauthorized)} 个未授权变更")
            
            return ValidationReport(
                is_valid=is_valid,
                authorized_changes=authorized,
                unauthorized_changes=unauthorized,
                warnings=warnings,
                recommendations=recommendations,
                validation_time=start_time
            )
            
        except Exception as e:
            logger.error(f"格式验证失败: {e}")
            
            return ValidationReport(
                is_valid=False,
                authorized_changes=[],
                unauthorized_changes=[],
                warnings=[f"格式验证过程出错: {e}"],
                recommendations=["建议检查文档状态并重新执行验证"],
                validation_time=start_time
            )
    
    def _generate_warnings(self, unauthorized_changes: List[FormatChange]) -> List[str]:
        """生成警告信息"""
        warnings = []
        
        if unauthorized_changes:
            warnings.append(f"检测到 {len(unauthorized_changes)} 个未授权的格式变更")
            
            # 按变更类型分组
            change_types = {}
            for change in unauthorized_changes:
                change_type = change.change_type
                if change_type not in change_types:
                    change_types[change_type] = 0
                change_types[change_type] += 1
            
            for change_type, count in change_types.items():
                warnings.append(f"未授权的 {change_type}: {count} 个")
        
        return warnings
    
    def _generate_recommendations(self, unauthorized_changes: List[FormatChange]) -> List[str]:
        """生成建议"""
        recommendations = []
        
        if unauthorized_changes:
            recommendations.append("建议回滚到执行前的文档状态")
            recommendations.append("检查任务规划器的格式保护设置")
            recommendations.append("确保所有格式化任务都有对应的批注授权")
            
            # 根据变更类型提供具体建议
            change_types = set(c.change_type for c in unauthorized_changes)
            
            if "heading_level_change" in change_types:
                recommendations.append("标题级别变更需要明确的批注授权")
            
            if "toc_structure_change" in change_types:
                recommendations.append("目录结构变更属于高风险操作，需要特别授权")
            
            if "hyperlink_address_change" in change_types:
                recommendations.append("超链接地址变更需要在批注中明确指定新地址")
        
        return recommendations
    
    def should_rollback(self, validation_report: ValidationReport) -> bool:
        """
        判断是否应该回滚
        
        Args:
            validation_report: 验证报告
            
        Returns:
            是否应该回滚
        """
        # 如果有未授权的格式变更，建议回滚
        if validation_report.unauthorized_count > 0:
            return True
        
        return False