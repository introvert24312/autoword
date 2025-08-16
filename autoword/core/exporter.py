"""
AutoWord Exporter
日志和导出系统，负责生成执行报告和日志
"""

import os
import json
import logging
from typing import List, Dict, Any, Optional
from datetime import datetime
from pathlib import Path

from .models import TaskPlan, ExecutionResult, TaskResult, Comment, DocumentStructure
from .exceptions import ValidationError


logger = logging.getLogger(__name__)


class Exporter:
    """导出器 - 负责生成各种报告和日志"""
    
    def __init__(self, output_dir: str = "output"):
        """
        初始化导出器
        
        Args:
            output_dir: 输出目录
        """
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(exist_ok=True)
    
    def export_plan(self, task_plan: TaskPlan, filename: Optional[str] = None) -> str:
        """
        导出任务计划为 JSON 文件
        
        Args:
            task_plan: 任务计划
            filename: 文件名（可选）
            
        Returns:
            导出文件路径
        """
        if filename is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"plan_{timestamp}.json"
        
        filepath = self.output_dir / filename
        
        plan_data = {
            "metadata": {
                "document_path": task_plan.document_path,
                "total_tasks": task_plan.total_tasks,
                "created_at": datetime.now().isoformat(),
                "version": "1.0"
            },
            "tasks": []
        }
        
        for task in task_plan.tasks:
            task_data = {
                "id": task.id,
                "type": task.type.value,
                "source_comment_id": task.source_comment_id,
                "locator": {
                    "by": task.locator.by.value,
                    "value": task.locator.value
                },
                "instruction": task.instruction,
                "dependencies": task.dependencies,
                "risk": task.risk.value,
                "requires_user_review": task.requires_user_review,
                "notes": task.notes
            }
            plan_data["tasks"].append(task_data)
        
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(plan_data, f, ensure_ascii=False, indent=2)
        
        logger.info(f"任务计划已导出: {filepath}")
        return str(filepath)
    
    def export_execution_log(self, execution_result: ExecutionResult, filename: Optional[str] = None) -> str:
        """
        导出执行日志为 JSON 文件
        
        Args:
            execution_result: 执行结果
            filename: 文件名（可选）
            
        Returns:
            导出文件路径
        """
        if filename is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"run_log_{timestamp}.json"
        
        filepath = self.output_dir / filename
        
        log_data = {
            "metadata": {
                "execution_time": execution_result.execution_time,
                "completed_at": datetime.now().isoformat(),
                "version": "1.0"
            },
            "summary": {
                "success": execution_result.success,
                "total_tasks": execution_result.total_tasks,
                "completed_tasks": execution_result.completed_tasks,
                "failed_tasks": execution_result.failed_tasks,
                "error_summary": execution_result.error_summary
            },
            "task_results": []
        }
        
        for result in execution_result.task_results:
            result_data = {
                "task_id": result.task_id,
                "success": result.success,
                "message": result.message,
                "execution_time": result.execution_time,
                "error_details": result.error_details
            }
            log_data["task_results"].append(result_data)
        
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(log_data, f, ensure_ascii=False, indent=2)
        
        logger.info(f"执行日志已导出: {filepath}")
        return str(filepath)
    
    def export_diff_report(self, 
                          before_structure: DocumentStructure,
                          after_structure: DocumentStructure,
                          filename: Optional[str] = None) -> str:
        """
        导出文档变更对比报告
        
        Args:
            before_structure: 执行前文档结构
            after_structure: 执行后文档结构
            filename: 文件名（可选）
            
        Returns:
            导出文件路径
        """
        if filename is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"diff_{timestamp}.md"
        
        filepath = self.output_dir / filename
        
        diff_content = self._generate_diff_markdown(before_structure, after_structure)
        
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(diff_content)
        
        logger.info(f"变更报告已导出: {filepath}")
        return str(filepath)
    
    def _generate_diff_markdown(self, before: DocumentStructure, after: DocumentStructure) -> str:
        """生成 Markdown 格式的变更报告"""
        lines = [
            "# 文档变更报告",
            "",
            f"**生成时间**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            "",
            "## 摘要",
            "",
            f"- 页数变化: {before.page_count} → {after.page_count}",
            f"- 字数变化: {before.word_count} → {after.word_count}",
            f"- 标题数量: {len(before.headings)} → {len(after.headings)}",
            f"- 样式数量: {len(before.styles)} → {len(after.styles)}",
            f"- 目录条目: {len(before.toc_entries)} → {len(after.toc_entries)}",
            f"- 超链接数量: {len(before.hyperlinks)} → {len(after.hyperlinks)}",
            "",
            "## 详细变更",
            ""
        ]
        
        # 标题变更
        lines.extend(self._compare_headings(before.headings, after.headings))
        
        # 样式变更
        lines.extend(self._compare_styles(before.styles, after.styles))
        
        # TOC 变更
        lines.extend(self._compare_toc_entries(before.toc_entries, after.toc_entries))
        
        # 超链接变更
        lines.extend(self._compare_hyperlinks(before.hyperlinks, after.hyperlinks))
        
        return "\n".join(lines)
    
    def _compare_headings(self, before_headings, after_headings) -> List[str]:
        """比较标题变更"""
        lines = ["### 标题变更", ""]
        
        before_dict = {h.text: h for h in before_headings}
        after_dict = {h.text: h for h in after_headings}
        
        # 新增标题
        added = set(after_dict.keys()) - set(before_dict.keys())
        if added:
            lines.append("**新增标题:**")
            for title in sorted(added):
                heading = after_dict[title]
                lines.append(f"- {heading.level}级: {title}")
            lines.append("")
        
        # 删除标题
        removed = set(before_dict.keys()) - set(after_dict.keys())
        if removed:
            lines.append("**删除标题:**")
            for title in sorted(removed):
                heading = before_dict[title]
                lines.append(f"- {heading.level}级: {title}")
            lines.append("")
        
        # 修改标题
        modified = []
        for title in set(before_dict.keys()) & set(after_dict.keys()):
            before_h = before_dict[title]
            after_h = after_dict[title]
            if before_h.level != after_h.level or before_h.style != after_h.style:
                modified.append((title, before_h, after_h))
        
        if modified:
            lines.append("**修改标题:**")
            for title, before_h, after_h in modified:
                changes = []
                if before_h.level != after_h.level:
                    changes.append(f"级别: {before_h.level} → {after_h.level}")
                if before_h.style != after_h.style:
                    changes.append(f"样式: {before_h.style} → {after_h.style}")
                lines.append(f"- {title}: {', '.join(changes)}")
            lines.append("")
        
        if not added and not removed and not modified:
            lines.append("无标题变更")
            lines.append("")
        
        return lines
    
    def _compare_styles(self, before_styles, after_styles) -> List[str]:
        """比较样式变更"""
        lines = ["### 样式变更", ""]
        
        before_names = {s.name for s in before_styles if s.in_use}
        after_names = {s.name for s in after_styles if s.in_use}
        
        added = after_names - before_names
        removed = before_names - after_names
        
        if added:
            lines.append("**新增样式:**")
            for name in sorted(added):
                lines.append(f"- {name}")
            lines.append("")
        
        if removed:
            lines.append("**移除样式:**")
            for name in sorted(removed):
                lines.append(f"- {name}")
            lines.append("")
        
        if not added and not removed:
            lines.append("无样式变更")
            lines.append("")
        
        return lines
    
    def _compare_toc_entries(self, before_entries, after_entries) -> List[str]:
        """比较目录条目变更"""
        lines = ["### 目录变更", ""]
        
        if len(before_entries) != len(after_entries):
            lines.append(f"**条目数量变化**: {len(before_entries)} → {len(after_entries)}")
            lines.append("")
        
        # 简化比较：只比较文本和页码
        before_items = [(e.text, e.page_number) for e in before_entries]
        after_items = [(e.text, e.page_number) for e in after_entries]
        
        if before_items != after_items:
            lines.append("**目录内容已更新**")
            lines.append("")
        else:
            lines.append("无目录变更")
            lines.append("")
        
        return lines
    
    def _compare_hyperlinks(self, before_links, after_links) -> List[str]:
        """比较超链接变更"""
        lines = ["### 超链接变更", ""]
        
        if len(before_links) != len(after_links):
            lines.append(f"**链接数量变化**: {len(before_links)} → {len(after_links)}")
            lines.append("")
        else:
            lines.append("无超链接变更")
            lines.append("")
        
        return lines
    
    def export_comments_data(self, comments: List[Comment], filename: Optional[str] = None) -> str:
        """
        导出批注数据
        
        Args:
            comments: 批注列表
            filename: 文件名（可选）
            
        Returns:
            导出文件路径
        """
        if filename is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"comments_{timestamp}.json"
        
        filepath = self.output_dir / filename
        
        comments_data = {
            "metadata": {
                "total_comments": len(comments),
                "exported_at": datetime.now().isoformat(),
                "version": "1.0"
            },
            "comments": []
        }
        
        for comment in comments:
            comment_data = {
                "id": comment.id,
                "author": comment.author,
                "page": comment.page,
                "anchor_text": comment.anchor_text,
                "comment_text": comment.comment_text,
                "range_start": comment.range_start,
                "range_end": comment.range_end
            }
            comments_data["comments"].append(comment_data)
        
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(comments_data, f, ensure_ascii=False, indent=2)
        
        logger.info(f"批注数据已导出: {filepath}")
        return str(filepath)


# 便捷函数
def export_execution_report(task_plan: TaskPlan, 
                           execution_result: ExecutionResult,
                           before_structure: Optional[DocumentStructure] = None,
                           after_structure: Optional[DocumentStructure] = None,
                           output_dir: str = "output") -> Dict[str, str]:
    """
    便捷函数：导出完整的执行报告
    
    Args:
        task_plan: 任务计划
        execution_result: 执行结果
        before_structure: 执行前文档结构
        after_structure: 执行后文档结构
        output_dir: 输出目录
        
    Returns:
        导出文件路径字典
    """
    exporter = Exporter(output_dir)
    
    exported_files = {}
    
    # 导出任务计划
    exported_files["plan"] = exporter.export_plan(task_plan)
    
    # 导出执行日志
    exported_files["log"] = exporter.export_execution_log(execution_result)
    
    # 导出变更报告
    if before_structure and after_structure:
        exported_files["diff"] = exporter.export_diff_report(before_structure, after_structure)
    
    return exported_files