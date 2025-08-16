"""
AutoWord Task Planner
任务规划器，负责生成和验证任务计划，实现四重格式保护防线
"""

import json
import logging
from typing import List, Dict, Any, Optional, Tuple
from datetime import datetime
from dataclasses import dataclass

from .models import (
    Comment, DocumentStructure, Task, TaskPlan, TaskType, RiskLevel, 
    LocatorType, Locator, ValidationResult
)
from .prompt_builder import PromptBuilder, PromptContext
from .llm_client import LLMClient, ModelType
from .constants import FORMAT_TYPES, ALLOWED_WITHOUT_COMMENT
from .exceptions import LLMError, ValidationError, FormatProtectionError
from .utils import load_json_schema


logger = logging.getLogger(__name__)


@dataclass
class PlanningResult:
    """规划结果数据类"""
    success: bool
    tasks: List[Task] = None  # 添加tasks属性
    task_plan: Optional[TaskPlan] = None
    original_tasks: List[Dict[str, Any]] = None
    filtered_tasks: List[Dict[str, Any]] = None
    skipped_tasks: List[Dict[str, Any]] = None
    validation_errors: List[str] = None
    planning_time: float = 0.0
    llm_response_time: float = 0.0
    error_message: Optional[str] = None


class FormatProtectionGuard:
    """格式保护守卫 - 实现四重防线的第2-4层"""
    
    def __init__(self):
        """初始化格式保护守卫"""
        self.format_types = FORMAT_TYPES
        self.allowed_without_comment = ALLOWED_WITHOUT_COMMENT
    
    def validate_task_authorization(self, task_data: Dict[str, Any]) -> Tuple[bool, str]:
        """
        验证任务授权 - 第2层防线：规划期过滤
        
        Args:
            task_data: 任务数据字典
            
        Returns:
            (是否通过验证, 错误消息)
        """
        task_type = task_data.get("type")
        source_comment_id = task_data.get("source_comment_id")
        
        # 检查格式化任务是否有批注授权
        if task_type in self.format_types:
            if not source_comment_id:
                return False, f"格式化任务 '{task_type}' 缺少批注授权 (source_comment_id)"
        
        # 检查非白名单任务是否有批注授权
        elif task_type not in self.allowed_without_comment:
            if not source_comment_id:
                return False, f"任务 '{task_type}' 不在白名单中且缺少批注授权"
        
        return True, ""
    
    def filter_unauthorized_tasks(self, tasks: List[Dict[str, Any]]) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
        """
        过滤未授权的任务
        
        Args:
            tasks: 原始任务列表
            
        Returns:
            (授权任务列表, 被过滤任务列表)
        """
        authorized_tasks = []
        filtered_tasks = []
        
        for task in tasks:
            is_authorized, error_msg = self.validate_task_authorization(task)
            
            if is_authorized:
                authorized_tasks.append(task)
            else:
                # 标记被过滤的原因
                filtered_task = task.copy()
                filtered_task["filtered"] = True
                filtered_task["filter_reason"] = error_msg
                filtered_tasks.append(filtered_task)
                
                logger.warning(f"过滤未授权任务: {task.get('id', 'unknown')} - {error_msg}")
        
        logger.info(f"格式保护过滤: {len(authorized_tasks)} 个授权任务, {len(filtered_tasks)} 个被过滤")
        return authorized_tasks, filtered_tasks
    
    def validate_task_safety(self, task: Task, comments: List[Comment]) -> ValidationResult:
        """
        验证任务安全性 - 第3层防线：执行期拦截
        
        Args:
            task: 任务对象
            comments: 批注列表
            
        Returns:
            验证结果
        """
        errors = []
        warnings = []
        
        # 检查格式化任务的批注授权
        if task.type.value in self.format_types:
            if not task.source_comment_id:
                errors.append(f"格式化任务 {task.id} 缺少批注授权")
            else:
                # 验证批注是否存在
                comment_exists = any(c.id == task.source_comment_id for c in comments)
                if not comment_exists:
                    errors.append(f"任务 {task.id} 引用的批注 {task.source_comment_id} 不存在")
        
        # 检查高风险任务
        if task.risk == RiskLevel.HIGH:
            warnings.append(f"高风险任务 {task.id} 需要特别注意")
        
        # 检查定位器有效性
        if not task.locator.value.strip():
            errors.append(f"任务 {task.id} 的定位器值为空")
        
        return ValidationResult(
            is_valid=len(errors) == 0,
            errors=errors,
            warnings=warnings,
            details={"task_id": task.id, "task_type": task.type.value}
        )


class TaskDependencyResolver:
    """任务依赖解析器"""
    
    def __init__(self):
        """初始化依赖解析器"""
        pass
    
    def resolve_dependencies(self, tasks: List[Task]) -> List[Task]:
        """
        解析任务依赖关系并排序
        
        Args:
            tasks: 任务列表
            
        Returns:
            按依赖关系排序的任务列表
        """
        # 创建任务映射
        task_map = {task.id: task for task in tasks}
        
        # 验证依赖关系
        self._validate_dependencies(tasks, task_map)
        
        # 拓扑排序
        sorted_tasks = self._topological_sort(tasks, task_map)
        
        logger.info(f"依赖解析完成: {len(sorted_tasks)} 个任务已排序")
        return sorted_tasks
    
    def _validate_dependencies(self, tasks: List[Task], task_map: Dict[str, Task]):
        """验证依赖关系的有效性"""
        for task in tasks:
            for dep_id in task.dependencies:
                if dep_id not in task_map:
                    logger.warning(f"任务 {task.id} 依赖的任务 {dep_id} 不存在")
    
    def _topological_sort(self, tasks: List[Task], task_map: Dict[str, Task]) -> List[Task]:
        """拓扑排序算法"""
        # 计算入度
        in_degree = {task.id: 0 for task in tasks}
        for task in tasks:
            for dep_id in task.dependencies:
                if dep_id in in_degree:
                    in_degree[task.id] += 1
        
        # 队列初始化（入度为0的任务）
        queue = [task for task in tasks if in_degree[task.id] == 0]
        result = []
        
        while queue:
            # 按风险级别和类型排序
            queue.sort(key=lambda t: (t.risk.value, t.type.value))
            current = queue.pop(0)
            result.append(current)
            
            # 更新依赖此任务的其他任务的入度
            for task in tasks:
                if current.id in task.dependencies:
                    in_degree[task.id] -= 1
                    if in_degree[task.id] == 0:
                        queue.append(task)
        
        # 检查是否有循环依赖
        if len(result) != len(tasks):
            remaining_tasks = [t for t in tasks if t not in result]
            logger.error(f"检测到循环依赖: {[t.id for t in remaining_tasks]}")
            # 将剩余任务添加到结果中（打破循环）
            result.extend(remaining_tasks)
        
        return result


class RiskAssessment:
    """风险评估器"""
    
    def __init__(self):
        """初始化风险评估器"""
        self.high_risk_types = {
            TaskType.APPLY_TEMPLATE,
            TaskType.REBUILD_TOC,
            TaskType.UPDATE_TOC_LEVELS
        }
        
        self.medium_risk_types = {
            TaskType.SET_HEADING_LEVEL,
            TaskType.REPLACE_HYPERLINK,
            TaskType.SET_PARAGRAPH_STYLE
        }
    
    def assess_task_risk(self, task_data: Dict[str, Any]) -> RiskLevel:
        """
        评估任务风险级别
        
        Args:
            task_data: 任务数据
            
        Returns:
            风险级别
        """
        task_type_str = task_data.get("type", "")
        
        try:
            task_type = TaskType(task_type_str)
        except ValueError:
            logger.warning(f"未知任务类型: {task_type_str}")
            return RiskLevel.MEDIUM
        
        # 基于任务类型的风险评估
        if task_type in self.high_risk_types:
            return RiskLevel.HIGH
        elif task_type in self.medium_risk_types:
            return RiskLevel.MEDIUM
        else:
            return RiskLevel.LOW
    
    def assess_batch_risk(self, tasks: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        评估批量任务的整体风险
        
        Args:
            tasks: 任务列表
            
        Returns:
            风险评估报告
        """
        risk_counts = {
            RiskLevel.LOW: 0,
            RiskLevel.MEDIUM: 0,
            RiskLevel.HIGH: 0
        }
        
        for task in tasks:
            risk = self.assess_task_risk(task)
            risk_counts[risk] += 1
        
        total_tasks = len(tasks)
        high_risk_ratio = risk_counts[RiskLevel.HIGH] / total_tasks if total_tasks > 0 else 0
        
        overall_risk = RiskLevel.LOW
        if high_risk_ratio > 0.3:  # 超过30%高风险任务
            overall_risk = RiskLevel.HIGH
        elif high_risk_ratio > 0.1 or risk_counts[RiskLevel.MEDIUM] > total_tasks * 0.5:
            overall_risk = RiskLevel.MEDIUM
        
        return {
            "overall_risk": overall_risk,
            "risk_distribution": {
                "low": risk_counts[RiskLevel.LOW],
                "medium": risk_counts[RiskLevel.MEDIUM], 
                "high": risk_counts[RiskLevel.HIGH]
            },
            "high_risk_ratio": high_risk_ratio,
            "total_tasks": total_tasks,
            "recommendations": self._generate_risk_recommendations(risk_counts, high_risk_ratio)
        }
    
    def _generate_risk_recommendations(self, risk_counts: Dict[RiskLevel, int], high_risk_ratio: float) -> List[str]:
        """生成风险建议"""
        recommendations = []
        
        if high_risk_ratio > 0.3:
            recommendations.append("建议分批执行，优先处理低风险任务")
            recommendations.append("高风险任务建议启用 dry-run 模式")
        
        if risk_counts[RiskLevel.HIGH] > 0:
            recommendations.append("建议在执行前创建额外备份")
            recommendations.append("高风险任务执行后立即验证结果")
        
        if risk_counts[RiskLevel.MEDIUM] > 10:
            recommendations.append("中等风险任务较多，建议分组执行")
        
        return recommendations


class TaskPlanner:
    """任务规划器 - 核心规划逻辑"""
    
    def __init__(self, 
                 schema_path: str = "schemas/tasks.schema.json",
                 default_model: ModelType = ModelType.GPT5):
        """
        初始化任务规划器
        
        Args:
            schema_path: JSON Schema 文件路径
            default_model: 默认使用的 LLM 模型
        """
        self.schema_path = schema_path
        self.default_model = default_model
        
        # 初始化组件
        self.prompt_builder = PromptBuilder(schema_path)
        self.llm_client = LLMClient()
        self.format_guard = FormatProtectionGuard()
        self.dependency_resolver = TaskDependencyResolver()
        self.risk_assessor = RiskAssessment()
    
    def generate_plan(self, 
                     document_structure: DocumentStructure,
                     comments: List[Comment],
                     document_path: Optional[str] = None,
                     model: Optional[ModelType] = None) -> PlanningResult:
        """
        生成任务计划
        
        Args:
            document_structure: 文档结构
            comments: 批注列表
            document_path: 文档路径
            model: 使用的 LLM 模型
            
        Returns:
            规划结果
        """
        start_time = datetime.now()
        model = model or self.default_model
        
        try:
            logger.info(f"开始生成任务计划: {len(comments)} 个批注")
            
            # 1. 构建提示词上下文
            context = PromptContext(
                document_structure=document_structure,
                comments=comments,
                document_path=document_path
            )
            
            # 2. 检查上下文长度
            length_check = self.prompt_builder.check_context_length(context)
            if not length_check["is_within_limit"]:
                logger.warning(f"上下文超限: {length_check['overflow_tokens']} tokens")
                return self._handle_context_overflow(context, model, start_time)
            
            # 3. 调用 LLM 生成任务
            llm_start_time = datetime.now()
            raw_response = self._call_llm(context, model)
            llm_end_time = datetime.now()
            llm_response_time = (llm_end_time - llm_start_time).total_seconds()
            
            # 4. 解析和验证 LLM 响应
            tasks_data = self._parse_llm_response(raw_response)
            
            # 5. 应用格式保护 - 第2层防线
            authorized_tasks, filtered_tasks = self.format_guard.filter_unauthorized_tasks(tasks_data)
            
            # 6. 风险评估
            risk_report = self.risk_assessor.assess_batch_risk(authorized_tasks)
            logger.info(f"风险评估: {risk_report['overall_risk'].value} 级别")
            
            # 7. 转换为任务对象
            task_objects = self._convert_to_task_objects(authorized_tasks, comments)
            
            # 8. 依赖解析和排序
            sorted_tasks = self.dependency_resolver.resolve_dependencies(task_objects)
            
            # 9. 创建任务计划
            task_plan = TaskPlan(
                tasks=sorted_tasks,
                document_path=document_path,
                total_tasks=len(sorted_tasks)
            )
            
            end_time = datetime.now()
            planning_time = (end_time - start_time).total_seconds()
            
            logger.info(f"任务计划生成完成: {len(sorted_tasks)} 个任务, 耗时 {planning_time:.2f}s")
            
            return PlanningResult(
                success=True,
                tasks=sorted_tasks,  # 添加tasks字段
                task_plan=task_plan,
                original_tasks=tasks_data,
                filtered_tasks=authorized_tasks,
                skipped_tasks=filtered_tasks,
                planning_time=planning_time,
                llm_response_time=llm_response_time
            )
            
        except Exception as e:
            end_time = datetime.now()
            planning_time = (end_time - start_time).total_seconds()
            
            logger.error(f"任务计划生成失败: {e}")
            
            return PlanningResult(
                success=False,
                tasks=[],  # 添加空的tasks列表
                error_message=str(e),
                planning_time=planning_time
            )
    
    def _call_llm(self, context: PromptContext, model: ModelType) -> str:
        """调用 LLM 生成任务"""
        system_prompt = self.prompt_builder.build_system_prompt()
        user_prompt = self.prompt_builder.build_user_prompt(context)
        
        try:
            # 使用 JSON 重试机制确保获得有效响应
            response = self.llm_client.call_with_json_retry(
                model, system_prompt, user_prompt, max_json_retries=2
            )
            
            if not response.success:
                raise LLMError(f"LLM 调用失败: {response.error}")
            
            return response.content
            
        except Exception as e:
            raise LLMError(f"LLM 调用失败: {e}")
    
    def _parse_llm_response(self, raw_response: str) -> List[Dict[str, Any]]:
        """解析 LLM 响应"""
        try:
            response_data = json.loads(raw_response)
            
            if "tasks" not in response_data:
                raise ValidationError("LLM 响应缺少 'tasks' 字段")
            
            tasks = response_data["tasks"]
            if not isinstance(tasks, list):
                raise ValidationError("'tasks' 字段必须是数组")
            
            logger.info(f"解析 LLM 响应: {len(tasks)} 个任务")
            return tasks
            
        except json.JSONDecodeError as e:
            raise ValidationError(f"LLM 响应不是有效的 JSON: {e}")
    
    def _convert_to_task_objects(self, tasks_data: List[Dict[str, Any]], comments: List[Comment]) -> List[Task]:
        """将任务数据转换为任务对象"""
        task_objects = []
        
        for i, task_data in enumerate(tasks_data):
            try:
                # 评估风险级别
                risk = self.risk_assessor.assess_task_risk(task_data)
                if "risk" not in task_data:
                    task_data["risk"] = risk.value
                
                # 创建定位器
                locator_data = task_data.get("locator", {})
                locator = Locator(
                    by=LocatorType(locator_data.get("by", "find")),
                    value=locator_data.get("value", "")
                )
                
                # 创建任务对象
                task = Task(
                    id=task_data.get("id", f"task_{i+1}"),
                    type=TaskType(task_data["type"]),
                    source_comment_id=task_data.get("source_comment_id"),
                    locator=locator,
                    instruction=task_data["instruction"],
                    dependencies=task_data.get("dependencies", []),
                    risk=RiskLevel(task_data.get("risk", "low")),
                    checks=task_data.get("checks", []),
                    requires_user_review=task_data.get("requires_user_review", False),
                    notes=task_data.get("notes")
                )
                
                task_objects.append(task)
                
            except Exception as e:
                logger.error(f"转换任务对象失败 (任务 {i+1}): {e}")
                continue
        
        return task_objects
    
    def _handle_context_overflow(self, context: PromptContext, model: ModelType, start_time: datetime) -> PlanningResult:
        """处理上下文溢出"""
        try:
            logger.info("处理上下文溢出，启用分块处理")
            
            # 分块处理
            chunks = self.prompt_builder.handle_context_overflow(context)
            all_tasks = []
            
            for i, chunk in enumerate(chunks, 1):
                logger.info(f"处理块 {i}/{len(chunks)}")
                
                raw_response = self._call_llm(chunk, model)
                chunk_tasks = self._parse_llm_response(raw_response)
                all_tasks.extend(chunk_tasks)
            
            # 合并结果并继续正常流程
            authorized_tasks, filtered_tasks = self.format_guard.filter_unauthorized_tasks(all_tasks)
            task_objects = self._convert_to_task_objects(authorized_tasks, context.comments)
            sorted_tasks = self.dependency_resolver.resolve_dependencies(task_objects)
            
            task_plan = TaskPlan(
                tasks=sorted_tasks,
                document_path=context.document_path,
                total_tasks=len(sorted_tasks)
            )
            
            end_time = datetime.now()
            planning_time = (end_time - start_time).total_seconds()
            
            logger.info(f"分块处理完成: {len(chunks)} 个块, {len(sorted_tasks)} 个任务")
            
            return PlanningResult(
                success=True,
                tasks=sorted_tasks,  # 添加tasks字段
                task_plan=task_plan,
                original_tasks=all_tasks,
                filtered_tasks=authorized_tasks,
                skipped_tasks=filtered_tasks,
                planning_time=planning_time
            )
            
        except Exception as e:
            end_time = datetime.now()
            planning_time = (end_time - start_time).total_seconds()
            
            return PlanningResult(
                success=False,
                tasks=[],  # 添加空的tasks列表
                error_message=f"分块处理失败: {e}",
                planning_time=planning_time
            )
    
    def validate_plan(self, task_plan: TaskPlan, comments: List[Comment]) -> ValidationResult:
        """
        验证任务计划 - 第3层防线：执行期拦截
        
        Args:
            task_plan: 任务计划
            comments: 批注列表
            
        Returns:
            验证结果
        """
        all_errors = []
        all_warnings = []
        
        for task in task_plan.tasks:
            result = self.format_guard.validate_task_safety(task, comments)
            all_errors.extend(result.errors)
            all_warnings.extend(result.warnings)
        
        return ValidationResult(
            is_valid=len(all_errors) == 0,
            errors=all_errors,
            warnings=all_warnings,
            details={
                "total_tasks": len(task_plan.tasks),
                "validation_time": datetime.now().isoformat()
            }
        )
    
    def close(self):
        """关闭规划器资源"""
        if hasattr(self, 'llm_client'):
            self.llm_client.close()


# 便捷函数
def create_task_plan(document_structure: DocumentStructure,
                    comments: List[Comment],
                    document_path: Optional[str] = None,
                    model: Optional[ModelType] = None) -> PlanningResult:
    """
    便捷函数：创建任务计划
    
    Args:
        document_structure: 文档结构
        comments: 批注列表
        document_path: 文档路径
        model: LLM 模型
        
    Returns:
        规划结果
    """
    planner = TaskPlanner()
    try:
        return planner.generate_plan(document_structure, comments, document_path, model)
    finally:
        planner.close()