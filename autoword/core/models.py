"""
AutoWord Core Data Models
核心数据模型定义
"""

from dataclasses import dataclass
from typing import List, Optional, Dict, Any, Union
from enum import Enum
from datetime import datetime
from pydantic import BaseModel, Field, field_validator


class TaskType(str, Enum):
    """任务类型枚举"""
    REWRITE = "rewrite"
    INSERT = "insert"
    DELETE = "delete"
    SET_PARAGRAPH_STYLE = "set_paragraph_style"
    SET_HEADING_LEVEL = "set_heading_level"
    APPLY_TEMPLATE = "apply_template"
    REPLACE_HYPERLINK = "replace_hyperlink"
    REBUILD_TOC = "rebuild_toc"
    UPDATE_TOC_LEVELS = "update_toc_levels"
    REFRESH_TOC_NUMBERS = "refresh_toc_numbers"


class RiskLevel(str, Enum):
    """风险等级枚举"""
    LOW = "low"
    MEDIUM = "medium"
    HIGH = "high"


class LocatorType(str, Enum):
    """定位器类型枚举"""
    BOOKMARK = "bookmark"
    RANGE = "range"
    HEADING = "heading"
    FIND = "find"


class Comment(BaseModel):
    """文档批注模型"""
    id: str = Field(..., description="批注唯一标识")
    author: str = Field(..., description="批注作者")
    page: int = Field(..., description="批注所在页码")
    anchor_text: str = Field(..., description="批注锚点文本")
    comment_text: str = Field(..., description="批注内容")
    range_start: int = Field(..., description="批注范围起始位置")
    range_end: int = Field(..., description="批注范围结束位置")
    created_date: Optional[datetime] = Field(None, description="创建时间")
    
    @field_validator('page')
    @classmethod
    def page_must_be_positive(cls, v):
        if v <= 0:
            raise ValueError('页码必须为正数')
        return v


class Locator(BaseModel):
    """任务定位器模型"""
    by: LocatorType = Field(..., description="定位方式")
    value: str = Field(..., description="定位值")
    
    @field_validator('value')
    @classmethod
    def value_must_not_be_empty(cls, v):
        if not v.strip():
            raise ValueError('定位值不能为空')
        return v.strip()


class Task(BaseModel):
    """任务模型"""
    id: str = Field(..., description="任务唯一标识")
    type: TaskType = Field(..., description="任务类型")
    source_comment_id: Optional[str] = Field(None, description="来源批注ID")
    locator: Locator = Field(..., description="任务定位器")
    instruction: str = Field(..., description="任务指令")
    dependencies: List[str] = Field(default_factory=list, description="依赖任务ID列表")
    risk: RiskLevel = Field(default=RiskLevel.LOW, description="风险等级")
    checks: List[str] = Field(default_factory=list, description="检查项列表")
    requires_user_review: bool = Field(default=False, description="是否需要用户审核")
    notes: Optional[str] = Field(None, description="备注信息")
    
    @field_validator('instruction')
    @classmethod
    def instruction_must_not_be_empty(cls, v):
        if not v.strip():
            raise ValueError('任务指令不能为空')
        return v.strip()


class Heading(BaseModel):
    """标题模型"""
    level: int = Field(..., description="标题级别")
    text: str = Field(..., description="标题文本")
    style: str = Field(..., description="标题样式")
    range_start: int = Field(..., description="标题范围起始位置")
    range_end: int = Field(..., description="标题范围结束位置")
    
    @field_validator('level')
    @classmethod
    def level_must_be_valid(cls, v):
        if v < 1 or v > 9:
            raise ValueError('标题级别必须在1-9之间')
        return v


class Style(BaseModel):
    """样式模型"""
    name: str = Field(..., description="样式名称")
    type: str = Field(..., description="样式类型")
    built_in: bool = Field(..., description="是否为内置样式")
    in_use: bool = Field(..., description="是否正在使用")


class TocEntry(BaseModel):
    """目录条目模型"""
    level: int = Field(..., description="目录级别")
    text: str = Field(..., description="目录文本")
    page_number: int = Field(..., description="页码")
    range_start: int = Field(..., description="范围起始位置")
    range_end: int = Field(..., description="范围结束位置")


class Hyperlink(BaseModel):
    """超链接模型"""
    text: str = Field(..., description="链接文本")
    address: str = Field(..., description="链接地址")
    type: str = Field(..., description="链接类型")
    range_start: int = Field(..., description="范围起始位置")
    range_end: int = Field(..., description="范围结束位置")


class Reference(BaseModel):
    """引用模型"""
    type: str = Field(..., description="引用类型")
    text: str = Field(..., description="引用文本")
    target: str = Field(..., description="引用目标")
    range_start: int = Field(..., description="范围起始位置")
    range_end: int = Field(..., description="范围结束位置")


class DocumentStructure(BaseModel):
    """文档结构模型"""
    headings: List[Heading] = Field(default_factory=list, description="标题列表")
    styles: List[Style] = Field(default_factory=list, description="样式列表")
    toc_entries: List[TocEntry] = Field(default_factory=list, description="目录条目列表")
    hyperlinks: List[Hyperlink] = Field(default_factory=list, description="超链接列表")
    references: List[Reference] = Field(default_factory=list, description="引用列表")
    page_count: int = Field(..., description="总页数")
    word_count: int = Field(..., description="总字数")
    
    @field_validator('page_count', 'word_count')
    @classmethod
    def counts_must_be_non_negative(cls, v):
        if v < 0:
            raise ValueError('计数值不能为负数')
        return v


class TaskPlan(BaseModel):
    """任务计划模型"""
    tasks: List[Task] = Field(..., description="任务列表")
    created_at: datetime = Field(default_factory=datetime.now, description="创建时间")
    document_path: Optional[str] = Field(None, description="文档路径")
    total_tasks: int = Field(..., description="任务总数")
    
    @field_validator('total_tasks')
    @classmethod
    def total_tasks_must_match(cls, v, info):
        if info.data and 'tasks' in info.data and len(info.data['tasks']) != v:
            raise ValueError('任务总数与任务列表长度不匹配')
        return v


class TaskResult(BaseModel):
    """任务执行结果模型"""
    task_id: str = Field(..., description="任务ID")
    success: bool = Field(..., description="是否成功")
    message: str = Field(..., description="执行消息")
    execution_time: float = Field(..., description="执行时间(秒)")
    error_details: Optional[str] = Field(None, description="错误详情")


class ExecutionResult(BaseModel):
    """执行结果模型"""
    success: bool = Field(..., description="整体是否成功")
    total_tasks: int = Field(..., description="任务总数")
    completed_tasks: int = Field(..., description="完成任务数")
    failed_tasks: int = Field(..., description="失败任务数")
    task_results: List[TaskResult] = Field(..., description="任务结果列表")
    execution_time: float = Field(..., description="总执行时间(秒)")
    error_summary: Optional[str] = Field(None, description="错误摘要")


class DocumentSnapshot(BaseModel):
    """文档快照模型"""
    timestamp: datetime = Field(default_factory=datetime.now, description="快照时间")
    document_path: str = Field(..., description="文档路径")
    structure: DocumentStructure = Field(..., description="文档结构")
    comments: List[Comment] = Field(..., description="批注列表")
    checksum: str = Field(..., description="文档校验和")


class ValidationResult(BaseModel):
    """验证结果模型"""
    is_valid: bool = Field(..., description="是否有效")
    errors: List[str] = Field(default_factory=list, description="错误列表")
    warnings: List[str] = Field(default_factory=list, description="警告列表")
    details: Dict[str, Any] = Field(default_factory=dict, description="详细信息")