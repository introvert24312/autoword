"""
AutoWord Error Handler
综合错误处理和用户反馈系统
"""

import logging
import traceback
from typing import Dict, List, Optional, Any, Tuple
from dataclasses import dataclass
from enum import Enum

from .exceptions import (
    AutoWordError, DocumentError, COMError, LLMError, 
    ValidationError, FormatProtectionError, TaskExecutionError,
    ConfigurationError
)


logger = logging.getLogger(__name__)


class ErrorSeverity(str, Enum):
    """错误严重程度"""
    INFO = "info"
    WARNING = "warning"
    ERROR = "error"
    CRITICAL = "critical"


class ErrorCategory(str, Enum):
    """错误类别"""
    DOCUMENT = "document"
    COM = "com"
    LLM = "llm"
    VALIDATION = "validation"
    FORMAT_PROTECTION = "format_protection"
    TASK_EXECUTION = "task_execution"
    CONFIGURATION = "configuration"
    SYSTEM = "system"
    UNKNOWN = "unknown"


@dataclass
class ErrorInfo:
    """错误信息"""
    category: ErrorCategory
    severity: ErrorSeverity
    message: str
    user_message: str
    suggestions: List[str]
    technical_details: Optional[str] = None
    error_code: Optional[str] = None
    recovery_actions: List[str] = None
    
    def __post_init__(self):
        if self.recovery_actions is None:
            self.recovery_actions = []


class ErrorHandler:
    """错误处理器"""
    
    def __init__(self):
        """初始化错误处理器"""
        self.error_mappings = self._build_error_mappings()
    
    def handle_error(self, error: Exception, context: Optional[str] = None) -> ErrorInfo:
        """
        处理错误并生成用户友好的错误信息
        
        Args:
            error: 异常对象
            context: 错误上下文
            
        Returns:
            错误信息对象
        """
        error_type = type(error)
        error_message = str(error)
        
        # 获取错误映射
        if error_type in self.error_mappings:
            error_info = self.error_mappings[error_type](error, context)
        else:
            error_info = self._handle_unknown_error(error, context)
        
        # 添加技术详细信息
        if logger.isEnabledFor(logging.DEBUG):
            error_info.technical_details = traceback.format_exc()
        
        # 记录错误
        self._log_error(error_info, error)
        
        return error_info
    
    def _build_error_mappings(self) -> Dict[type, callable]:
        """构建错误映射"""
        return {
            DocumentError: self._handle_document_error,
            COMError: self._handle_com_error,
            LLMError: self._handle_llm_error,
            ValidationError: self._handle_validation_error,
            FormatProtectionError: self._handle_format_protection_error,
            TaskExecutionError: self._handle_task_execution_error,
            ConfigurationError: self._handle_configuration_error,
            FileNotFoundError: self._handle_file_not_found_error,
            PermissionError: self._handle_permission_error,
            ConnectionError: self._handle_connection_error,
            TimeoutError: self._handle_timeout_error,
        }
    
    def _handle_document_error(self, error: DocumentError, context: Optional[str]) -> ErrorInfo:
        """处理文档错误"""
        message = str(error)
        
        if "找不到" in message or "不存在" in message:
            return ErrorInfo(
                category=ErrorCategory.DOCUMENT,
                severity=ErrorSeverity.ERROR,
                message=message,
                user_message="无法找到指定的文档文件",
                suggestions=[
                    "检查文档路径是否正确",
                    "确认文档文件是否存在",
                    "检查文件是否被其他程序占用"
                ],
                error_code="DOC_001",
                recovery_actions=[
                    "选择其他文档文件",
                    "检查文件权限",
                    "关闭可能占用文件的程序"
                ]
            )
        elif "权限" in message or "访问" in message:
            return ErrorInfo(
                category=ErrorCategory.DOCUMENT,
                severity=ErrorSeverity.ERROR,
                message=message,
                user_message="没有足够的权限访问文档",
                suggestions=[
                    "以管理员身份运行程序",
                    "检查文件权限设置",
                    "确认文档未被保护"
                ],
                error_code="DOC_002",
                recovery_actions=[
                    "右键选择"以管理员身份运行"",
                    "修改文件权限",
                    "联系文档所有者"
                ]
            )
        else:
            return ErrorInfo(
                category=ErrorCategory.DOCUMENT,
                severity=ErrorSeverity.ERROR,
                message=message,
                user_message="文档处理时发生错误",
                suggestions=[
                    "检查文档格式是否正确",
                    "尝试用 Word 打开文档验证",
                    "检查文档是否损坏"
                ],
                error_code="DOC_003"
            )
    
    def _handle_com_error(self, error: COMError, context: Optional[str]) -> ErrorInfo:
        """处理 COM 错误"""
        message = str(error)
        
        if "Word" in message and ("未安装" in message or "找不到" in message):
            return ErrorInfo(
                category=ErrorCategory.COM,
                severity=ErrorSeverity.CRITICAL,
                message=message,
                user_message="系统中未安装 Microsoft Word 或版本不兼容",
                suggestions=[
                    "安装 Microsoft Word 2016 或更高版本",
                    "检查 Word 是否正确安装",
                    "重新注册 Word COM 组件"
                ],
                error_code="COM_001",
                recovery_actions=[
                    "安装 Microsoft Office",
                    "运行 Office 修复工具",
                    "重新启动计算机"
                ]
            )
        elif "COM" in message and "初始化" in message:
            return ErrorInfo(
                category=ErrorCategory.COM,
                severity=ErrorSeverity.ERROR,
                message=message,
                user_message="Word COM 组件初始化失败",
                suggestions=[
                    "关闭所有 Word 进程后重试",
                    "重新启动计算机",
                    "检查系统权限"
                ],
                error_code="COM_002",
                recovery_actions=[
                    "结束 Word 进程",
                    "重启应用程序",
                    "以管理员身份运行"
                ]
            )
        else:
            return ErrorInfo(
                category=ErrorCategory.COM,
                severity=ErrorSeverity.ERROR,
                message=message,
                user_message="Word 组件操作失败",
                suggestions=[
                    "检查 Word 是否正常工作",
                    "尝试手动打开文档",
                    "重新启动 Word"
                ],
                error_code="COM_003"
            )
    
    def _handle_llm_error(self, error: LLMError, context: Optional[str]) -> ErrorInfo:
        """处理 LLM 错误"""
        message = str(error)
        
        if "API" in message and ("密钥" in message or "key" in message.lower()):
            return ErrorInfo(
                category=ErrorCategory.LLM,
                severity=ErrorSeverity.ERROR,
                message=message,
                user_message="LLM API 密钥配置错误",
                suggestions=[
                    "检查环境变量 GPT5_KEY 或 CLAUDE37_KEY",
                    "确认 API 密钥格式正确",
                    "验证 API 密钥是否有效"
                ],
                error_code="LLM_001",
                recovery_actions=[
                    "重新设置环境变量",
                    "联系 API 提供商验证密钥",
                    "尝试使用其他模型"
                ]
            )
        elif "网络" in message or "连接" in message or "timeout" in message.lower():
            return ErrorInfo(
                category=ErrorCategory.LLM,
                severity=ErrorSeverity.WARNING,
                message=message,
                user_message="网络连接问题导致 LLM 服务不可用",
                suggestions=[
                    "检查网络连接",
                    "稍后重试",
                    "检查防火墙设置"
                ],
                error_code="LLM_002",
                recovery_actions=[
                    "检查网络连接",
                    "等待网络恢复后重试",
                    "联系网络管理员"
                ]
            )
        elif "JSON" in message:
            return ErrorInfo(
                category=ErrorCategory.LLM,
                severity=ErrorSeverity.WARNING,
                message=message,
                user_message="LLM 响应格式错误，正在重试",
                suggestions=[
                    "系统会自动重试",
                    "如果持续失败，请尝试其他模型"
                ],
                error_code="LLM_003",
                recovery_actions=[
                    "等待自动重试完成",
                    "切换到其他 LLM 模型"
                ]
            )
        else:
            return ErrorInfo(
                category=ErrorCategory.LLM,
                severity=ErrorSeverity.ERROR,
                message=message,
                user_message="LLM 服务调用失败",
                suggestions=[
                    "检查 API 配置",
                    "稍后重试",
                    "尝试其他模型"
                ],
                error_code="LLM_004"
            )
    
    def _handle_validation_error(self, error: ValidationError, context: Optional[str]) -> ErrorInfo:
        """处理验证错误"""
        return ErrorInfo(
            category=ErrorCategory.VALIDATION,
            severity=ErrorSeverity.WARNING,
            message=str(error),
            user_message="数据验证失败",
            suggestions=[
                "检查输入数据格式",
                "确认数据完整性",
                "查看详细错误信息"
            ],
            error_code="VAL_001"
        )
    
    def _handle_format_protection_error(self, error: FormatProtectionError, context: Optional[str]) -> ErrorInfo:
        """处理格式保护错误"""
        return ErrorInfo(
            category=ErrorCategory.FORMAT_PROTECTION,
            severity=ErrorSeverity.WARNING,
            message=str(error),
            user_message="格式保护系统阻止了未授权的格式变更",
            suggestions=[
                "确认批注中明确要求格式变更",
                "检查任务是否有对应的批注授权",
                "这是安全保护机制，防止意外的格式变更"
            ],
            error_code="FMT_001",
            recovery_actions=[
                "在文档中添加明确的格式变更批注",
                "检查批注ID是否正确",
                "确认格式变更确实是必要的"
            ]
        )
    
    def _handle_task_execution_error(self, error: TaskExecutionError, context: Optional[str]) -> ErrorInfo:
        """处理任务执行错误"""
        message = str(error)
        
        if "定位" in message:
            return ErrorInfo(
                category=ErrorCategory.TASK_EXECUTION,
                severity=ErrorSeverity.WARNING,
                message=message,
                user_message="无法在文档中找到指定的位置",
                suggestions=[
                    "检查文档内容是否已变更",
                    "确认定位文本是否准确",
                    "尝试使用更精确的定位方式"
                ],
                error_code="TASK_001",
                recovery_actions=[
                    "手动检查文档内容",
                    "更新定位信息",
                    "跳过此任务继续执行其他任务"
                ]
            )
        else:
            return ErrorInfo(
                category=ErrorCategory.TASK_EXECUTION,
                severity=ErrorSeverity.ERROR,
                message=message,
                user_message="任务执行过程中发生错误",
                suggestions=[
                    "检查任务参数是否正确",
                    "确认文档状态正常",
                    "查看详细错误信息"
                ],
                error_code="TASK_002"
            )
    
    def _handle_configuration_error(self, error: ConfigurationError, context: Optional[str]) -> ErrorInfo:
        """处理配置错误"""
        return ErrorInfo(
            category=ErrorCategory.CONFIGURATION,
            severity=ErrorSeverity.ERROR,
            message=str(error),
            user_message="系统配置错误",
            suggestions=[
                "检查配置文件",
                "确认环境变量设置",
                "重新安装应用程序"
            ],
            error_code="CFG_001",
            recovery_actions=[
                "重置配置到默认值",
                "重新配置系统",
                "联系技术支持"
            ]
        )
    
    def _handle_file_not_found_error(self, error: FileNotFoundError, context: Optional[str]) -> ErrorInfo:
        """处理文件未找到错误"""
        return ErrorInfo(
            category=ErrorCategory.SYSTEM,
            severity=ErrorSeverity.ERROR,
            message=str(error),
            user_message="找不到指定的文件",
            suggestions=[
                "检查文件路径是否正确",
                "确认文件是否存在",
                "检查文件名拼写"
            ],
            error_code="SYS_001",
            recovery_actions=[
                "浏览并选择正确的文件",
                "检查文件是否被移动或删除",
                "从备份中恢复文件"
            ]
        )
    
    def _handle_permission_error(self, error: PermissionError, context: Optional[str]) -> ErrorInfo:
        """处理权限错误"""
        return ErrorInfo(
            category=ErrorCategory.SYSTEM,
            severity=ErrorSeverity.ERROR,
            message=str(error),
            user_message="没有足够的权限执行此操作",
            suggestions=[
                "以管理员身份运行程序",
                "检查文件权限",
                "确认用户权限"
            ],
            error_code="SYS_002",
            recovery_actions=[
                "右键选择"以管理员身份运行"",
                "修改文件权限",
                "联系系统管理员"
            ]
        )
    
    def _handle_connection_error(self, error: ConnectionError, context: Optional[str]) -> ErrorInfo:
        """处理连接错误"""
        return ErrorInfo(
            category=ErrorCategory.SYSTEM,
            severity=ErrorSeverity.WARNING,
            message=str(error),
            user_message="网络连接失败",
            suggestions=[
                "检查网络连接",
                "稍后重试",
                "检查防火墙设置"
            ],
            error_code="SYS_003",
            recovery_actions=[
                "检查网络连接",
                "重启网络适配器",
                "联系网络管理员"
            ]
        )
    
    def _handle_timeout_error(self, error: TimeoutError, context: Optional[str]) -> ErrorInfo:
        """处理超时错误"""
        return ErrorInfo(
            category=ErrorCategory.SYSTEM,
            severity=ErrorSeverity.WARNING,
            message=str(error),
            user_message="操作超时",
            suggestions=[
                "稍后重试",
                "检查网络连接",
                "增加超时时间"
            ],
            error_code="SYS_004",
            recovery_actions=[
                "等待一段时间后重试",
                "检查系统性能",
                "优化网络连接"
            ]
        )
    
    def _handle_unknown_error(self, error: Exception, context: Optional[str]) -> ErrorInfo:
        """处理未知错误"""
        return ErrorInfo(
            category=ErrorCategory.UNKNOWN,
            severity=ErrorSeverity.ERROR,
            message=str(error),
            user_message="发生了未知错误",
            suggestions=[
                "重新启动应用程序",
                "检查系统状态",
                "联系技术支持"
            ],
            error_code="UNK_001",
            recovery_actions=[
                "重启应用程序",
                "检查系统日志",
                "报告错误给开发团队"
            ]
        )
    
    def _log_error(self, error_info: ErrorInfo, original_error: Exception):
        """记录错误"""
        log_level = {
            ErrorSeverity.INFO: logging.INFO,
            ErrorSeverity.WARNING: logging.WARNING,
            ErrorSeverity.ERROR: logging.ERROR,
            ErrorSeverity.CRITICAL: logging.CRITICAL
        }.get(error_info.severity, logging.ERROR)
        
        logger.log(
            log_level,
            f"[{error_info.error_code}] {error_info.category.value.upper()}: {error_info.message}",
            exc_info=original_error if log_level >= logging.ERROR else None
        )
    
    def get_recovery_suggestions(self, error_category: ErrorCategory) -> List[str]:
        """
        获取特定错误类别的通用恢复建议
        
        Args:
            error_category: 错误类别
            
        Returns:
            恢复建议列表
        """
        suggestions = {
            ErrorCategory.DOCUMENT: [
                "检查文档是否存在且可访问",
                "确认文档格式正确",
                "尝试用 Word 手动打开文档"
            ],
            ErrorCategory.COM: [
                "确认 Microsoft Word 已正确安装",
                "重启应用程序",
                "以管理员身份运行"
            ],
            ErrorCategory.LLM: [
                "检查网络连接",
                "验证 API 密钥配置",
                "稍后重试"
            ],
            ErrorCategory.VALIDATION: [
                "检查输入数据格式",
                "确认数据完整性"
            ],
            ErrorCategory.FORMAT_PROTECTION: [
                "确认批注中有明确的格式变更要求",
                "这是安全保护机制"
            ],
            ErrorCategory.TASK_EXECUTION: [
                "检查任务参数",
                "确认文档状态正常"
            ],
            ErrorCategory.CONFIGURATION: [
                "检查配置文件",
                "重置到默认配置"
            ],
            ErrorCategory.SYSTEM: [
                "检查系统权限",
                "重启应用程序"
            ]
        }
        
        return suggestions.get(error_category, ["联系技术支持"])


# 全局错误处理器实例
error_handler = ErrorHandler()


def handle_error(error: Exception, context: Optional[str] = None) -> ErrorInfo:
    """
    便捷函数：处理错误
    
    Args:
        error: 异常对象
        context: 错误上下文
        
    Returns:
        错误信息对象
    """
    return error_handler.handle_error(error, context)