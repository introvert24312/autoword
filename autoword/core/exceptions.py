"""
AutoWord Custom Exceptions
自定义异常类
"""


class AutoWordError(Exception):
    """AutoWord 基础异常类"""
    pass


class DocumentError(AutoWordError):
    """文档相关异常"""
    pass


class COMError(AutoWordError):
    """Word COM 相关异常"""
    pass


class LLMError(AutoWordError):
    """LLM 服务相关异常"""
    pass


class ValidationError(AutoWordError):
    """数据验证异常"""
    pass


class FormatProtectionError(AutoWordError):
    """格式保护异常"""
    pass


class TaskExecutionError(AutoWordError):
    """任务执行异常"""
    pass


class ConfigurationError(AutoWordError):
    """配置异常"""
    pass

class APIKeyError(AutoWordError):
    """API密钥相关异常"""
    pass