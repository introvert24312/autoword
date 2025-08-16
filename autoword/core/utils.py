"""
AutoWord Utility Functions
工具函数
"""

import os
import json
import hashlib
from datetime import datetime
from typing import Dict, Any, Optional
from pathlib import Path

from .constants import BACKUP_DATE_FORMAT, SUPPORTED_EXTENSIONS
from .exceptions import ValidationError, ConfigurationError


def get_api_key(key_name: str) -> str:
    """
    从环境变量获取 API 密钥
    
    Args:
        key_name: 环境变量名称
        
    Returns:
        API 密钥
        
    Raises:
        ConfigurationError: 如果密钥不存在
    """
    key = os.getenv(key_name)
    if not key:
        raise ConfigurationError(f"Missing {key_name} environment variable")
    return key


def validate_document_path(file_path: str) -> Path:
    """
    验证文档路径
    
    Args:
        file_path: 文档路径
        
    Returns:
        验证后的路径对象
        
    Raises:
        ValidationError: 如果路径无效
    """
    path = Path(file_path)
    
    if not path.exists():
        raise ValidationError(f"文档文件不存在: {file_path}")
    
    if not path.is_file():
        raise ValidationError(f"路径不是文件: {file_path}")
    
    if path.suffix.lower() not in SUPPORTED_EXTENSIONS:
        raise ValidationError(f"不支持的文件类型: {path.suffix}")
    
    return path


def generate_backup_path(original_path: str) -> str:
    """
    生成备份文件路径
    
    Args:
        original_path: 原始文件路径
        
    Returns:
        备份文件路径
    """
    path = Path(original_path)
    timestamp = datetime.now().strftime(BACKUP_DATE_FORMAT)
    backup_name = f"{path.stem}_backup_{timestamp}{path.suffix}"
    return str(path.parent / backup_name)


def calculate_file_checksum(file_path: str) -> str:
    """
    计算文件校验和
    
    Args:
        file_path: 文件路径
        
    Returns:
        MD5 校验和
    """
    hash_md5 = hashlib.md5()
    with open(file_path, "rb") as f:
        for chunk in iter(lambda: f.read(4096), b""):
            hash_md5.update(chunk)
    return hash_md5.hexdigest()


def load_json_schema(schema_path: str) -> Dict[str, Any]:
    """
    加载 JSON Schema
    
    Args:
        schema_path: Schema 文件路径
        
    Returns:
        Schema 字典
        
    Raises:
        ValidationError: 如果 Schema 无效
    """
    try:
        with open(schema_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError) as e:
        raise ValidationError(f"无法加载 JSON Schema: {e}")


def sanitize_filename(filename: str) -> str:
    """
    清理文件名，移除非法字符
    
    Args:
        filename: 原始文件名
        
    Returns:
        清理后的文件名
    """
    # Windows 文件名非法字符
    illegal_chars = '<>:"/\\|?*'
    for char in illegal_chars:
        filename = filename.replace(char, '_')
    
    # 移除前后空格和点
    filename = filename.strip(' .')
    
    # 确保不为空
    if not filename:
        filename = "untitled"
    
    return filename


def format_duration(seconds: float) -> str:
    """
    格式化持续时间
    
    Args:
        seconds: 秒数
        
    Returns:
        格式化的时间字符串
    """
    if seconds < 1:
        return f"{seconds*1000:.0f}ms"
    elif seconds < 60:
        return f"{seconds:.1f}s"
    else:
        minutes = int(seconds // 60)
        remaining_seconds = seconds % 60
        return f"{minutes}m {remaining_seconds:.1f}s"


def truncate_text(text: str, max_length: int = 100, suffix: str = "...") -> str:
    """
    截断文本
    
    Args:
        text: 原始文本
        max_length: 最大长度
        suffix: 截断后缀
        
    Returns:
        截断后的文本
    """
    if len(text) <= max_length:
        return text
    
    return text[:max_length - len(suffix)] + suffix


def ensure_directory(directory: str) -> Path:
    """
    确保目录存在
    
    Args:
        directory: 目录路径
        
    Returns:
        目录路径对象
    """
    path = Path(directory)
    path.mkdir(parents=True, exist_ok=True)
    return path