"""
Configuration Manager
配置管理器 - 负责应用程序配置的安全存储和管理
"""

import os
import json
import logging
from typing import Dict, Any, Optional
from pathlib import Path
from dataclasses import dataclass, field, asdict
from cryptography.fernet import Fernet
import base64

from autoword.core.llm_client import ModelType


logger = logging.getLogger(__name__)


@dataclass
class AppConfig:
    """应用配置数据类"""
    selected_model: str = "claude"
    api_keys: Dict[str, str] = field(default_factory=lambda: {
        "claude": "sk-3w1JFbWUq7tKjpLlopdkISQ9F6fpLhHx5viD0frh43ESE9Io",
        "gpt": "sk-NhjnJtqlZMx4PGTqvkGlH4POT82HHBrBnBbWOat99Bs5VZXi"
    })
    last_input_path: str = ""
    last_output_path: str = ""
    window_geometry: Dict[str, int] = field(default_factory=lambda: {
        "x": 100, "y": 100, "width": 800, "height": 600
    })
    auto_save_settings: bool = True
    theme: str = "dark"
    language: str = "zh_CN"


class SecureStorage:
    """安全存储类 - 负责API密钥的加密存储"""
    
    def __init__(self):
        self.key_file = Path.home() / ".autoword_key"
        self.key = self._get_or_create_key()
        self.cipher = Fernet(self.key)
    
    def _get_or_create_key(self) -> bytes:
        """获取或创建加密密钥"""
        try:
            if self.key_file.exists():
                with open(self.key_file, 'rb') as f:
                    return f.read()
            else:
                # 创建新密钥
                key = Fernet.generate_key()
                # 确保目录存在
                self.key_file.parent.mkdir(exist_ok=True)
                # 保存密钥
                with open(self.key_file, 'wb') as f:
                    f.write(key)
                # 设置文件权限（仅所有者可读写）
                if os.name != 'nt':  # Unix系统
                    os.chmod(self.key_file, 0o600)
                logger.info("创建新的加密密钥")
                return key
        except Exception as e:
            logger.error(f"密钥管理失败: {e}")
            raise
    
    def encrypt(self, data: str) -> str:
        """加密数据"""
        if not data:
            return ""
        try:
            encrypted_data = self.cipher.encrypt(data.encode('utf-8'))
            return base64.b64encode(encrypted_data).decode('utf-8')
        except Exception as e:
            logger.error(f"数据加密失败: {e}")
            raise
    
    def decrypt(self, encrypted_data: str) -> str:
        """解密数据"""
        if not encrypted_data:
            return ""
        try:
            decoded_data = base64.b64decode(encrypted_data.encode('utf-8'))
            decrypted_data = self.cipher.decrypt(decoded_data)
            return decrypted_data.decode('utf-8')
        except Exception as e:
            logger.error(f"数据解密失败: {e}")
            return ""  # 解密失败返回空字符串


class ConfigurationManager:
    """配置管理器"""
    
    def __init__(self):
        self.config_dir = Path.home() / ".autoword"
        self.config_file = self.config_dir / "config.json"
        self.secure_storage = SecureStorage()
        self.config = AppConfig()
        
        # 确保配置目录存在
        self.config_dir.mkdir(exist_ok=True)
        
        # 加载配置
        self.load_config()
    
    def load_config(self) -> AppConfig:
        """加载配置"""
        try:
            if self.config_file.exists():
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config_data = json.load(f)
                
                # 解密API密钥
                if 'api_keys' in config_data:
                    decrypted_keys = {}
                    for model, encrypted_key in config_data['api_keys'].items():
                        if encrypted_key:  # 只解密非空密钥
                            decrypted_keys[model] = self.secure_storage.decrypt(encrypted_key)
                        else:
                            decrypted_keys[model] = ""
                    config_data['api_keys'] = decrypted_keys
                
                # 更新配置对象，但保留默认API密钥
                for key, value in config_data.items():
                    if hasattr(self.config, key):
                        if key == 'api_keys':
                            # 合并API密钥，优先使用配置文件中的，但保留默认值作为后备
                            default_keys = self.config.api_keys.copy()
                            for model, key_value in value.items():
                                if key_value:  # 只覆盖非空密钥
                                    default_keys[model] = key_value
                            self.config.api_keys = default_keys
                        else:
                            setattr(self.config, key, value)
                
                logger.info("配置加载成功")
            else:
                logger.info("配置文件不存在，使用默认配置")
                # 确保默认API密钥存在
                if not self.config.api_keys.get("claude"):
                    self.config.api_keys["claude"] = "sk-3w1JFbWUq7tKjpLlopdkISQ9F6fpLhHx5viD0frh43ESE9Io"
                if not self.config.api_keys.get("gpt"):
                    self.config.api_keys["gpt"] = "sk-NhjnJtqlZMx4PGTqvkGlH4POT82HHBrBnBbWOat99Bs5VZXi"
                self.save_config()  # 创建默认配置文件
                
        except Exception as e:
            logger.error(f"配置加载失败: {e}")
            # 使用默认配置并确保API密钥存在
            self.config = AppConfig()
            if not self.config.api_keys.get("claude"):
                self.config.api_keys["claude"] = "sk-3w1JFbWUq7tKjpLlopdkISQ9F6fpLhHx5viD0frh43ESE9Io"
            if not self.config.api_keys.get("gpt"):
                self.config.api_keys["gpt"] = "sk-NhjnJtqlZMx4PGTqvkGlH4POT82HHBrBnBbWOat99Bs5VZXi"
        
        return self.config
    
    def save_config(self):
        """保存配置"""
        try:
            config_data = asdict(self.config)
            
            # 加密API密钥
            if config_data['api_keys']:
                encrypted_keys = {}
                for model, key in config_data['api_keys'].items():
                    if key:  # 只加密非空密钥
                        encrypted_keys[model] = self.secure_storage.encrypt(key)
                    else:
                        encrypted_keys[model] = ""
                config_data['api_keys'] = encrypted_keys
            
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config_data, f, indent=2, ensure_ascii=False)
            
            logger.info("配置保存成功")
            
        except Exception as e:
            logger.error(f"配置保存失败: {e}")
            raise
    
    def get_api_key(self, model_type: str) -> Optional[str]:
        """获取API密钥"""
        return self.config.api_keys.get(model_type, "")
    
    def set_api_key(self, model_type: str, api_key: str):
        """设置API密钥"""
        self.config.api_keys[model_type] = api_key
        if self.config.auto_save_settings:
            self.save_config()
    
    def get_last_paths(self) -> Dict[str, str]:
        """获取最后使用的路径"""
        return {
            "input": self.config.last_input_path,
            "output": self.config.last_output_path
        }
    
    def set_last_paths(self, input_path: str = None, output_path: str = None):
        """设置最后使用的路径"""
        if input_path is not None:
            self.config.last_input_path = input_path
        if output_path is not None:
            self.config.last_output_path = output_path
        
        if self.config.auto_save_settings:
            self.save_config()
    
    def get_selected_model(self) -> str:
        """获取选中的模型"""
        return self.config.selected_model
    
    def set_selected_model(self, model_type: str):
        """设置选中的模型"""
        self.config.selected_model = model_type
        if self.config.auto_save_settings:
            self.save_config()
    
    def get_window_geometry(self) -> Dict[str, int]:
        """获取窗口几何信息"""
        return self.config.window_geometry.copy()
    
    def set_window_geometry(self, x: int, y: int, width: int, height: int):
        """设置窗口几何信息"""
        self.config.window_geometry = {
            "x": x, "y": y, "width": width, "height": height
        }
        if self.config.auto_save_settings:
            self.save_config()
    
    def validate_config(self) -> Dict[str, str]:
        """验证配置完整性"""
        errors = {}
        
        # 检查选中的模型是否有API密钥
        selected_model = self.config.selected_model
        api_key = self.get_api_key(selected_model)
        
        if not api_key:
            errors['api_key'] = f"未设置{selected_model}模型的API密钥"
        
        # 检查模型类型是否有效
        valid_models = ["claude", "gpt"]
        if selected_model not in valid_models:
            errors['model'] = f"无效的模型类型: {selected_model}"
        
        return errors
    
    def reset_config(self):
        """重置配置为默认值"""
        try:
            self.config = AppConfig()
            self.save_config()
            logger.info("配置已重置为默认值")
        except Exception as e:
            logger.error(f"配置重置失败: {e}")
            raise
    
    def export_config(self, file_path: str, include_api_keys: bool = False):
        """导出配置到文件"""
        try:
            config_data = asdict(self.config)
            
            if not include_api_keys:
                # 移除API密钥
                config_data['api_keys'] = {}
            
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(config_data, f, indent=2, ensure_ascii=False)
            
            logger.info(f"配置已导出到: {file_path}")
            
        except Exception as e:
            logger.error(f"配置导出失败: {e}")
            raise
    
    def import_config(self, file_path: str):
        """从文件导入配置"""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                config_data = json.load(f)
            
            # 验证配置数据
            for key, value in config_data.items():
                if hasattr(self.config, key):
                    setattr(self.config, key, value)
            
            self.save_config()
            logger.info(f"配置已从文件导入: {file_path}")
            
        except Exception as e:
            logger.error(f"配置导入失败: {e}")
            raise
    
    def get_model_display_name(self, model_type: str) -> str:
        """获取模型显示名称"""
        model_names = {
            "claude": "Claude 3.7 Sonnet", 
            "gpt": "GPT-5"
        }
        return model_names.get(model_type, model_type)
    
    def get_available_models(self) -> Dict[str, str]:
        """获取可用模型列表"""
        return {
            "claude": "Claude 3.7 Sonnet",
            "gpt": "GPT-5"
        }