"""
AutoWord vNext Core Configuration Module
核心配置模块，提供统一的配置管理
"""

import os
import json
import http.client
from typing import Optional, Dict, Any, List
from dataclasses import dataclass, field
from pathlib import Path

@dataclass
class LLMConfig:
    """LLM配置类"""
    provider: str = "custom"  # custom, openai, anthropic
    model: str = "claude-3-7-sonnet-20250219"
    api_key: str = ""
    base_url: str = "globalai.vip"
    temperature: float = 0.1
    max_tokens: int = 4000
    timeout: int = 30
    retry_attempts: int = 3

@dataclass
class LocalizationConfig:
    """本地化配置类"""
    language: str = "zh-CN"
    style_aliases: Dict[str, str] = field(default_factory=lambda: {
        "Heading 1": "标题 1",
        "Heading 2": "标题 2", 
        "Heading 3": "标题 3",
        "Normal": "正文",
        "Title": "标题"
    })
    font_fallbacks: Dict[str, List[str]] = field(default_factory=lambda: {
        "楷体": ["楷体", "楷体_GB2312", "STKaiti"],
        "宋体": ["宋体", "SimSun", "NSimSun"],
        "黑体": ["黑体", "SimHei", "Microsoft YaHei"]
    })
    enable_font_fallback: bool = True
    log_fallbacks: bool = True

@dataclass
class ValidationConfig:
    """验证配置类"""
    strict_mode: bool = True
    rollback_on_failure: bool = True
    chapter_assertions: bool = True
    style_assertions: bool = True
    toc_assertions: bool = True
    pagination_assertions: bool = True
    forbidden_level1_headings: List[str] = field(default_factory=lambda: ["摘要", "参考文献"])

@dataclass
class AuditConfig:
    """审计配置类"""
    save_snapshots: bool = True
    generate_diff_reports: bool = True
    output_directory: str = "./audit_output"
    keep_audit_days: int = 30
    compress_old_audits: bool = True

@dataclass
class ExecutorConfig:
    """执行器配置类"""
    batch_operations: bool = True
    com_timeout: int = 30
    retry_com_errors: bool = True
    max_com_retries: int = 3
    enable_localization: bool = True

@dataclass
class VNextConfig:
    """vNext主配置类"""
    llm: LLMConfig = field(default_factory=LLMConfig)
    localization: LocalizationConfig = field(default_factory=LocalizationConfig)
    validation: ValidationConfig = field(default_factory=ValidationConfig)
    audit: AuditConfig = field(default_factory=AuditConfig)
    executor: ExecutorConfig = field(default_factory=ExecutorConfig)

class CustomLLMClient:
    """自定义LLM客户端，支持你的API服务"""
    
    def __init__(self, config: LLMConfig):
        self.config = config
        self.base_url = config.base_url
        self.api_key = config.api_key
        self.model = config.model
        
    def generate_plan(self, structure_json: str, user_intent: str) -> str:
        """生成执行计划"""
        try:
            # 构建提示词
            prompt = self._build_prompt(structure_json, user_intent)
            
            # 调用API
            response = self._call_api(prompt)
            
            return response
            
        except Exception as e:
            raise Exception(f"LLM API调用失败: {e}")
    
    def _build_prompt(self, structure_json: str, user_intent: str) -> str:
        """构建提示词"""
        return f"""你是一个Word文档处理专家。根据用户意图和文档结构，生成JSON格式的执行计划。

文档结构:
{structure_json}

用户意图:
{user_intent}

请生成一个JSON格式的执行计划，包含以下操作类型：
- delete_section_by_heading: 删除章节
- update_toc: 更新目录
- delete_toc: 删除目录
- set_style_rule: 设置样式
- reassign_paragraphs_to_style: 重新分配段落样式
- clear_direct_formatting: 清除直接格式

返回格式：
{{
  "schema_version": "plan.v1",
  "ops": [
    {{
      "operation_type": "操作类型",
      "参数名": "参数值"
    }}
  ]
}}

只返回JSON，不要其他内容。"""

    def _call_api(self, prompt: str) -> str:
        """调用自定义API"""
        try:
            conn = http.client.HTTPSConnection(self.base_url)
            
            payload = json.dumps({
                "model": self.model,
                "messages": [
                    {
                        "role": "system",
                        "content": "You are a helpful assistant that generates JSON execution plans for Word document processing."
                    },
                    {
                        "role": "user", 
                        "content": prompt
                    }
                ],
                "temperature": self.config.temperature,
                "max_tokens": self.config.max_tokens
            })
            
            headers = {
                'Accept': 'application/json',
                'Authorization': f'Bearer {self.api_key}',
                'Content-Type': 'application/json',
                'Host': self.base_url,
                'Connection': 'keep-alive'
            }
            
            conn.request("POST", "/v1/chat/completions", payload, headers)
            res = conn.getresponse()
            data = res.read()
            
            response_json = json.loads(data.decode("utf-8"))
            
            if 'choices' in response_json and len(response_json['choices']) > 0:
                content = response_json['choices'][0]['message']['content']
                return content.strip()
            else:
                raise Exception(f"API响应格式错误: {response_json}")
                
        except Exception as e:
            raise Exception(f"API调用失败: {e}")

def load_config(config_path: str = "vnext_config.json") -> VNextConfig:
    """加载配置文件"""
    if not os.path.exists(config_path):
        return VNextConfig()
    
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            config_data = json.load(f)
        
        # 创建配置对象
        config = VNextConfig()
        
        # 加载LLM配置
        if 'llm' in config_data:
            llm_data = config_data['llm']
            config.llm = LLMConfig(
                provider=llm_data.get('provider', 'custom'),
                model=llm_data.get('model', 'claude-3-7-sonnet-20250219'),
                api_key=llm_data.get('api_key', ''),
                base_url=llm_data.get('base_url', 'globalai.vip'),
                temperature=llm_data.get('temperature', 0.1),
                max_tokens=llm_data.get('max_tokens', 4000)
            )
        
        # 加载其他配置...
        if 'localization' in config_data:
            loc_data = config_data['localization']
            config.localization = LocalizationConfig(
                language=loc_data.get('language', 'zh-CN'),
                style_aliases=loc_data.get('style_aliases', config.localization.style_aliases),
                font_fallbacks=loc_data.get('font_fallbacks', config.localization.font_fallbacks)
            )
        
        return config
        
    except Exception as e:
        print(f"配置加载失败: {e}")
        return VNextConfig()

def save_config(config: VNextConfig, config_path: str = "vnext_config.json"):
    """保存配置文件"""
    try:
        config_data = {
            "llm": {
                "provider": config.llm.provider,
                "model": config.llm.model,
                "api_key": config.llm.api_key,
                "base_url": config.llm.base_url,
                "temperature": config.llm.temperature,
                "max_tokens": config.llm.max_tokens
            },
            "localization": {
                "language": config.localization.language,
                "style_aliases": config.localization.style_aliases,
                "font_fallbacks": config.localization.font_fallbacks
            },
            "validation": {
                "strict_mode": config.validation.strict_mode,
                "rollback_on_failure": config.validation.rollback_on_failure
            },
            "audit": {
                "save_snapshots": config.audit.save_snapshots,
                "generate_diff_reports": config.audit.generate_diff_reports,
                "output_directory": config.audit.output_directory
            }
        }
        
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(config_data, f, indent=2, ensure_ascii=False)
            
    except Exception as e:
        raise Exception(f"配置保存失败: {e}")