"""
AutoWord LLM Client
支持 GPT-5 和 Claude 3.7 的 LLM 客户端（中转API）
"""

import os
import json
import logging
import time
import http.client
from typing import Optional, Dict, Any
from enum import Enum
from dataclasses import dataclass

from .exceptions import LLMError, APIKeyError


logger = logging.getLogger(__name__)


class ModelType(Enum):
    """支持的模型类型"""
    GPT5 = "gpt-4o"  # 使用gpt-4o作为GPT5的模型名
    CLAUDE37 = "claude-3-7-sonnet-20250219"


@dataclass
class LLMResponse:
    """LLM 响应数据类"""
    success: bool
    content: str
    model: str
    usage: Optional[Dict[str, Any]] = None
    error: Optional[str] = None
    
    def to_dict(self) -> Dict[str, Any]:
        """转换为字典"""
        return {
            "success": self.success,
            "content": self.content,
            "model": self.model,
            "usage": self.usage,
            "error": self.error
        }


class LLMClient:
    """LLM 客户端"""
    
    def __init__(self, 
                 api_keys: Optional[Dict[str, str]] = None,
                 base_url: str = "globalai.vip",
                 timeout: int = 60,
                 max_retries: int = 3):
        """
        初始化 LLM 客户端
        
        Args:
            api_keys: API密钥字典 {"gpt": "key", "claude": "key"}
            base_url: API 基础URL
            timeout: 请求超时时间
            max_retries: 最大重试次数
        """
        self.base_url = base_url
        self.timeout = timeout
        self.max_retries = max_retries
        
        # API密钥 - 默认值作为后备
        self.api_keys = api_keys or {
            "gpt": "sk-NhjnJtqlZMx4PGTqvkGlH4POT82HHBrBnBbWOat99Bs5VZXi",
            "claude": "sk-3w1JFbWUq7tKjpLlopdkISQ9F6fpLhHx5viD0frh43ESE9Io"
        }
    
    def _get_api_key(self, model_type: ModelType) -> str:
        """获取API密钥"""
        if model_type == ModelType.GPT5:
            key = self.api_keys.get("gpt", "")
        elif model_type == ModelType.CLAUDE37:
            key = self.api_keys.get("claude", "")
        else:
            raise APIKeyError(f"不支持的模型类型: {model_type}")
        
        if not key:
            raise APIKeyError(f"未设置{model_type.value}的API密钥")
        
        return key
    
    def _make_request(self, model_type: ModelType, messages: list, 
                     temperature: float = 0.7) -> Dict[str, Any]:
        """发送API请求"""
        api_key = self._get_api_key(model_type)
        
        payload = {
            "model": model_type.value,
            "messages": messages,
            "temperature": temperature,
            "max_tokens": 4000
        }
        
        headers = {
            'Accept': 'application/json',
            'Authorization': f'Bearer {api_key}',
            'Content-Type': 'application/json',
            'Host': self.base_url,
            'Connection': 'keep-alive'
        }
        
        for attempt in range(self.max_retries):
            try:
                conn = http.client.HTTPSConnection(self.base_url, timeout=self.timeout)
                
                payload_json = json.dumps(payload)
                conn.request("POST", "/v1/chat/completions", payload_json, headers)
                
                response = conn.getresponse()
                data = response.read()
                
                if response.status == 200:
                    result = json.loads(data.decode("utf-8"))
                    return result
                else:
                    error_msg = f"API请求失败: HTTP {response.status}"
                    if attempt < self.max_retries - 1:
                        logger.warning(f"{error_msg}, 重试中... ({attempt + 1}/{self.max_retries})")
                        time.sleep(2 ** attempt)  # 指数退避
                        continue
                    else:
                        raise LLMError(error_msg)
                        
            except Exception as e:
                if attempt < self.max_retries - 1:
                    logger.warning(f"请求异常: {str(e)}, 重试中... ({attempt + 1}/{self.max_retries})")
                    time.sleep(2 ** attempt)
                    continue
                else:
                    raise LLMError(f"API请求失败: {str(e)}")
            finally:
                try:
                    conn.close()
                except:
                    pass
        
        raise LLMError("所有重试都失败了")
    
    def _parse_response(self, response_data: Dict[str, Any], model_type: ModelType) -> LLMResponse:
        """解析API响应"""
        try:
            if "choices" in response_data and len(response_data["choices"]) > 0:
                content = response_data["choices"][0]["message"]["content"]
                usage = response_data.get("usage", {})
                
                # 清理JSON标记
                content = self._clean_json_content(content)
                
                return LLMResponse(
                    success=True,
                    content=content,
                    model=model_type.value,
                    usage=usage
                )
            else:
                error_msg = response_data.get("error", {}).get("message", "未知错误")
                return LLMResponse(
                    success=False,
                    content="",
                    model=model_type.value,
                    error=error_msg
                )
                
        except Exception as e:
            return LLMResponse(
                success=False,
                content="",
                model=model_type.value,
                error=f"响应解析失败: {str(e)}"
            )
    
    def _clean_json_content(self, content: str) -> str:
        """清理JSON内容，移除markdown标记和修复常见格式问题"""
        content = content.strip()
        
        # 移除```json和```标记
        if content.startswith('```json'):
            content = content[7:]
        elif content.startswith('```'):
            content = content[3:]
            
        if content.endswith('```'):
            content = content[:-3]
        
        content = content.strip()
        
        # 修复常见的JSON格式问题
        import re
        
        # 移除注释行（// 和 /* */ 风格）
        content = re.sub(r'//.*?\n', '\n', content)
        content = re.sub(r'/\*.*?\*/', '', content, flags=re.DOTALL)
        
        # 修复尾随逗号问题
        content = re.sub(r',(\s*[}\]])', r'\1', content)
        
        # 修复缺少逗号的问题
        content = re.sub(r'"\s*\n\s*"', '",\n"', content)
        content = re.sub(r'}\s*\n\s*{', '},\n{', content)
        content = re.sub(r']\s*\n\s*\[', '],\n[', content)
        content = re.sub(r'"\s+"', '", "', content)  # 修复字符串间缺少逗号
        
        # 修复未引用的字符串值
        content = re.sub(r':\s*([a-zA-Z_][a-zA-Z0-9_]*)\s*([,}\]])', r': "\1"\2', content)
        
        # 修复单引号为双引号
        content = re.sub(r"'([^']*)'", r'"\1"', content)
        
        # 确保JSON对象/数组的开始和结束
        content = content.strip()
        if not content.startswith(('{', '[')):
            # 尝试找到第一个 { 或 [
            start_pos = min(
                (content.find('{') if content.find('{') != -1 else len(content)),
                (content.find('[') if content.find('[') != -1 else len(content))
            )
            if start_pos < len(content):
                content = content[start_pos:]
        
        return content.strip()
    
    def call_model(self, 
                   model_type: ModelType,
                   system_prompt: str,
                   user_prompt: str,
                   temperature: float = 0.7) -> LLMResponse:
        """
        调用指定模型
        
        Args:
            model_type: 模型类型
            system_prompt: 系统提示词
            user_prompt: 用户提示词
            temperature: 温度参数
            
        Returns:
            LLM响应
        """
        messages = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ]
        
        try:
            response_data = self._make_request(model_type, messages, temperature)
            return self._parse_response(response_data, model_type)
        except Exception as e:
            logger.error(f"模型调用失败: {str(e)}")
            return LLMResponse(
                success=False,
                content="",
                model=model_type.value,
                error=str(e)
            )
    
    def call_gpt5(self, system_prompt: str, user_prompt: str, 
                  temperature: float = 0.7) -> LLMResponse:
        """调用GPT-5模型"""
        return self.call_model(ModelType.GPT5, system_prompt, user_prompt, temperature)
    
    def call_claude37(self, system_prompt: str, user_prompt: str,
                      temperature: float = 0.7) -> LLMResponse:
        """调用Claude 3.7模型"""
        return self.call_model(ModelType.CLAUDE37, system_prompt, user_prompt, temperature)
    
    def call_with_json_retry(self, 
                           model_type: ModelType,
                           system_prompt: str,
                           user_prompt: str,
                           max_json_retries: int = 3) -> LLMResponse:
        """
        调用模型并重试JSON解析
        
        Args:
            model_type: 模型类型
            system_prompt: 系统提示词
            user_prompt: 用户提示词
            max_json_retries: JSON解析最大重试次数
            
        Returns:
            LLM响应
        """
        original_system_prompt = system_prompt
        
        for attempt in range(max_json_retries):
            response = self.call_model(model_type, system_prompt, user_prompt)
            
            if not response.success:
                return response
            
            # 尝试多种JSON修复策略
            json_candidates = [
                response.content,  # 原始内容
                self._clean_json_content(response.content),  # 清理后的内容
                self._aggressive_json_fix(response.content),  # 激进修复
            ]
            
            for i, candidate in enumerate(json_candidates):
                try:
                    parsed = json.loads(candidate)
                    # JSON解析成功，返回修复后的内容
                    return LLMResponse(
                        success=True,
                        content=candidate,
                        model=response.model,
                        usage=response.usage
                    )
                except json.JSONDecodeError as e:
                    if i == len(json_candidates) - 1:  # 最后一个候选也失败了
                        logger.warning(f"JSON解析失败 (尝试 {i+1}/{len(json_candidates)}): {str(e)}")
            
            # 所有修复策略都失败了，准备重试
            if attempt < max_json_retries - 1:
                logger.warning(f"JSON解析失败，重试中... ({attempt + 1}/{max_json_retries})")
                # 逐步增强系统提示词
                if attempt == 0:
                    system_prompt = original_system_prompt + "\n\n重要：请确保返回有效的JSON格式，不要包含任何额外的文本、说明或markdown标记。"
                elif attempt == 1:
                    system_prompt = original_system_prompt + "\n\n严格要求：只返回纯JSON，格式如下：\n{\n  \"tasks\": [\n    {\"id\": \"task_1\", \"type\": \"rewrite\", ...}\n  ]\n}\n不要添加任何解释文字。"
                continue
            else:
                # 最后一次尝试也失败了
                return LLMResponse(
                    success=False,
                    content=response.content,
                    model=model_type.value,
                    error=f"JSON解析失败，已尝试 {max_json_retries} 次: 最后错误为格式问题"
                )
        
        return response
    
    def _aggressive_json_fix(self, content: str) -> str:
        """激进的JSON修复策略"""
        import re
        
        content = content.strip()
        
        # 移除所有markdown标记
        content = re.sub(r'```[a-zA-Z]*\n?', '', content)
        content = re.sub(r'```', '', content)
        
        # 移除前后的非JSON文本
        # 找到第一个 { 或 [
        start_match = re.search(r'[{\[]', content)
        if start_match:
            content = content[start_match.start():]
        
        # 找到最后一个 } 或 ]
        end_match = None
        for match in re.finditer(r'[}\]]', content):
            end_match = match
        if end_match:
            content = content[:end_match.end()]
        
        # 修复常见的JSON错误
        content = re.sub(r',(\s*[}\]])', r'\1', content)  # 移除尾随逗号
        content = re.sub(r'([}\]])(\s*)([{\[])', r'\1,\2\3', content)  # 添加缺失的逗号
        content = re.sub(r'"\s*\n\s*"', '",\n"', content)  # 修复字符串间的逗号
        
        # 修复未引用的键名
        content = re.sub(r'([{,]\s*)([a-zA-Z_][a-zA-Z0-9_]*)\s*:', r'\1"\2":', content)
        
        # 修复未引用的字符串值（但不影响数字、布尔值、null）
        content = re.sub(r':\s*([a-zA-Z_][a-zA-Z0-9_\s]*?)(\s*[,}\]])', 
                        lambda m: f': "{m.group(1).strip()}"{m.group(2)}' 
                        if m.group(1).strip() not in ['true', 'false', 'null'] and not m.group(1).strip().isdigit()
                        else m.group(0), content)
        
        return content.strip()


# 便捷函数
def call_gpt5(system_prompt: str, user_prompt: str, temperature: float = 0.7) -> LLMResponse:
    """便捷函数：调用GPT-5"""
    client = LLMClient()
    return client.call_gpt5(system_prompt, user_prompt, temperature)


def call_claude37(system_prompt: str, user_prompt: str, temperature: float = 0.7) -> LLMResponse:
    """便捷函数：调用Claude 3.7"""
    client = LLMClient()
    return client.call_claude37(system_prompt, user_prompt, temperature)