"""
AutoWord LLM Client
LLM 服务客户端，支持 GPT-5 和 Claude 3.7
"""

import os
import json
import time
import logging
from typing import Dict, Any, Optional, Union
from dataclasses import dataclass
from enum import Enum

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

from .constants import API_BASE, API_TIMEOUT
from .exceptions import LLMError, ConfigurationError


logger = logging.getLogger(__name__)


class ModelType(str, Enum):
    """支持的模型类型"""
    GPT5 = "gpt-5"
    CLAUDE37 = "claude-3-7-sonnet-20250219"


@dataclass
class LLMResponse:
    """LLM 响应数据类"""
    content: str
    model: str
    usage: Optional[Dict[str, Any]] = None
    response_time: float = 0.0
    
    def to_dict(self) -> Dict[str, Any]:
        """转换为字典"""
        return {
            "content": self.content,
            "model": self.model,
            "usage": self.usage,
            "response_time": self.response_time
        }


class LLMClient:
    """LLM 客户端类"""
    
    def __init__(self, 
                 api_base: str = API_BASE,
                 timeout: int = API_TIMEOUT,
                 max_retries: int = 3,
                 retry_delay: float = 1.0):
        """
        初始化 LLM 客户端
        
        Args:
            api_base: API 基础 URL
            timeout: 请求超时时间(秒)
            max_retries: 最大重试次数
            retry_delay: 重试延迟(秒)
        """
        self.api_base = api_base
        self.timeout = timeout
        self.max_retries = max_retries
        self.retry_delay = retry_delay
        
        # 配置 HTTP 会话
        self.session = requests.Session()
        
        # 配置重试策略
        retry_strategy = Retry(
            total=max_retries,
            backoff_factor=retry_delay,
            status_forcelist=[429, 500, 502, 503, 504],
            allowed_methods=["POST"]
        )
        
        adapter = HTTPAdapter(max_retries=retry_strategy)
        self.session.mount("http://", adapter)
        self.session.mount("https://", adapter)
        
        # 设置默认请求头
        self.session.headers.update({
            "Accept": "application/json",
            "Content-Type": "application/json",
            "Connection": "keep-alive",
            "User-Agent": "AutoWord/1.0"
        })
    
    def _get_api_key(self, model: ModelType) -> str:
        """
        获取指定模型的 API 密钥
        
        Args:
            model: 模型类型
            
        Returns:
            API 密钥
            
        Raises:
            ConfigurationError: 如果密钥不存在
        """
        if model == ModelType.GPT5:
            key = os.getenv("GPT5_KEY")
            if not key:
                raise ConfigurationError("Missing GPT5_KEY environment variable")
        elif model == ModelType.CLAUDE37:
            key = os.getenv("CLAUDE37_KEY")
            if not key:
                raise ConfigurationError("Missing CLAUDE37_KEY environment variable")
        else:
            raise ConfigurationError(f"Unsupported model: {model}")
        
        return key
    
    def _build_payload(self, 
                      model: ModelType,
                      system_prompt: str,
                      user_prompt: str,
                      json_mode: bool = True,
                      temperature: float = 0.0,
                      top_p: float = 1.0) -> Dict[str, Any]:
        """
        构建请求载荷
        
        Args:
            model: 模型类型
            system_prompt: 系统提示词
            user_prompt: 用户提示词
            json_mode: 是否启用 JSON 模式
            temperature: 温度参数
            top_p: Top-p 参数
            
        Returns:
            请求载荷字典
        """
        payload = {
            "model": model.value,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            "temperature": temperature,
            "top_p": top_p
        }
        
        if json_mode:
            payload["response_format"] = {"type": "json_object"}
        
        return payload
    
    def _make_request(self, 
                     model: ModelType,
                     payload: Dict[str, Any]) -> Dict[str, Any]:
        """
        发送 HTTP 请求
        
        Args:
            model: 模型类型
            payload: 请求载荷
            
        Returns:
            响应数据
            
        Raises:
            LLMError: 请求失败时抛出
        """
        api_key = self._get_api_key(model)
        
        # 设置授权头
        headers = {"Authorization": f"Bearer {api_key}"}
        if self.api_base == API_BASE:
            headers["Host"] = "globalai.vip"
        
        try:
            start_time = time.time()
            
            response = self.session.post(
                f"{self.api_base}/chat/completions",
                headers=headers,
                json=payload,
                timeout=self.timeout
            )
            
            response_time = time.time() - start_time
            
            # 检查 HTTP 状态码
            if response.status_code == 401:
                raise LLMError(f"Authentication failed for {model.value}. Check API key.")
            elif response.status_code == 429:
                raise LLMError(f"Rate limit exceeded for {model.value}. Please try again later.")
            elif response.status_code >= 400:
                error_msg = f"HTTP {response.status_code}: {response.text}"
                raise LLMError(f"API request failed for {model.value}: {error_msg}")
            
            response.raise_for_status()
            
            # 解析响应
            try:
                data = response.json()
            except json.JSONDecodeError as e:
                raise LLMError(f"Invalid JSON response from {model.value}: {e}")
            
            # 检查响应格式
            if "choices" not in data or not data["choices"]:
                raise LLMError(f"Invalid response format from {model.value}: missing choices")
            
            choice = data["choices"][0]
            if "message" not in choice or "content" not in choice["message"]:
                raise LLMError(f"Invalid response format from {model.value}: missing message content")
            
            # 构建响应对象
            return {
                "content": choice["message"]["content"],
                "usage": data.get("usage"),
                "response_time": response_time
            }
            
        except requests.exceptions.Timeout:
            raise LLMError(f"Request timeout for {model.value} after {self.timeout} seconds")
        except requests.exceptions.ConnectionError:
            raise LLMError(f"Connection error for {model.value}. Check network connectivity.")
        except requests.exceptions.RequestException as e:
            raise LLMError(f"Request failed for {model.value}: {e}")
    
    def _parse_json_response(self, content: str, model: ModelType) -> Dict[str, Any]:
        """
        解析 JSON 响应内容
        
        Args:
            content: 响应内容
            model: 模型类型
            
        Returns:
            解析后的 JSON 数据
            
        Raises:
            LLMError: JSON 解析失败时抛出
        """
        try:
            return json.loads(content)
        except json.JSONDecodeError as e:
            logger.warning(f"JSON parse error for {model.value}: {e}")
            logger.warning(f"Raw content: {content[:500]}...")
            
            # 尝试提取 JSON 部分
            content = content.strip()
            
            # 查找 JSON 对象的开始和结束
            start_idx = content.find('{')
            end_idx = content.rfind('}')
            
            if start_idx != -1 and end_idx != -1 and end_idx > start_idx:
                json_part = content[start_idx:end_idx + 1]
                try:
                    return json.loads(json_part)
                except json.JSONDecodeError:
                    pass
            
            raise LLMError(f"Failed to parse JSON response from {model.value}: {e}")
    
    def call_model(self,
                  model: ModelType,
                  system_prompt: str,
                  user_prompt: str,
                  json_mode: bool = True,
                  temperature: float = 0.0,
                  top_p: float = 1.0) -> LLMResponse:
        """
        调用指定模型
        
        Args:
            model: 模型类型
            system_prompt: 系统提示词
            user_prompt: 用户提示词
            json_mode: 是否启用 JSON 模式
            temperature: 温度参数
            top_p: Top-p 参数
            
        Returns:
            LLM 响应对象
            
        Raises:
            LLMError: 调用失败时抛出
        """
        logger.info(f"Calling {model.value} with {len(user_prompt)} chars")
        
        payload = self._build_payload(
            model, system_prompt, user_prompt, 
            json_mode, temperature, top_p
        )
        
        # 执行请求（带重试）
        last_error = None
        for attempt in range(self.max_retries + 1):
            try:
                response_data = self._make_request(model, payload)
                
                return LLMResponse(
                    content=response_data["content"],
                    model=model.value,
                    usage=response_data.get("usage"),
                    response_time=response_data["response_time"]
                )
                
            except LLMError as e:
                last_error = e
                if attempt < self.max_retries:
                    wait_time = self.retry_delay * (2 ** attempt)
                    logger.warning(f"Attempt {attempt + 1} failed for {model.value}: {e}")
                    logger.info(f"Retrying in {wait_time} seconds...")
                    time.sleep(wait_time)
                else:
                    logger.error(f"All {self.max_retries + 1} attempts failed for {model.value}")
                    break
        
        raise last_error or LLMError(f"Unknown error calling {model.value}")
    
    def call_gpt5(self,
                 system_prompt: str,
                 user_prompt: str,
                 json_mode: bool = True) -> LLMResponse:
        """
        调用 GPT-5 模型
        
        Args:
            system_prompt: 系统提示词
            user_prompt: 用户提示词
            json_mode: 是否启用 JSON 模式
            
        Returns:
            LLM 响应对象
        """
        return self.call_model(
            ModelType.GPT5, system_prompt, user_prompt, json_mode
        )
    
    def call_claude37(self,
                     system_prompt: str,
                     user_prompt: str,
                     json_mode: bool = True) -> LLMResponse:
        """
        调用 Claude 3.7 模型
        
        Args:
            system_prompt: 系统提示词
            user_prompt: 用户提示词
            json_mode: 是否启用 JSON 模式
            
        Returns:
            LLM 响应对象
        """
        return self.call_model(
            ModelType.CLAUDE37, system_prompt, user_prompt, json_mode
        )
    
    def call_with_json_retry(self,
                           model: ModelType,
                           system_prompt: str,
                           user_prompt: str,
                           max_json_retries: int = 2) -> Dict[str, Any]:
        """
        调用模型并解析 JSON，失败时重试
        
        Args:
            model: 模型类型
            system_prompt: 系统提示词
            user_prompt: 用户提示词
            max_json_retries: JSON 解析最大重试次数
            
        Returns:
            解析后的 JSON 数据
            
        Raises:
            LLMError: 调用或解析失败时抛出
        """
        for attempt in range(max_json_retries + 1):
            try:
                response = self.call_model(
                    model, system_prompt, user_prompt, json_mode=True
                )
                
                return self._parse_json_response(response.content, model)
                
            except LLMError as e:
                if "parse JSON" in str(e) and attempt < max_json_retries:
                    logger.warning(f"JSON parse failed, retrying... (attempt {attempt + 1})")
                    # 在重试时强调 JSON 格式要求
                    enhanced_system = system_prompt + "\n\nIMPORTANT: You MUST respond with valid JSON only. No additional text or explanations."
                    system_prompt = enhanced_system
                    continue
                else:
                    raise e
        
        raise LLMError(f"Failed to get valid JSON response from {model.value} after {max_json_retries + 1} attempts")
    
    def close(self):
        """关闭客户端会话"""
        if hasattr(self, 'session'):
            self.session.close()
    
    def __enter__(self):
        """上下文管理器入口"""
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """上下文管理器出口"""
        self.close()


# 便捷函数
def call_gpt5(system_prompt: str, user_prompt: str, json_mode: bool = True) -> str:
    """
    便捷函数：调用 GPT-5
    
    Args:
        system_prompt: 系统提示词
        user_prompt: 用户提示词
        json_mode: 是否启用 JSON 模式
        
    Returns:
        响应内容
    """
    with LLMClient() as client:
        response = client.call_gpt5(system_prompt, user_prompt, json_mode)
        return response.content


def call_claude37(system_prompt: str, user_prompt: str, json_mode: bool = True) -> str:
    """
    便捷函数：调用 Claude 3.7
    
    Args:
        system_prompt: 系统提示词
        user_prompt: 用户提示词
        json_mode: 是否启用 JSON 模式
        
    Returns:
        响应内容
    """
    with LLMClient() as client:
        response = client.call_claude37(system_prompt, user_prompt, json_mode)
        return response.content