"""
Test AutoWord LLM Client
测试 LLM 客户端
"""

import json
import pytest
from unittest.mock import Mock, patch, MagicMock
from requests.exceptions import Timeout, ConnectionError, HTTPError

from autoword.core.llm_client import (
    LLMClient, ModelType, LLMResponse,
    call_gpt5, call_claude37
)
from autoword.core.exceptions import LLMError, ConfigurationError


class TestLLMClient:
    """测试 LLMClient 类"""
    
    def setup_method(self):
        """测试前设置"""
        self.client = LLMClient()
    
    def teardown_method(self):
        """测试后清理"""
        self.client.close()
    
    @patch.dict('os.environ', {'GPT5_KEY': 'test-gpt5-key'})
    def test_get_api_key_gpt5(self):
        """测试获取 GPT-5 API 密钥"""
        key = self.client._get_api_key(ModelType.GPT5)
        assert key == 'test-gpt5-key'
    
    @patch.dict('os.environ', {'CLAUDE37_KEY': 'test-claude-key'})
    def test_get_api_key_claude37(self):
        """测试获取 Claude 3.7 API 密钥"""
        key = self.client._get_api_key(ModelType.CLAUDE37)
        assert key == 'test-claude-key'
    
    @patch.dict('os.environ', {}, clear=True)
    def test_missing_api_key(self):
        """测试缺少 API 密钥"""
        with pytest.raises(ConfigurationError, match="Missing GPT5_KEY"):
            self.client._get_api_key(ModelType.GPT5)
    
    def test_build_payload(self):
        """测试构建请求载荷"""
        payload = self.client._build_payload(
            ModelType.GPT5,
            "System prompt",
            "User prompt",
            json_mode=True
        )
        
        expected = {
            "model": "gpt-5",
            "messages": [
                {"role": "system", "content": "System prompt"},
                {"role": "user", "content": "User prompt"}
            ],
            "temperature": 0.0,
            "top_p": 1.0,
            "response_format": {"type": "json_object"}
        }
        
        assert payload == expected
    
    def test_build_payload_no_json_mode(self):
        """测试构建非 JSON 模式载荷"""
        payload = self.client._build_payload(
            ModelType.GPT5,
            "System prompt",
            "User prompt",
            json_mode=False
        )
        
        assert "response_format" not in payload
    
    @patch('autoword.core.llm_client.time.time')
    @patch.dict('os.environ', {'GPT5_KEY': 'test-key'})
    def test_make_request_success(self, mock_time):
        """测试成功的请求"""
        # 模拟时间
        mock_time.side_effect = [1000.0, 1001.5]  # 1.5秒响应时间
        
        # 模拟响应
        mock_response = Mock()
        mock_response.status_code = 200
        mock_response.json.return_value = {
            "choices": [
                {
                    "message": {
                        "content": '{"result": "success"}'
                    }
                }
            ],
            "usage": {"total_tokens": 100}
        }
        
        with patch.object(self.client.session, 'post', return_value=mock_response):
            result = self.client._make_request(ModelType.GPT5, {"test": "payload"})
        
        assert result["content"] == '{"result": "success"}'
        assert result["usage"] == {"total_tokens": 100}
        assert result["response_time"] == 1.5
    
    @patch.dict('os.environ', {'GPT5_KEY': 'test-key'})
    def test_make_request_http_error(self):
        """测试 HTTP 错误"""
        mock_response = Mock()
        mock_response.status_code = 401
        mock_response.text = "Unauthorized"
        
        with patch.object(self.client.session, 'post', return_value=mock_response):
            with pytest.raises(LLMError, match="Authentication failed"):
                self.client._make_request(ModelType.GPT5, {"test": "payload"})
    
    @patch.dict('os.environ', {'GPT5_KEY': 'test-key'})
    def test_make_request_timeout(self):
        """测试请求超时"""
        with patch.object(self.client.session, 'post', side_effect=Timeout()):
            with pytest.raises(LLMError, match="Request timeout"):
                self.client._make_request(ModelType.GPT5, {"test": "payload"})
    
    @patch.dict('os.environ', {'GPT5_KEY': 'test-key'})
    def test_make_request_connection_error(self):
        """测试连接错误"""
        with patch.object(self.client.session, 'post', side_effect=ConnectionError()):
            with pytest.raises(LLMError, match="Connection error"):
                self.client._make_request(ModelType.GPT5, {"test": "payload"})
    
    def test_parse_json_response_success(self):
        """测试成功解析 JSON"""
        content = '{"tasks": [{"id": "1", "type": "rewrite"}]}'
        result = self.client._parse_json_response(content, ModelType.GPT5)
        
        expected = {"tasks": [{"id": "1", "type": "rewrite"}]}
        assert result == expected
    
    def test_parse_json_response_with_extra_text(self):
        """测试解析包含额外文本的 JSON"""
        content = 'Here is the JSON:\n{"tasks": [{"id": "1"}]}\nEnd of response.'
        result = self.client._parse_json_response(content, ModelType.GPT5)
        
        expected = {"tasks": [{"id": "1"}]}
        assert result == expected
    
    def test_parse_json_response_invalid(self):
        """测试解析无效 JSON"""
        content = 'This is not JSON at all'
        
        with pytest.raises(LLMError, match="Failed to parse JSON"):
            self.client._parse_json_response(content, ModelType.GPT5)
    
    @patch.object(LLMClient, '_make_request')
    @patch.dict('os.environ', {'GPT5_KEY': 'test-key'})
    def test_call_model_success(self, mock_make_request):
        """测试成功调用模型"""
        mock_make_request.return_value = {
            "content": '{"result": "success"}',
            "usage": {"total_tokens": 100},
            "response_time": 1.5
        }
        
        response = self.client.call_model(
            ModelType.GPT5,
            "System prompt",
            "User prompt"
        )
        
        assert isinstance(response, LLMResponse)
        assert response.content == '{"result": "success"}'
        assert response.model == "gpt-5"
        assert response.usage == {"total_tokens": 100}
        assert response.response_time == 1.5
    
    @patch.object(LLMClient, '_make_request')
    @patch('autoword.core.llm_client.time.sleep')
    @patch.dict('os.environ', {'GPT5_KEY': 'test-key'})
    def test_call_model_with_retry(self, mock_sleep, mock_make_request):
        """测试带重试的模型调用"""
        # 前两次失败，第三次成功
        mock_make_request.side_effect = [
            LLMError("First attempt failed"),
            LLMError("Second attempt failed"),
            {
                "content": '{"result": "success"}',
                "usage": {"total_tokens": 100},
                "response_time": 1.5
            }
        ]
        
        response = self.client.call_model(
            ModelType.GPT5,
            "System prompt",
            "User prompt"
        )
        
        assert response.content == '{"result": "success"}'
        assert mock_make_request.call_count == 3
        assert mock_sleep.call_count == 2  # 两次重试延迟
    
    @patch.object(LLMClient, 'call_model')
    def test_call_gpt5(self, mock_call_model):
        """测试 GPT-5 便捷方法"""
        mock_response = LLMResponse(
            content='{"result": "success"}',
            model="gpt-5",
            response_time=1.0
        )
        mock_call_model.return_value = mock_response
        
        response = self.client.call_gpt5("System", "User")
        
        mock_call_model.assert_called_once_with(
            ModelType.GPT5, "System", "User", True
        )
        assert response == mock_response
    
    @patch.object(LLMClient, 'call_model')
    def test_call_claude37(self, mock_call_model):
        """测试 Claude 3.7 便捷方法"""
        mock_response = LLMResponse(
            content='{"result": "success"}',
            model="claude-3-7-sonnet-20250219",
            response_time=1.0
        )
        mock_call_model.return_value = mock_response
        
        response = self.client.call_claude37("System", "User")
        
        mock_call_model.assert_called_once_with(
            ModelType.CLAUDE37, "System", "User", True
        )
        assert response == mock_response
    
    @patch.object(LLMClient, 'call_model')
    @patch.object(LLMClient, '_parse_json_response')
    def test_call_with_json_retry_success(self, mock_parse_json, mock_call_model):
        """测试 JSON 重试成功"""
        mock_response = LLMResponse(
            content='{"result": "success"}',
            model="gpt-5",
            response_time=1.0
        )
        mock_call_model.return_value = mock_response
        mock_parse_json.return_value = {"result": "success"}
        
        result = self.client.call_with_json_retry(
            ModelType.GPT5, "System", "User"
        )
        
        assert result == {"result": "success"}
        assert mock_call_model.call_count == 1
        assert mock_parse_json.call_count == 1
    
    @patch.object(LLMClient, 'call_model')
    @patch.object(LLMClient, '_parse_json_response')
    def test_call_with_json_retry_with_retries(self, mock_parse_json, mock_call_model):
        """测试 JSON 重试机制"""
        mock_response = LLMResponse(
            content='invalid json',
            model="gpt-5",
            response_time=1.0
        )
        mock_call_model.return_value = mock_response
        
        # 前两次解析失败，第三次成功
        mock_parse_json.side_effect = [
            LLMError("Failed to parse JSON"),
            LLMError("Failed to parse JSON"),
            {"result": "success"}
        ]
        
        result = self.client.call_with_json_retry(
            ModelType.GPT5, "System", "User", max_json_retries=2
        )
        
        assert result == {"result": "success"}
        assert mock_call_model.call_count == 3
        assert mock_parse_json.call_count == 3


class TestLLMResponse:
    """测试 LLMResponse 数据类"""
    
    def test_llm_response_creation(self):
        """测试创建 LLM 响应"""
        response = LLMResponse(
            content='{"result": "success"}',
            model="gpt-5",
            usage={"total_tokens": 100},
            response_time=1.5
        )
        
        assert response.content == '{"result": "success"}'
        assert response.model == "gpt-5"
        assert response.usage == {"total_tokens": 100}
        assert response.response_time == 1.5
    
    def test_llm_response_to_dict(self):
        """测试转换为字典"""
        response = LLMResponse(
            content='{"result": "success"}',
            model="gpt-5",
            usage={"total_tokens": 100},
            response_time=1.5
        )
        
        expected = {
            "content": '{"result": "success"}',
            "model": "gpt-5",
            "usage": {"total_tokens": 100},
            "response_time": 1.5
        }
        
        assert response.to_dict() == expected


class TestConvenienceFunctions:
    """测试便捷函数"""
    
    @patch('autoword.core.llm_client.LLMClient')
    def test_call_gpt5_function(self, mock_client_class):
        """测试 call_gpt5 便捷函数"""
        mock_client = Mock()
        mock_response = LLMResponse(
            content='{"result": "success"}',
            model="gpt-5",
            response_time=1.0
        )
        mock_client.call_gpt5.return_value = mock_response
        mock_client_class.return_value.__enter__.return_value = mock_client
        
        result = call_gpt5("System", "User")
        
        assert result == '{"result": "success"}'
        mock_client.call_gpt5.assert_called_once_with("System", "User", True)
    
    @patch('autoword.core.llm_client.LLMClient')
    def test_call_claude37_function(self, mock_client_class):
        """测试 call_claude37 便捷函数"""
        mock_client = Mock()
        mock_response = LLMResponse(
            content='{"result": "success"}',
            model="claude-3-7-sonnet-20250219",
            response_time=1.0
        )
        mock_client.call_claude37.return_value = mock_response
        mock_client_class.return_value.__enter__.return_value = mock_client
        
        result = call_claude37("System", "User")
        
        assert result == '{"result": "success"}'
        mock_client.call_claude37.assert_called_once_with("System", "User", True)


if __name__ == "__main__":
    pytest.main([__file__])