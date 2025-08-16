"""
AutoWord LLM Client Demo
LLM 客户端使用示例
"""

import os
import sys
import json

# 添加项目根目录到路径
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from autoword.core.llm_client import LLMClient, ModelType
from autoword.core.constants import SYSTEM_PROMPT


def demo_basic_usage():
    """演示基本用法"""
    print("=== AutoWord LLM Client Demo ===\n")
    
    # 检查环境变量
    gpt5_key = os.getenv("GPT5_KEY")
    claude_key = os.getenv("CLAUDE37_KEY")
    
    if not gpt5_key and not claude_key:
        print("❌ 未找到 API 密钥环境变量")
        print("请设置 GPT5_KEY 或 CLAUDE37_KEY 环境变量")
        return
    
    # 创建客户端
    with LLMClient() as client:
        print("✅ LLM 客户端已创建")
        
        # 示例文档结构和批注
        sample_structure = """
        文档结构摘要:
        - 标题: "项目报告"
        - 1级标题: "概述", "方法", "结果", "结论"
        - 样式: 正文、标题1、标题2
        - TOC: 自动生成，3级
        """
        
        sample_comments = """
        批注列表:
        1. ID: comment_1, 作者: 张三, 页码: 1
           锚点: "项目背景介绍"
           内容: "这段需要重写，更加简洁明了"
        
        2. ID: comment_2, 作者: 李四, 页码: 2  
           锚点: "数据分析方法"
           内容: "插入一个关于统计方法的段落"
        """
        
        # 示例 JSON Schema（简化版）
        sample_schema = {
            "type": "object",
            "properties": {
                "tasks": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "properties": {
                            "id": {"type": "string"},
                            "type": {"type": "string"},
                            "instruction": {"type": "string"}
                        }
                    }
                }
            }
        }
        
        user_prompt = f"""
        文档结构摘要:
        {sample_structure}
        
        批注列表:
        {sample_comments}
        
        输出 Schema:
        {json.dumps(sample_schema, ensure_ascii=False, indent=2)}
        
        约束:
        - 除非批注明确要求，否则不要改变格式
        - 允许无批注的操作: rewrite/insert/delete
        """
        
        # 尝试调用可用的模型
        if gpt5_key:
            try:
                print("\n🚀 调用 GPT-5...")
                response = client.call_gpt5(SYSTEM_PROMPT, user_prompt)
                print(f"✅ GPT-5 响应成功")
                print(f"📊 响应时间: {response.response_time:.2f}s")
                print(f"🔤 内容长度: {len(response.content)} 字符")
                
                # 尝试解析 JSON
                try:
                    parsed = json.loads(response.content)
                    print(f"✅ JSON 解析成功，包含 {len(parsed.get('tasks', []))} 个任务")
                except json.JSONDecodeError:
                    print("⚠️ 响应不是有效的 JSON 格式")
                
                print(f"📝 响应内容预览:\n{response.content[:200]}...")
                
            except Exception as e:
                print(f"❌ GPT-5 调用失败: {e}")
        
        if claude_key:
            try:
                print("\n🚀 调用 Claude 3.7...")
                response = client.call_claude37(SYSTEM_PROMPT, user_prompt)
                print(f"✅ Claude 3.7 响应成功")
                print(f"📊 响应时间: {response.response_time:.2f}s")
                print(f"🔤 内容长度: {len(response.content)} 字符")
                
                # 尝试解析 JSON
                try:
                    parsed = json.loads(response.content)
                    print(f"✅ JSON 解析成功，包含 {len(parsed.get('tasks', []))} 个任务")
                except json.JSONDecodeError:
                    print("⚠️ 响应不是有效的 JSON 格式")
                
                print(f"📝 响应内容预览:\n{response.content[:200]}...")
                
            except Exception as e:
                print(f"❌ Claude 3.7 调用失败: {e}")


def demo_json_retry():
    """演示 JSON 重试功能"""
    print("\n=== JSON 重试功能演示 ===\n")
    
    # 检查环境变量
    if not os.getenv("GPT5_KEY") and not os.getenv("CLAUDE37_KEY"):
        print("❌ 未找到 API 密钥，跳过演示")
        return
    
    with LLMClient() as client:
        simple_prompt = """
        请生成一个简单的任务列表 JSON:
        {
            "tasks": [
                {"id": "1", "type": "rewrite", "instruction": "重写第一段"}
            ]
        }
        """
        
        try:
            # 使用可用的模型
            model = ModelType.GPT5 if os.getenv("GPT5_KEY") else ModelType.CLAUDE37
            
            print(f"🚀 使用 {model.value} 测试 JSON 重试...")
            
            result = client.call_with_json_retry(
                model, 
                SYSTEM_PROMPT, 
                simple_prompt,
                max_json_retries=2
            )
            
            print("✅ JSON 重试成功")
            print(f"📋 解析结果: {json.dumps(result, ensure_ascii=False, indent=2)}")
            
        except Exception as e:
            print(f"❌ JSON 重试失败: {e}")


if __name__ == "__main__":
    demo_basic_usage()
    demo_json_retry()
    
    print("\n=== 演示完成 ===")
    print("💡 提示: 设置环境变量 GPT5_KEY 或 CLAUDE37_KEY 来测试真实 API 调用")