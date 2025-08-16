#!/usr/bin/env python3
"""
测试任务生成
"""

import sys
import json
sys.path.append('.')

from autoword.core.llm_client import LLMClient, ModelType

def test_task_generation():
    print("测试任务生成...")
    
    client = LLMClient()
    
    # 简化的系统提示词
    system_prompt = """你是一个Word文档自动化助手。根据用户提供的批注，生成对应的任务列表。

请严格按照以下JSON格式返回，不要添加任何其他文字：

{
  "tasks": [
    {
      "id": "task_1",
      "type": "rewrite",
      "locator": {"by": "find", "value": "要修改的文本"},
      "instruction": "具体的修改指令",
      "source_comment_id": "comment_1"
    }
  ]
}

支持的任务类型：
- rewrite: 重写内容
- insert: 插入内容  
- delete: 删除内容
- set_heading_level: 设置标题级别
- set_paragraph_style: 设置段落样式

定位方式：
- find: 通过文本查找
- heading: 通过标题查找
- bookmark: 通过书签查找"""

    # 简化的用户提示词
    user_prompt = """文档中有以下批注：

批注1 (comment_1): 作者张三 - "重写项目背景部分，增加市场分析和竞争对手分析，使内容更加详细和专业"
位置: "项目背景"

批注2 (comment_2): 作者李四 - "将此标题设置为2级标题"  
位置: "技术架构图"

批注3 (comment_3): 作者王五 - "删除过时的技术栈说明段落"
位置: "旧技术栈"

请为这些批注生成对应的任务。"""

    print("发送请求到GPT5...")
    response = client.call_gpt5(system_prompt, user_prompt)
    
    print(f"响应成功: {response.success}")
    print(f"响应内容:\n{response.content}")
    
    if response.success:
        try:
            tasks_data = json.loads(response.content)
            print(f"\nJSON解析成功!")
            print(f"任务数量: {len(tasks_data.get('tasks', []))}")
            
            for i, task in enumerate(tasks_data.get('tasks', []), 1):
                print(f"任务{i}: {task.get('type')} - {task.get('instruction')}")
                
        except json.JSONDecodeError as e:
            print(f"\nJSON解析失败: {e}")
            print("尝试清理响应内容...")
            
            # 尝试提取JSON部分
            content = response.content.strip()
            if content.startswith('```json'):
                content = content[7:]
            if content.endswith('```'):
                content = content[:-3]
            content = content.strip()
            
            try:
                tasks_data = json.loads(content)
                print("清理后JSON解析成功!")
                print(f"任务数量: {len(tasks_data.get('tasks', []))}")
            except:
                print("清理后仍然解析失败")

if __name__ == "__main__":
    test_task_generation()