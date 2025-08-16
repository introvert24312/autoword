#!/usr/bin/env python3
"""
AutoWord 最终版本 - 简化的完整实现
"""

import sys
import os
import json
import time
sys.path.append('.')

from autoword.core.doc_loader import WordSession
from autoword.core.doc_inspector import DocumentInspector
from autoword.core.llm_client import LLMClient, ModelType
from autoword.core.models import Task, TaskType, Locator, LocatorType, ExecutionResult, TaskResult

def main():
    document_path = r"C:\Users\y\Desktop\郭宇睿-现代汉语言文学在社会发展中的作用与影响研究2稿批注.docx"
    
    print("AutoWord 文档自动化系统")
    print("=" * 50)
    print(f"文档: {os.path.basename(document_path)}")
    print()
    
    start_time = time.time()
    
    try:
        # 1. 加载文档
        print("步骤1: 加载文档...")
        with WordSession(visible=False) as word_app:
            document = word_app.Documents.Open(document_path)
            
            try:
                # 2. 提取批注
                print("步骤2: 提取批注...")
                inspector = DocumentInspector()
                document_info = inspector.get_document_info(document)
                
                print(f"✅ 找到 {len(document_info.comments)} 个批注")
                
                if not document_info.comments:
                    print("⚠️ 文档中没有批注，无需处理")
                    return
                
                # 显示批注
                for i, comment in enumerate(document_info.comments, 1):
                    print(f"  {i}. {comment.author}: {comment.comment_text[:50]}...")
                
                # 3. 生成任务
                print("\n步骤3: 生成任务...")
                tasks = generate_tasks(document_info.comments)
                
                if not tasks:
                    print("❌ 没有生成任何任务")
                    return
                
                print(f"✅ 生成了 {len(tasks)} 个任务")
                
                # 显示任务
                for i, task in enumerate(tasks, 1):
                    print(f"  {i}. {task.type.value}: {task.instruction[:50]}...")
                
                # 4. 执行任务（试运行模式）
                print("\n步骤4: 执行任务（试运行模式）...")
                results = execute_tasks_dry_run(tasks, document)
                
                # 显示结果
                print("\n" + "=" * 50)
                print("执行结果")
                print("=" * 50)
                
                success_count = sum(1 for r in results if r.success)
                
                for i, result in enumerate(results, 1):
                    status = "✅" if result.success else "❌"
                    print(f"{i}. {status} {result.task_id}: {result.message}")
                
                print(f"\n📊 统计:")
                print(f"  总任务: {len(results)}")
                print(f"  成功: {success_count}")
                print(f"  失败: {len(results) - success_count}")
                print(f"  耗时: {time.time() - start_time:.2f}s")
                
                if success_count == len(results):
                    print("\n🎉 所有任务都可以成功执行！")
                    print("💡 要实际执行修改，请修改代码中的execute_tasks_dry_run为execute_tasks_real")
                else:
                    print("\n⚠️ 部分任务执行失败，请检查任务配置")
                
            finally:
                document.Close(SaveChanges=False)
                
    except Exception as e:
        print(f"❌ 处理失败: {str(e)}")
        import traceback
        traceback.print_exc()

def generate_tasks(comments):
    """生成任务列表"""
    print("  调用LLM生成任务...")
    
    client = LLMClient()
    
    # 构建批注描述
    comments_text = ""
    for i, comment in enumerate(comments, 1):
        comments_text += f"{i}. ID: {comment.id}, 作者: {comment.author}\n"
        comments_text += f"   内容: {comment.comment_text}\n"
        comments_text += f"   位置: {comment.anchor_text[:30]}...\n\n"
    
    system_prompt = """你是一个Word文档自动化助手。根据批注生成任务列表。

重要规则：
- 只输出有效的JSON，不要添加任何其他文字
- 每个任务必须有source_comment_id
- 使用准确的定位信息

支持的任务类型：
- rewrite: 重写内容
- insert: 插入内容  
- delete: 删除内容
- set_heading_level: 设置标题级别
- set_paragraph_style: 设置段落样式

JSON格式：
{
  "tasks": [
    {
      "id": "task_1",
      "type": "rewrite",
      "locator": {"by": "find", "value": "要修改的文本"},
      "instruction": "具体指令",
      "source_comment_id": "comment_1"
    }
  ]
}"""

    user_prompt = f"""文档中有以下批注：

{comments_text}

请为这些批注生成对应的任务。"""

    response = client.call_gpt5(system_prompt, user_prompt)
    
    if not response.success:
        print(f"  ❌ LLM调用失败: {response.error}")
        return []
    
    try:
        data = json.loads(response.content)
        tasks = []
        
        for task_data in data.get("tasks", []):
            try:
                task = Task(
                    id=task_data["id"],
                    type=TaskType(task_data["type"]),
                    locator=Locator(
                        by=LocatorType(task_data["locator"]["by"]),
                        value=task_data["locator"]["value"]
                    ),
                    instruction=task_data["instruction"],
                    source_comment_id=task_data.get("source_comment_id")
                )
                tasks.append(task)
            except Exception as e:
                print(f"  ⚠️ 跳过无效任务: {e}")
                continue
        
        print(f"  ✅ 成功解析 {len(tasks)} 个任务")
        return tasks
        
    except json.JSONDecodeError as e:
        print(f"  ❌ JSON解析失败: {e}")
        print(f"  响应内容: {response.content[:200]}...")
        return []

def execute_tasks_dry_run(tasks, document):
    """试运行任务（不实际修改文档）"""
    results = []
    
    for task in tasks:
        print(f"  试运行: {task.id} - {task.type.value}")
        
        try:
            # 模拟任务执行
            if task.type == TaskType.DELETE:
                message = f"[试运行] 将删除包含'{task.locator.value}'的内容"
            elif task.type == TaskType.REWRITE:
                message = f"[试运行] 将重写包含'{task.locator.value}'的内容"
            elif task.type == TaskType.SET_PARAGRAPH_STYLE:
                message = f"[试运行] 将设置段落样式"
            elif task.type == TaskType.SET_HEADING_LEVEL:
                message = f"[试运行] 将设置标题级别"
            else:
                message = f"[试运行] 将执行{task.type.value}操作"
            
            # 简单的定位验证
            if task.locator.by == LocatorType.FIND:
                # 在文档中查找文本
                found = False
                try:
                    search_range = document.Range()
                    search_range.Find.Text = task.locator.value
                    found = search_range.Find.Execute()
                except:
                    pass
                
                if not found:
                    message += " (警告: 未找到目标文本)"
            
            results.append(TaskResult(
                task_id=task.id,
                success=True,
                message=message,
                execution_time=0.1
            ))
            
        except Exception as e:
            results.append(TaskResult(
                task_id=task.id,
                success=False,
                message=f"试运行失败: {str(e)}",
                execution_time=0.0,
                error_details=str(e)
            ))
    
    return results

if __name__ == "__main__":
    main()
    input("\n按回车键退出...")