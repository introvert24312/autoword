#!/usr/bin/env python3
"""
AutoWord 完整工作流程演示
展示从文档分析到任务执行的完整流程
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from autoword.core.enhanced_executor import EnhancedExecutor, WorkflowMode
from autoword.core.llm_client import ModelType
from autoword.core.models import Comment, Document
from autoword.core.exceptions import AutoWordError
import time


def demo_complete_workflow():
    """演示完整的工作流程"""
    print("=== AutoWord 完整工作流程演示 ===")
    print()
    
    # 检查 API 密钥
    api_key_available = bool(os.getenv('GPT5_KEY') or os.getenv('CLAUDE37_KEY'))
    
    if not api_key_available:
        print("⚠️ 未检测到 API 密钥，将使用模拟模式演示")
        print("💡 设置环境变量 GPT5_KEY 或 CLAUDE37_KEY 以使用真实 API")
        print()
    
    # 创建执行器
    print("🚀 创建 AutoWord 执行器...")
    executor = EnhancedExecutor(
        llm_model=ModelType.GPT5 if os.getenv('GPT5_KEY') else ModelType.CLAUDE37,
        visible=False  # 隐藏 Word 窗口
    )
    print("✅ 执行器创建完成")
    print()
    
    # 演示不同的工作流程模式
    demo_modes = [
        {
            "name": "试运行模式",
            "mode": WorkflowMode.DRY_RUN,
            "description": "预览修改效果，不实际修改文档",
            "icon": "🔍"
        },
        {
            "name": "安全模式", 
            "mode": WorkflowMode.SAFE,
            "description": "自动备份，支持回滚保护",
            "icon": "🛡️"
        },
        {
            "name": "正常模式",
            "mode": WorkflowMode.NORMAL,
            "description": "直接执行修改，高效快速",
            "icon": "⚡"
        }
    ]
    
    print("📋 支持的工作流程模式:")
    for mode_info in demo_modes:
        print(f"  {mode_info['icon']} {mode_info['name']}: {mode_info['description']}")
    print()
    
    # 模拟文档处理
    print("📄 模拟文档处理流程...")
    
    # 创建示例文档信息
    sample_document = {
        "path": "sample_document.docx",
        "title": "项目技术方案",
        "page_count": 8,
        "word_count": 3500,
        "comments": [
            {
                "id": "comment_1",
                "author": "张三",
                "text": "重写项目背景部分，增加市场分析和竞争对手分析",
                "anchor_text": "项目背景",
                "page_number": 1
            },
            {
                "id": "comment_2",
                "author": "李四", 
                "text": "将技术架构图的标题设置为2级标题",
                "anchor_text": "技术架构图",
                "page_number": 3
            },
            {
                "id": "comment_3",
                "author": "王五",
                "text": "删除过时的技术栈说明段落",
                "anchor_text": "旧技术栈",
                "page_number": 4
            },
            {
                "id": "comment_4",
                "author": "赵六",
                "text": "在结论部分插入项目时间线表格",
                "anchor_text": "项目结论",
                "page_number": 7
            }
        ]
    }
    
    print(f"📊 文档信息:")
    print(f"  📁 文件: {sample_document['title']}")
    print(f"  📄 页数: {sample_document['page_count']} 页")
    print(f"  📝 字数: {sample_document['word_count']} 字")
    print(f"  💬 批注: {len(sample_document['comments'])} 个")
    print()
    
    print("💬 批注详情:")
    for i, comment in enumerate(sample_document['comments'], 1):
        print(f"  {i}. 👤 {comment['author']} (第{comment['page_number']}页)")
        print(f"     📍 位置: {comment['anchor_text']}")
        print(f"     📝 内容: {comment['text']}")
        print()
    
    # 演示工作流程分析
    print("🔍 工作流程分析:")
    print("  1. 📂 文档加载 - 使用 Word COM API 打开文档")
    print("  2. 🔍 结构分析 - 提取标题、样式、目录、超链接等信息")
    print("  3. 💬 批注解析 - 分析批注内容和位置信息")
    print("  4. 📝 提示构建 - 生成包含文档上下文的 LLM 提示词")
    print("  5. 🤖 任务规划 - 调用 LLM 生成结构化任务列表")
    print("  6. 🛡️ 格式保护 - 四重防线验证任务安全性")
    print("  7. ⚡ 任务执行 - 使用 Word COM API 执行具体操作")
    print("  8. 📊 结果报告 - 生成详细的执行报告")
    print()
    
    # 演示格式保护机制
    print("🛡️ 四重格式保护机制:")
    protection_layers = [
        {
            "layer": "第1层 - 提示词硬约束",
            "description": "LLM 系统提示词明确禁止未授权格式变更",
            "example": "'不要更改格式，除非批注明确要求'"
        },
        {
            "layer": "第2层 - 规划期过滤",
            "description": "自动过滤无批注授权的格式化任务",
            "example": "过滤掉没有 source_comment_id 的样式设置任务"
        },
        {
            "layer": "第3层 - 执行期拦截",
            "description": "执行前再次校验批注授权和任务安全性",
            "example": "验证格式化任务是否有对应的批注ID"
        },
        {
            "layer": "第4层 - 事后校验回滚",
            "description": "检测未授权变更并自动回滚到备份状态",
            "example": "检测到样式变更后自动回滚文档"
        }
    ]
    
    for i, layer in enumerate(protection_layers, 1):
        print(f"  {i}. {layer['layer']}")
        print(f"     📋 功能: {layer['description']}")
        print(f"     💡 示例: {layer['example']}")
        print()
    
    # 演示任务类型
    print("📝 支持的任务类型:")
    task_types = [
        {
            "category": "内容操作",
            "tasks": [
                {"type": "REWRITE", "desc": "重写文本内容", "auth": "无需授权"},
                {"type": "INSERT", "desc": "插入新内容", "auth": "无需授权"},
                {"type": "DELETE", "desc": "删除指定内容", "auth": "无需授权"}
            ]
        },
        {
            "category": "格式操作",
            "tasks": [
                {"type": "SET_PARAGRAPH_STYLE", "desc": "设置段落样式", "auth": "需要批注授权"},
                {"type": "SET_HEADING_LEVEL", "desc": "设置标题级别", "auth": "需要批注授权"},
                {"type": "APPLY_TEMPLATE", "desc": "应用文档模板", "auth": "需要批注授权"}
            ]
        },
        {
            "category": "结构操作",
            "tasks": [
                {"type": "REBUILD_TOC", "desc": "重建目录", "auth": "需要批注授权"},
                {"type": "REPLACE_HYPERLINK", "desc": "替换超链接", "auth": "需要批注授权"},
                {"type": "REFRESH_TOC_NUMBERS", "desc": "刷新目录页码", "auth": "系统任务"}
            ]
        }
    ]
    
    for category in task_types:
        print(f"  📂 {category['category']}:")
        for task in category['tasks']:
            auth_icon = "🔓" if "无需" in task['auth'] else "🔒" if "需要" in task['auth'] else "⚙️"
            print(f"    {auth_icon} {task['type']}: {task['desc']} ({task['auth']})")
        print()
    
    # 模拟任务生成
    print("🤖 模拟 LLM 任务生成:")
    simulated_tasks = [
        {
            "id": "task_1",
            "type": "REWRITE",
            "source_comment_id": "comment_1",
            "locator": {"by": "find", "value": "项目背景"},
            "instruction": "重写项目背景部分，增加市场分析和竞争对手分析，使内容更加详细和专业",
            "risk": "LOW"
        },
        {
            "id": "task_2",
            "type": "SET_HEADING_LEVEL",
            "source_comment_id": "comment_2",
            "locator": {"by": "heading", "value": "技术架构图"},
            "instruction": "将标题级别设置为2级",
            "risk": "MEDIUM"
        },
        {
            "id": "task_3",
            "type": "DELETE",
            "source_comment_id": "comment_3",
            "locator": {"by": "find", "value": "旧技术栈"},
            "instruction": "删除过时的技术栈说明段落",
            "risk": "LOW"
        },
        {
            "id": "task_4",
            "type": "INSERT",
            "source_comment_id": "comment_4",
            "locator": {"by": "find", "value": "项目结论"},
            "instruction": "插入项目时间线表格",
            "risk": "LOW"
        }
    ]
    
    for i, task in enumerate(simulated_tasks, 1):
        risk_icon = "🟢" if task['risk'] == 'LOW' else "🟡" if task['risk'] == 'MEDIUM' else "🔴"
        auth_status = "🔒" if task['type'] in ['SET_HEADING_LEVEL', 'SET_PARAGRAPH_STYLE'] else "🔓"
        print(f"  {i}. {risk_icon} {auth_status} {task['id']}: {task['type']}")
        print(f"     📍 定位: {task['locator']['by']} = '{task['locator']['value']}'")
        print(f"     📝 指令: {task['instruction']}")
        print(f"     💬 来源: {task['source_comment_id']}")
        print()
    
    # 演示执行结果
    print("📊 模拟执行结果:")
    execution_results = [
        {"task_id": "task_1", "success": True, "message": "重写完成，内容更加详细专业", "time": 2.5},
        {"task_id": "task_2", "success": True, "message": "标题级别已设置为2级", "time": 0.8},
        {"task_id": "task_3", "success": True, "message": "过时段落已删除", "time": 0.5},
        {"task_id": "task_4", "success": True, "message": "时间线表格已插入", "time": 1.2}
    ]
    
    total_time = sum(r["time"] for r in execution_results)
    success_count = sum(1 for r in execution_results if r["success"])
    
    for i, result in enumerate(execution_results, 1):
        status_icon = "✅" if result["success"] else "❌"
        print(f"  {i}. {status_icon} {result['task_id']}: {result['message']} ({result['time']}s)")
    
    print()
    print(f"🎉 执行完成统计:")
    print(f"  ✅ 成功任务: {success_count}/{len(execution_results)}")
    print(f"  ⏱️ 总耗时: {total_time:.1f}s")
    print(f"  📈 成功率: {success_count/len(execution_results)*100:.1f}%")
    print()
    
    print("🚀 AutoWord 让文档编辑更智能，让格式保护更安全！")


if __name__ == "__main__":
    try:
        demo_complete_workflow()
        
    except KeyboardInterrupt:
        print("\n👋 演示已停止")
    except Exception as e:
        print(f"\n❌ 演示过程中出现错误: {str(e)}")