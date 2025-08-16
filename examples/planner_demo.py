"""
AutoWord Task Planner Demo
任务规划器使用示例
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import json
from autoword.core.planner import TaskPlanner, FormatProtectionGuard, RiskAssessment
from autoword.core.models import (
    Comment, DocumentStructure, Heading, Style, TaskType, RiskLevel
)
from autoword.core.llm_client import ModelType


def demo_format_protection():
    """演示格式保护功能"""
    print("=== 格式保护守卫演示 ===\n")
    
    guard = FormatProtectionGuard()
    
    # 测试任务数据
    test_tasks = [
        {
            "id": "task_1",
            "type": "rewrite",
            "source_comment_id": None,
            "instruction": "重写这段内容"
        },
        {
            "id": "task_2",
            "type": "set_paragraph_style",
            "source_comment_id": "comment_1",
            "instruction": "设置段落样式为标题1"
        },
        {
            "id": "task_3",
            "type": "set_paragraph_style",
            "source_comment_id": None,  # 无授权
            "instruction": "设置段落样式（未授权）"
        },
        {
            "id": "task_4",
            "type": "insert",
            "source_comment_id": None,
            "instruction": "插入新段落"
        },
        {
            "id": "task_5",
            "type": "apply_template",
            "source_comment_id": None,  # 无授权
            "instruction": "应用模板（未授权）"
        }
    ]
    
    print(f"📋 原始任务: {len(test_tasks)} 个")
    for task in test_tasks:
        print(f"  - {task['id']}: {task['type']} ({'有授权' if task['source_comment_id'] else '无授权'})")
    
    # 应用格式保护
    authorized, filtered = guard.filter_unauthorized_tasks(test_tasks)
    
    print(f"\n✅ 授权任务: {len(authorized)} 个")
    for task in authorized:
        print(f"  - {task['id']}: {task['type']}")
    
    print(f"\n❌ 被过滤任务: {len(filtered)} 个")
    for task in filtered:
        print(f"  - {task['id']}: {task['type']} - {task['filter_reason']}")
    
    print(f"\n🛡️ 格式保护效果: 阻止了 {len(filtered)} 个未授权的格式化任务")


def demo_risk_assessment():
    """演示风险评估功能"""
    print("\n=== 风险评估器演示 ===\n")
    
    assessor = RiskAssessment()
    
    # 测试任务数据
    test_tasks = [
        {"type": "rewrite"},
        {"type": "insert"},
        {"type": "delete"},
        {"type": "set_paragraph_style"},
        {"type": "set_heading_level"},
        {"type": "apply_template"},
        {"type": "rebuild_toc"},
        {"type": "replace_hyperlink"},
        {"type": "refresh_toc_numbers"}
    ]
    
    print("📊 单个任务风险评估:")
    for task in test_tasks:
        risk = assessor.assess_task_risk(task)
        risk_icon = {"low": "🟢", "medium": "🟡", "high": "🔴"}[risk.value]
        print(f"  {risk_icon} {task['type']}: {risk.value.upper()}")
    
    # 批量风险评估
    print(f"\n📈 批量风险评估:")
    risk_report = assessor.assess_batch_risk(test_tasks)
    
    print(f"  整体风险级别: {risk_report['overall_risk'].value.upper()}")
    print(f"  风险分布:")
    print(f"    - 低风险: {risk_report['risk_distribution']['low']} 个")
    print(f"    - 中等风险: {risk_report['risk_distribution']['medium']} 个")
    print(f"    - 高风险: {risk_report['risk_distribution']['high']} 个")
    print(f"  高风险比例: {risk_report['high_risk_ratio']:.1%}")
    
    if risk_report['recommendations']:
        print(f"  💡 建议:")
        for rec in risk_report['recommendations']:
            print(f"    - {rec}")


def demo_task_planning_simulation():
    """演示任务规划模拟（不调用真实LLM）"""
    print("\n=== 任务规划模拟演示 ===\n")
    
    # 创建模拟的文档结构
    structure = DocumentStructure(
        headings=[
            Heading(level=1, text="第一章 概述", style="标题 1", range_start=0, range_end=20),
            Heading(level=2, text="1.1 背景", style="标题 2", range_start=100, range_end=120),
            Heading(level=2, text="1.2 目标", style="标题 2", range_start=200, range_end=220),
        ],
        styles=[
            Style(name="标题 1", type="paragraph", built_in=True, in_use=True),
            Style(name="标题 2", type="paragraph", built_in=True, in_use=True),
            Style(name="正文", type="paragraph", built_in=True, in_use=True),
        ],
        toc_entries=[],
        hyperlinks=[],
        references=[],
        page_count=5,
        word_count=1500
    )
    
    # 创建模拟批注
    comments = [
        Comment(
            id="comment_1",
            author="张三",
            page=1,
            anchor_text="项目背景部分内容过于简单",
            comment_text="这里需要补充更详细的背景信息，建议增加2-3段内容",
            range_start=110,
            range_end=150
        ),
        Comment(
            id="comment_2",
            author="李四",
            page=2,
            anchor_text="目标描述不够清晰",
            comment_text="重新组织这段文字，使目标更加明确和具体",
            range_start=210,
            range_end=250
        ),
        Comment(
            id="comment_3",
            author="王五",
            page=1,
            anchor_text="标题格式",
            comment_text="将这个标题改为2级标题",
            range_start=10,
            range_end=30
        )
    ]
    
    print(f"📄 文档信息:")
    print(f"  - 页数: {structure.page_count}")
    print(f"  - 字数: {structure.word_count}")
    print(f"  - 标题: {len(structure.headings)} 个")
    print(f"  - 批注: {len(comments)} 个")
    
    # 模拟 LLM 响应
    mock_llm_response = {
        "tasks": [
            {
                "id": "task_1",
                "type": "rewrite",
                "source_comment_id": "comment_1",
                "locator": {"by": "range", "value": "110-150"},
                "instruction": "重写背景部分，增加2-3段详细内容",
                "risk": "low",
                "notes": "基于张三的批注"
            },
            {
                "id": "task_2", 
                "type": "rewrite",
                "source_comment_id": "comment_2",
                "locator": {"by": "range", "value": "210-250"},
                "instruction": "重新组织目标描述，使其更加明确具体",
                "risk": "low",
                "notes": "基于李四的批注"
            },
            {
                "id": "task_3",
                "type": "set_heading_level",
                "source_comment_id": "comment_3",
                "locator": {"by": "range", "value": "10-30"},
                "instruction": "将标题级别设置为2级",
                "risk": "medium",
                "notes": "基于王五的批注"
            },
            {
                "id": "task_4",
                "type": "set_paragraph_style",
                "source_comment_id": None,  # 无授权，应被过滤
                "locator": {"by": "find", "value": "某段落"},
                "instruction": "设置段落样式",
                "risk": "medium"
            },
            {
                "id": "task_5",
                "type": "refresh_toc_numbers",
                "source_comment_id": None,  # 允许无授权
                "locator": {"by": "heading", "value": "目录"},
                "instruction": "刷新目录页码",
                "risk": "low"
            }
        ]
    }
    
    print(f"\n🤖 模拟 LLM 生成任务: {len(mock_llm_response['tasks'])} 个")
    
    # 应用格式保护
    guard = FormatProtectionGuard()
    authorized, filtered = guard.filter_unauthorized_tasks(mock_llm_response['tasks'])
    
    print(f"\n🛡️ 格式保护结果:")
    print(f"  - 授权任务: {len(authorized)} 个")
    print(f"  - 被过滤任务: {len(filtered)} 个")
    
    # 风险评估
    assessor = RiskAssessment()
    risk_report = assessor.assess_batch_risk(authorized)
    
    print(f"\n📊 风险评估结果:")
    print(f"  - 整体风险: {risk_report['overall_risk'].value.upper()}")
    print(f"  - 低风险: {risk_report['risk_distribution']['low']} 个")
    print(f"  - 中等风险: {risk_report['risk_distribution']['medium']} 个")
    print(f"  - 高风险: {risk_report['risk_distribution']['high']} 个")
    
    # 显示最终任务列表
    print(f"\n📋 最终任务列表:")
    for i, task in enumerate(authorized, 1):
        risk_icon = {"low": "🟢", "medium": "🟡", "high": "🔴"}[task.get('risk', 'low')]
        comment_info = f" (批注: {task['source_comment_id']})" if task.get('source_comment_id') else " (系统任务)"
        print(f"  {i}. {risk_icon} {task['id']}: {task['type']}{comment_info}")
        print(f"     指令: {task['instruction']}")
    
    if filtered:
        print(f"\n❌ 被过滤的任务:")
        for task in filtered:
            print(f"  - {task['id']}: {task['type']} - {task['filter_reason']}")


def demo_planning_workflow():
    """演示完整的规划工作流程"""
    print("\n=== 完整规划工作流程演示 ===\n")
    
    print("🔄 任务规划工作流程:")
    print("  1. 📥 接收文档结构和批注")
    print("  2. 🏗️ 构建 LLM 提示词")
    print("  3. 🤖 调用 LLM 生成任务")
    print("  4. 🛡️ 应用格式保护过滤")
    print("  5. 📊 进行风险评估")
    print("  6. 🔗 解析任务依赖关系")
    print("  7. 📋 生成最终任务计划")
    print("  8. ✅ 验证任务安全性")
    
    print(f"\n🛡️ 四重格式保护防线:")
    print("  1. 🎯 提示词硬约束: LLM 系统提示词明确禁止未授权格式变更")
    print("  2. 🔍 规划期过滤: 过滤无批注来源的格式类任务")
    print("  3. 🚫 执行期拦截: 执行前再次校验批注授权")
    print("  4. 🔄 事后校验回滚: 检测未授权变更并自动回滚")
    
    print(f"\n⚡ 性能优化特性:")
    print("  - 🧠 智能上下文管理: 自动检测和处理上下文溢出")
    print("  - 📦 分块处理: 大文档按标题或批注智能分块")
    print("  - 🔄 重试机制: LLM 调用失败自动重试")
    print("  - 📈 风险评估: 智能评估任务风险并提供建议")
    print("  - 🔗 依赖解析: 自动解析任务依赖关系并排序")


if __name__ == "__main__":
    demo_format_protection()
    demo_risk_assessment()
    demo_task_planning_simulation()
    demo_planning_workflow()
    
    print("\n=== 演示完成 ===")
    print("💡 提示: 任务规划器已准备好与 Word COM 执行器集成使用")
    print("🔒 安全保障: 四重格式保护防线确保文档格式安全")