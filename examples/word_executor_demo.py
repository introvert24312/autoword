"""
AutoWord Word Executor Demo
Word COM 执行器使用示例
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from autoword.core.word_executor import WordExecutor, ExecutionMode, execute_task_list
from autoword.core.models import (
    Task, TaskType, RiskLevel, LocatorType, Locator, Comment
)


def demo_task_creation():
    """演示任务创建"""
    print("=== Word 执行器任务创建演示 ===\n")
    
    # 创建各种类型的任务
    tasks = [
        Task(
            id="task_1",
            type=TaskType.REWRITE,
            source_comment_id="comment_1",
            locator=Locator(by=LocatorType.FIND, value="项目背景"),
            instruction="重写项目背景部分，使其更加详细和具体",
            risk=RiskLevel.LOW,
            notes="基于张三的批注"
        ),
        Task(
            id="task_2",
            type=TaskType.INSERT,
            source_comment_id="comment_2",
            locator=Locator(by=LocatorType.HEADING, value="技术方案"),
            instruction="在技术方案章节后插入架构图说明",
            risk=RiskLevel.LOW,
            notes="基于李四的批注"
        ),
        Task(
            id="task_3",
            type=TaskType.SET_HEADING_LEVEL,
            source_comment_id="comment_3",
            locator=Locator(by=LocatorType.RANGE, value="100-120"),
            instruction="将此标题设置为2级标题",
            risk=RiskLevel.MEDIUM,
            notes="基于王五的批注"
        ),
        Task(
            id="task_4",
            type=TaskType.DELETE,
            source_comment_id="comment_4",
            locator=Locator(by=LocatorType.FIND, value="过时的内容"),
            instruction="删除这段过时的内容",
            risk=RiskLevel.LOW,
            notes="基于赵六的批注"
        ),
        Task(
            id="task_5",
            type=TaskType.SET_PARAGRAPH_STYLE,
            source_comment_id="comment_5",
            locator=Locator(by=LocatorType.BOOKMARK, value="important_section"),
            instruction="将此段落设置为强调样式",
            risk=RiskLevel.MEDIUM,
            notes="基于孙七的批注"
        )
    ]
    
    print(f"📋 创建了 {len(tasks)} 个任务:")
    for task in tasks:
        risk_icon = {"low": "🟢", "medium": "🟡", "high": "🔴"}[task.risk.value]
        print(f"  {risk_icon} {task.id}: {task.type.value}")
        print(f"     定位: {task.locator.by.value} = '{task.locator.value}'")
        print(f"     指令: {task.instruction}")
        print(f"     批注: {task.source_comment_id}")
        print()
    
    return tasks


def demo_execution_modes():
    """演示执行模式"""
    print("=== 执行模式演示 ===\n")
    
    print("🔧 支持的执行模式:")
    print("  1. 🟢 NORMAL - 正常执行模式，实际修改文档")
    print("  2. 🔍 DRY_RUN - 试运行模式，不实际修改文档，仅模拟执行")
    print("  3. 🛡️ SAFE - 安全模式，额外验证和保护措施")
    
    print(f"\n📊 执行器特性:")
    print("  - 🎯 精确定位: 支持书签、范围、标题、文本查找四种定位方式")
    print("  - 🛡️ 格式保护: 第3层防线 - 执行期拦截未授权格式变更")
    print("  - 🔄 错误恢复: 单个任务失败不影响其他任务执行")
    print("  - 📝 详细日志: 记录每个任务的执行状态和耗时")
    print("  - 💾 自动保存: 执行完成后自动保存文档")


def demo_task_locators():
    """演示任务定位器"""
    print("\n=== 任务定位器演示 ===\n")
    
    locator_examples = [
        {
            "type": "书签定位",
            "locator": Locator(by=LocatorType.BOOKMARK, value="section_1"),
            "description": "通过书签名称精确定位",
            "use_case": "适用于预先标记的重要位置"
        },
        {
            "type": "范围定位",
            "locator": Locator(by=LocatorType.RANGE, value="100-200"),
            "description": "通过字符位置范围定位",
            "use_case": "适用于已知精确位置的文本"
        },
        {
            "type": "标题定位",
            "locator": Locator(by=LocatorType.HEADING, value="第一章 概述"),
            "description": "通过标题文本定位",
            "use_case": "适用于结构化文档的章节操作"
        },
        {
            "type": "文本查找",
            "locator": Locator(by=LocatorType.FIND, value="项目背景介绍"),
            "description": "通过文本内容查找定位",
            "use_case": "适用于内容相关的修改操作"
        }
    ]
    
    print("🎯 支持的定位方式:")
    for example in locator_examples:
        print(f"\n  📍 {example['type']}")
        print(f"     定位器: {example['locator'].by.value} = '{example['locator'].value}'")
        print(f"     说明: {example['description']}")
        print(f"     适用场景: {example['use_case']}")


def demo_format_protection():
    """演示格式保护功能"""
    print("\n=== 格式保护演示 ===\n")
    
    print("🛡️ 第3层防线 - 执行期拦截:")
    print("  执行器在执行每个任务前会进行安全检查:")
    
    # 模拟批注
    comments = [
        Comment(
            id="comment_1",
            author="张三",
            page=1,
            anchor_text="标题格式",
            comment_text="将这个标题改为2级标题",
            range_start=100,
            range_end=120
        )
    ]
    
    # 授权的格式化任务
    authorized_task = Task(
        id="authorized_task",
        type=TaskType.SET_HEADING_LEVEL,
        source_comment_id="comment_1",  # 有批注授权
        locator=Locator(by=LocatorType.RANGE, value="100-120"),
        instruction="设置为2级标题"
    )
    
    # 未授权的格式化任务
    unauthorized_task = Task(
        id="unauthorized_task",
        type=TaskType.SET_PARAGRAPH_STYLE,
        source_comment_id=None,  # 无批注授权
        locator=Locator(by=LocatorType.FIND, value="某段落"),
        instruction="设置段落样式"
    )
    
    print(f"\n  ✅ 授权任务示例:")
    print(f"     任务: {authorized_task.type.value}")
    print(f"     批注ID: {authorized_task.source_comment_id}")
    print(f"     状态: 通过格式保护检查")
    
    print(f"\n  ❌ 未授权任务示例:")
    print(f"     任务: {unauthorized_task.type.value}")
    print(f"     批注ID: {unauthorized_task.source_comment_id}")
    print(f"     状态: 被格式保护阻止")
    
    print(f"\n  🔒 保护机制:")
    print("     1. 检查格式化任务是否有对应的批注授权")
    print("     2. 验证批注ID是否存在于批注列表中")
    print("     3. 阻止所有未授权的格式变更操作")
    print("     4. 允许内容类操作（重写、插入、删除）无需授权")


def demo_execution_workflow():
    """演示执行工作流程"""
    print("\n=== 执行工作流程演示 ===\n")
    
    print("🔄 Word 执行器工作流程:")
    print("  1. 📂 打开文档 - 使用 Word COM 打开目标文档")
    print("  2. 🛡️ 安全检查 - 对每个任务进行格式保护验证")
    print("  3. 🎯 任务定位 - 使用定位器精确找到目标位置")
    print("  4. ⚡ 执行操作 - 根据任务类型执行相应的 Word 操作")
    print("  5. 📊 记录结果 - 记录执行状态、耗时和错误信息")
    print("  6. 💾 保存文档 - 自动保存修改后的文档")
    print("  7. 📋 生成报告 - 返回详细的执行结果报告")
    
    print(f"\n⚡ 性能特性:")
    print("  - 🔀 并发安全: 单线程 COM 操作确保稳定性")
    print("  - 🔄 错误隔离: 单个任务失败不影响其他任务")
    print("  - 📈 进度跟踪: 实时报告执行进度和状态")
    print("  - 🎛️ 模式切换: 支持正常、试运行、安全三种模式")
    print("  - 🔍 详细日志: 完整记录每个操作的详细信息")


def demo_error_handling():
    """演示错误处理"""
    print("\n=== 错误处理演示 ===\n")
    
    print("🚨 错误处理策略:")
    print("  1. 🎯 定位失败 - 自动尝试模糊匹配和备选方案")
    print("  2. 🛡️ 格式保护 - 阻止未授权操作并记录详细原因")
    print("  3. 📄 COM 异常 - 捕获 Word COM 错误并提供友好提示")
    print("  4. 🔄 任务隔离 - 单个任务失败不中断整体执行")
    print("  5. 📊 详细报告 - 提供完整的错误信息和建议")
    
    print(f"\n🔧 恢复机制:")
    print("  - 📍 定位失败时自动尝试模糊匹配")
    print("  - 🔍 找不到目标时使用文档开始位置")
    print("  - 📝 记录所有失败原因供后续分析")
    print("  - 🎛️ 试运行模式可预先检测潜在问题")


if __name__ == "__main__":
    print("=== AutoWord Word 执行器演示 ===\n")
    
    tasks = demo_task_creation()
    demo_execution_modes()
    demo_task_locators()
    demo_format_protection()
    demo_execution_workflow()
    demo_error_handling()
    
    print("\n=== 演示完成 ===")
    print("💡 提示: Word 执行器已准备好执行任务")
    print("🔒 安全保障: 第3层格式保护防线确保执行安全")
    print("⚡ 高性能: 优化的 COM 操作和错误处理机制")
    
    print(f"\n📖 使用示例:")
    print("```python")
    print("from autoword.core.word_executor import execute_task_list")
    print("")
    print("# 执行任务列表")
    print("result = execute_task_list(")
    print("    tasks=tasks,")
    print("    document_path='document.docx',")
    print("    visible=False,")
    print("    dry_run=False")
    print(")")
    print("")
    print("print(f'执行结果: {result.success}')")
    print("print(f'完成任务: {result.completed_tasks}/{result.total_tasks}')")
    print("```")