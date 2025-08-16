"""
AutoWord Enhanced Executor Demo
增强执行器演示 - 完整的四重格式保护防线
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from autoword.core.enhanced_executor import EnhancedWordExecutor, execute_with_full_protection
from autoword.core.models import (
    Task, TaskType, RiskLevel, LocatorType, Locator, Comment
)


def demo_four_layer_protection():
    """演示四重格式保护防线"""
    print("=== 四重格式保护防线演示 ===\n")
    
    print("🛡️ AutoWord 四重格式保护防线:")
    print("  1️⃣ 提示词硬约束 - LLM 系统提示词明确禁止未授权格式变更")
    print("  2️⃣ 规划期过滤 - 过滤无批注来源的格式类任务")
    print("  3️⃣ 执行期拦截 - 执行前再次校验批注授权")
    print("  4️⃣ 事后校验回滚 - 检测未授权变更并自动回滚")
    
    print(f"\n🔒 保护范围:")
    print("  - 标题级别和样式变更")
    print("  - 段落样式设置")
    print("  - 模板应用")
    print("  - 目录结构变更")
    print("  - 超链接地址修改")
    
    print(f"\n✅ 允许的无授权操作:")
    print("  - 内容重写 (rewrite)")
    print("  - 内容插入 (insert)")
    print("  - 内容删除 (delete)")
    print("  - 目录页码刷新 (refresh_toc_numbers)")


def demo_enhanced_task_types():
    """演示增强的任务类型"""
    print("\n=== 增强任务类型演示 ===\n")
    
    # 创建各种增强任务
    enhanced_tasks = [
        {
            "task": Task(
                id="format_task_1",
                type=TaskType.APPLY_TEMPLATE,
                source_comment_id="comment_1",
                locator=Locator(by=LocatorType.FIND, value="整个文档"),
                instruction="应用模板: 公司标准模板",
                risk=RiskLevel.HIGH
            ),
            "description": "应用文档模板 - 高风险格式化操作"
        },
        {
            "task": Task(
                id="format_task_2",
                type=TaskType.REBUILD_TOC,
                source_comment_id="comment_2",
                locator=Locator(by=LocatorType.HEADING, value="目录"),
                instruction="重建目录，包含1-3级标题",
                risk=RiskLevel.HIGH
            ),
            "description": "重建目录 - 高风险结构化操作"
        },
        {
            "task": Task(
                id="format_task_3",
                type=TaskType.UPDATE_TOC_LEVELS,
                source_comment_id="comment_3",
                locator=Locator(by=LocatorType.FIND, value="目录"),
                instruction="更新目录级别为1-4级",
                risk=RiskLevel.MEDIUM
            ),
            "description": "更新目录级别 - 中等风险操作"
        },
        {
            "task": Task(
                id="safe_task_1",
                type=TaskType.REFRESH_TOC_NUMBERS,
                source_comment_id=None,  # 允许无授权
                locator=Locator(by=LocatorType.FIND, value="目录"),
                instruction="刷新目录页码",
                risk=RiskLevel.LOW
            ),
            "description": "刷新目录页码 - 安全操作，无需授权"
        }
    ]
    
    print("📋 增强任务类型:")
    for item in enhanced_tasks:
        task = item["task"]
        risk_icon = {"low": "🟢", "medium": "🟡", "high": "🔴"}[task.risk.value]
        auth_status = "🔓 需要授权" if task.source_comment_id else "🔒 无需授权"
        
        print(f"\n  {risk_icon} {task.type.value}")
        print(f"     {item['description']}")
        print(f"     {auth_status}")
        print(f"     指令: {task.instruction}")


def demo_format_validation():
    """演示格式验证功能"""
    print("\n=== 格式验证功能演示 ===\n")
    
    print("🔍 第4层防线 - 事后校验回滚:")
    print("  执行完成后自动比较文档前后状态，检测所有格式变更")
    
    print(f"\n📊 检测的变更类型:")
    change_types = [
        ("heading_level_change", "标题级别变更", "检测标题从1级改为2级等变更"),
        ("heading_style_change", "标题样式变更", "检测标题样式的改变"),
        ("style_usage_change", "样式使用变更", "检测样式从未使用变为使用"),
        ("toc_structure_change", "目录结构变更", "检测目录条目数量变化"),
        ("toc_levels_change", "目录级别变更", "检测目录级别分布变化"),
        ("hyperlink_address_change", "超链接地址变更", "检测链接地址的修改")
    ]
    
    for change_type, name, description in change_types:
        print(f"  • {name}: {description}")
    
    print(f"\n⚖️ 授权验证流程:")
    print("  1. 识别所有格式变更")
    print("  2. 查找对应的授权任务")
    print("  3. 验证任务是否有批注授权")
    print("  4. 标记未授权变更")
    print("  5. 生成验证报告")
    print("  6. 自动回滚未授权变更")


def demo_rollback_mechanism():
    """演示回滚机制"""
    print("\n=== 自动回滚机制演示 ===\n")
    
    print("🔄 回滚触发条件:")
    print("  - 检测到未授权的格式变更")
    print("  - 执行过程中发生严重错误")
    print("  - 用户手动触发回滚")
    
    print(f"\n💾 备份策略:")
    print("  - 执行前自动创建带时间戳的备份文件")
    print("  - 备份文件命名格式: 原文件名_backup_YYYYMMDD_HHMMSS.docx")
    print("  - 支持多版本备份管理")
    
    print(f"\n🔧 回滚过程:")
    print("  1. 检测到需要回滚的情况")
    print("  2. 记录详细的回滚原因")
    print("  3. 从备份文件恢复原始文档")
    print("  4. 更新执行结果状态")
    print("  5. 生成回滚报告")
    
    print(f"\n📋 回滚报告内容:")
    print("  - 回滚原因和触发条件")
    print("  - 被回滚的具体变更")
    print("  - 恢复的文档版本信息")
    print("  - 建议的后续操作")


def demo_enhanced_execution_workflow():
    """演示增强执行工作流程"""
    print("\n=== 增强执行工作流程演示 ===\n")
    
    print("🔄 完整执行流程:")
    steps = [
        ("1. 预检查", "验证文档路径、Word COM 可用性"),
        ("2. 创建备份", "生成带时间戳的备份文件"),
        ("3. 执行前快照", "记录文档当前状态"),
        ("4. 任务执行", "使用基础执行器执行所有任务"),
        ("5. 执行后快照", "记录文档修改后状态"),
        ("6. 格式验证", "比较前后状态，检测未授权变更"),
        ("7. 自动回滚", "如发现问题自动恢复到原始状态"),
        ("8. 生成报告", "创建详细的执行和验证报告")
    ]
    
    for step, description in steps:
        print(f"  {step}: {description}")
    
    print(f"\n⚡ 性能优化:")
    print("  - 智能快照：只记录关键结构信息")
    print("  - 增量比较：只检测实际发生的变更")
    print("  - 并行验证：格式验证与任务执行并行")
    print("  - 缓存机制：重复操作使用缓存结果")


def demo_protection_status():
    """演示保护状态查询"""
    print("\n=== 格式保护状态查询演示 ===\n")
    
    # 创建增强执行器
    executor = EnhancedWordExecutor(enable_validation=True)
    
    # 获取保护状态
    status = executor.get_protection_status()
    
    print("🛡️ 当前格式保护状态:")
    for layer, status_text in status["four_layer_protection"].items():
        layer_num = layer.replace("layer_", "第") + "层"
        print(f"  {layer_num}: {status_text}")
    
    print(f"\n⚙️ 配置状态:")
    print(f"  - 格式验证: {'启用' if status['validation_enabled'] else '禁用'}")
    print(f"  - 自动备份: {'启用' if status['auto_backup'] else '禁用'}")
    print(f"  - 回滚能力: {'支持' if status['rollback_capability'] else '不支持'}")


def demo_usage_examples():
    """演示使用示例"""
    print("\n=== 使用示例演示 ===\n")
    
    print("📖 基本使用示例:")
    print("```python")
    print("from autoword.core.enhanced_executor import execute_with_full_protection")
    print("")
    print("# 使用完整格式保护执行任务")
    print("result = execute_with_full_protection(")
    print("    tasks=tasks,")
    print("    document_path='document.docx',")
    print("    comments=comments,")
    print("    visible=False,")
    print("    dry_run=False,")
    print("    auto_rollback=True")
    print(")")
    print("")
    print("# 检查执行结果")
    print("if result.success:")
    print("    print(f'执行成功: {result.completed_tasks}/{result.total_tasks}')")
    print("    if result.validation_report:")
    print("        print(f'格式验证: {\"通过\" if result.validation_report.is_valid else \"失败\"}')")
    print("else:")
    print("    print(f'执行失败: {result.error_summary}')")
    print("    if result.rollback_performed:")
    print("        print('已自动回滚到原始状态')")
    print("```")
    
    print(f"\n🔍 试运行示例:")
    print("```python")
    print("# 试运行预验证")
    print("dry_result = execute_with_full_protection(")
    print("    tasks=tasks,")
    print("    document_path='document.docx',")
    print("    comments=comments,")
    print("    dry_run=True")
    print(")")
    print("")
    print("if dry_result.success:")
    print("    print('试运行通过，可以正式执行')")
    print("else:")
    print("    print(f'试运行发现问题: {dry_result.error_summary}')")
    print("```")


if __name__ == "__main__":
    print("=== AutoWord 增强执行器演示 ===\n")
    
    demo_four_layer_protection()
    demo_enhanced_task_types()
    demo_format_validation()
    demo_rollback_mechanism()
    demo_enhanced_execution_workflow()
    demo_protection_status()
    demo_usage_examples()
    
    print("\n=== 演示完成 ===")
    print("🛡️ 四重格式保护防线确保文档安全")
    print("⚡ 增强执行器提供完整的自动化解决方案")
    print("🔄 自动回滚机制保障操作可逆性")
    print("📊 详细报告帮助分析和优化执行过程")