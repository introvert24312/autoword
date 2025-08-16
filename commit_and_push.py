#!/usr/bin/env python3
"""
自动提交和推送到GitHub
"""

import subprocess
import sys
import os
from datetime import datetime

def run_command(command, description):
    """运行命令并处理错误"""
    print(f"🔄 {description}...")
    try:
        result = subprocess.run(command, shell=True, capture_output=True, text=True, encoding='utf-8')
        if result.returncode == 0:
            print(f"✅ {description} 成功")
            if result.stdout.strip():
                print(f"   输出: {result.stdout.strip()}")
            return True
        else:
            print(f"❌ {description} 失败")
            if result.stderr.strip():
                print(f"   错误: {result.stderr.strip()}")
            return False
    except Exception as e:
        print(f"❌ {description} 异常: {str(e)}")
        return False

def main():
    """主函数"""
    print("=== AutoWord 项目自动提交和推送 ===")
    print()
    
    # 检查Git状态
    if not run_command("git status", "检查Git状态"):
        print("❌ Git状态检查失败，请确保在Git仓库中")
        return False
    
    # 添加所有文件
    if not run_command("git add .", "添加所有文件"):
        print("❌ 添加文件失败")
        return False
    
    # 创建提交信息
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    commit_message = f"""🚀 AutoWord 文档自动化系统完整实现

✨ 核心功能:
- 🤖 LLM客户端 (GPT-5/Claude 3.7)
- 📝 智能提示词构建器
- 🎯 任务规划器 (四重格式保护)
- ⚡ Word COM执行器
- 🛡️ 格式验证器
- 🔗 TOC和超链接管理
- 📊 日志导出系统
- 🔄 文档验证和回滚
- 🚀 增强执行器 (主入口)

🧪 测试覆盖:
- 单元测试: 200+ 测试用例
- 集成测试: 完整工作流程
- 模拟测试: 所有外部依赖

📚 文档和示例:
- 完整的README文档
- 功能演示脚本
- API使用示例

🛡️ 安全特性:
- 四重格式保护机制
- 自动备份和回滚
- 批注授权验证
- 未授权变更检测

提交时间: {timestamp}
"""
    
    # 提交更改
    if not run_command(f'git commit -m "{commit_message}"', "提交更改"):
        print("❌ 提交失败")
        return False
    
    # 推送到远程仓库
    if not run_command("git push origin main", "推送到GitHub"):
        print("❌ 推送失败，尝试推送到master分支")
        if not run_command("git push origin master", "推送到GitHub (master)"):
            print("❌ 推送到master分支也失败")
            return False
    
    print()
    print("🎉 AutoWord 项目已成功提交并推送到GitHub!")
    print("💰 赚钱系统已就绪，可以开始商业化运营!")
    print()
    print("📋 项目特点:")
    print("  🚀 完整的文档自动化解决方案")
    print("  🛡️ 企业级安全和格式保护")
    print("  🤖 最新LLM技术集成")
    print("  📊 详细的执行报告和日志")
    print("  🔧 灵活的配置和扩展性")
    print()
    print("💡 商业化建议:")
    print("  📈 SaaS服务: 按文档处理量收费")
    print("  🏢 企业版: 私有部署和定制开发")
    print("  🎓 培训服务: 文档自动化咨询")
    print("  🔌 API服务: 集成到其他系统")
    print()
    print("🔗 项目地址: https://github.com/your-repo/autoword")
    
    return True

if __name__ == "__main__":
    try:
        success = main()
        sys.exit(0 if success else 1)
    except KeyboardInterrupt:
        print("\n👋 操作已取消")
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ 发生异常: {str(e)}")
        sys.exit(1)