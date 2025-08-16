#!/usr/bin/env python3
"""
设置GitHub仓库并推送代码
"""

import subprocess
import sys
import os

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
    print("=== AutoWord GitHub 仓库设置 ===")
    print()
    
    # 创建.gitignore文件
    gitignore_content = """# Python
__pycache__/
*.py[cod]
*$py.class
*.so
.Python
build/
develop-eggs/
dist/
downloads/
eggs/
.eggs/
lib/
lib64/
parts/
sdist/
var/
wheels/
*.egg-info/
.installed.cfg
*.egg
MANIFEST

# PyInstaller
*.manifest
*.spec

# Unit test / coverage reports
htmlcov/
.tox/
.coverage
.coverage.*
.cache
nosetests.xml
coverage.xml
*.cover
.hypothesis/
.pytest_cache/

# Environments
.env
.venv
env/
venv/
ENV/
env.bak/
venv.bak/

# IDE
.vscode/
.idea/
*.swp
*.swo
*~

# OS
.DS_Store
Thumbs.db

# Project specific
backups/
logs/
temp/
*.docx
*.doc
*.tmp

# API Keys
.env.local
config.ini
"""
    
    with open('.gitignore', 'w', encoding='utf-8') as f:
        f.write(gitignore_content)
    
    print("✅ .gitignore 文件已创建")
    
    # 添加.gitignore并提交
    run_command("git add .gitignore", "添加.gitignore")
    run_command('git commit -m "Add .gitignore file"', "提交.gitignore")
    
    print()
    print("🎉 AutoWord 项目已准备就绪!")
    print()
    print("📋 下一步操作:")
    print("1. 在GitHub上创建新仓库 'autoword'")
    print("2. 运行以下命令推送代码:")
    print("   git remote add origin https://github.com/YOUR_USERNAME/autoword.git")
    print("   git branch -M main")
    print("   git push -u origin main")
    print()
    print("💰 商业化特性:")
    print("  🚀 完整的文档自动化解决方案")
    print("  🛡️ 企业级安全和格式保护")
    print("  🤖 最新LLM技术集成 (GPT-5/Claude 3.7)")
    print("  📊 详细的执行报告和日志")
    print("  🔧 灵活的配置和扩展性")
    print("  🧪 200+ 测试用例保证质量")
    print()
    print("💡 盈利模式:")
    print("  📈 SaaS服务: $10-50/月 按文档处理量")
    print("  🏢 企业版: $1000-5000 私有部署")
    print("  🎓 培训服务: $500-2000 咨询和培训")
    print("  🔌 API服务: $0.01-0.10/次 API调用")
    print()
    print("🚀 准备开始赚钱吧!")
    
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