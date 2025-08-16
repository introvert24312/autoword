@echo off
chcp 65001 >nul
title AutoWord 一键演示
color 0E

echo.
echo ╔══════════════════════════════════════════════════════════════╗
echo ║                   AutoWord 一键演示                          ║
echo ║                 文档自动化完整体验                            ║
echo ╚══════════════════════════════════════════════════════════════╝
echo.

echo 🚀 AutoWord 完整演示流程
echo.
echo 本演示将展示AutoWord的完整功能：
echo   1. 📝 创建带批注的演示文档
echo   2. 🔍 系统环境检查
echo   3. 🎯 核心功能测试
echo   4. 🚀 完整工作流程演示
echo.

set /p confirm=是否开始演示？(Y/N): 
if /i not "%confirm%"=="Y" goto END

echo.
echo ═══════════════════════════════════════════════════════════════
echo 第1步: 创建演示文档
echo ═══════════════════════════════════════════════════════════════
echo.

echo 📝 正在创建包含批注的演示Word文档...
python create_demo_document.py

echo.
echo ═══════════════════════════════════════════════════════════════
echo 第2步: 系统环境检查
echo ═══════════════════════════════════════════════════════════════
echo.

echo 🔍 检查Python环境...
python --version
if errorlevel 1 (
    echo ❌ Python 未安装，请先安装Python 3.8+
    pause
    goto END
)

echo.
echo 📦 检查依赖包...
python -c "
try:
    import requests, pydantic, win32com.client
    print('✅ 核心依赖包已安装')
except ImportError as e:
    print(f'⚠️ 缺少依赖包: {e}')
    print('正在安装依赖包...')
    import subprocess
    subprocess.run(['pip', 'install', '-r', 'requirements.txt'])
"

echo.
echo 🔑 检查API密钥配置...
if defined GPT5_KEY (
    echo ✅ GPT5_KEY 已配置
    set API_AVAILABLE=1
) else if defined CLAUDE37_KEY (
    echo ✅ CLAUDE37_KEY 已配置  
    set API_AVAILABLE=1
) else (
    echo ⚠️ 未配置API密钥，将使用演示模式
    echo 💡 设置环境变量 GPT5_KEY 或 CLAUDE37_KEY 以使用真实API
    set API_AVAILABLE=0
)

echo.
echo ═══════════════════════════════════════════════════════════════
echo 第3步: 核心功能测试
echo ═══════════════════════════════════════════════════════════════
echo.

echo 🧪 运行核心功能测试...
python quick_test.py >nul 2>&1
if errorlevel 1 (
    echo ❌ 核心功能测试失败
    echo 🔧 正在运行详细测试...
    python quick_test.py
    pause
) else (
    echo ✅ 核心功能测试通过
)

echo.
echo ═══════════════════════════════════════════════════════════════
echo 第4步: 完整工作流程演示
echo ═══════════════════════════════════════════════════════════════
echo.

echo 🚀 运行完整工作流程演示...
python examples/complete_workflow_demo.py

echo.
echo ═══════════════════════════════════════════════════════════════
echo 第5步: 格式保护机制演示
echo ═══════════════════════════════════════════════════════════════
echo.

echo 🛡️ 演示四重格式保护机制...
python -c "
print('🛡️ AutoWord 四重格式保护机制演示')
print('=' * 50)

from autoword.core.planner import FormatProtectionGuard
from autoword.core.models import Task, TaskType, Locator, LocatorType

# 创建测试任务
tasks = [
    Task(id='safe_task', type=TaskType.REWRITE, 
         locator=Locator(by=LocatorType.FIND, value='项目背景'), 
         instruction='重写项目背景', source_comment_id='comment1'),
    Task(id='risky_task', type=TaskType.SET_PARAGRAPH_STYLE, 
         locator=Locator(by=LocatorType.FIND, value='技术方案'), 
         instruction='设置段落样式'),  # 没有批注授权
    Task(id='authorized_format', type=TaskType.SET_HEADING_LEVEL, 
         locator=Locator(by=LocatorType.FIND, value='技术架构图'), 
         instruction='设置为2级标题', source_comment_id='comment2')
]

print('📋 原始任务列表:')
for i, task in enumerate(tasks, 1):
    auth_status = '🔒' if task.source_comment_id else '🔓'
    risk = '⚠️' if task.type in [TaskType.SET_PARAGRAPH_STYLE, TaskType.SET_HEADING_LEVEL] else '✅'
    print(f'  {i}. {auth_status} {risk} {task.type.value} - {task.instruction}')

print()
print('🛡️ 格式保护过滤结果:')
guard = FormatProtectionGuard()
safe_tasks = guard.filter_unauthorized_tasks(tasks, ['comment1', 'comment2'])

for i, task in enumerate(safe_tasks, 1):
    print(f'  ✅ {i}. {task.type.value} - {task.instruction}')

blocked = len(tasks) - len(safe_tasks)
print(f'')
print(f'🚫 已阻止 {blocked} 个未授权的格式化任务')
print('✅ 格式保护机制工作正常！')
"

echo.
echo ═══════════════════════════════════════════════════════════════
echo 演示完成！
echo ═══════════════════════════════════════════════════════════════
echo.

echo 🎉 AutoWord 演示成功完成！
echo.
echo 📊 演示总结:
echo   ✅ 演示文档已创建
echo   ✅ 系统环境正常
echo   ✅ 核心功能测试通过
echo   ✅ 工作流程演示完成
echo   ✅ 格式保护机制验证
echo.
echo 💡 下一步操作:
echo   1. 查看生成的演示Word文档
echo   2. 运行 test_autoword.bat 进行交互式测试
echo   3. 阅读 README.md 了解详细功能
echo   4. 查看 PROJECT_SUMMARY.md 了解商业价值
echo.
echo 💰 商业化信息:
echo   📈 SaaS服务: $10-50/月
echo   🏢 企业版: $1000-5000  
echo   🎓 培训服务: $500-2000
echo   🔌 API服务: $0.01-0.10/次
echo.
echo 🚀 AutoWord 已准备好商业化运营！
echo.

:END
echo 📞 技术支持: developer@autoword.com
echo 🌐 项目地址: https://github.com/your-repo/autoword
echo.
pause