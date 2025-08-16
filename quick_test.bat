@echo off
chcp 65001 >nul
title AutoWord 快速测试
color 0B

echo.
echo ⚡ AutoWord 快速功能测试
echo.

echo 🔍 检查Python环境...
python --version
if errorlevel 1 (
    echo ❌ Python 未安装
    pause
    exit
)

echo.
echo 📦 检查核心模块...
python -c "
try:
    from autoword.core.models import Task, Comment, TaskType
    from autoword.core.llm_client import LLMClient
    from autoword.core.enhanced_executor import EnhancedExecutor
    print('✅ 核心模块导入成功')
except ImportError as e:
    print(f'❌ 模块导入失败: {e}')
    exit(1)
"

echo.
echo 🧪 运行核心功能测试...
python -c "
import sys
import os
sys.path.append('.')

print('🚀 AutoWord 核心功能快速测试')
print('=' * 50)

# 测试数据模型
print('📊 测试数据模型...')
from autoword.core.models import Comment, Task, TaskType, Locator, LocatorType

comment = Comment(
    id='test_comment',
    author='测试用户',
    text='这是一个测试批注',
    anchor_text='测试文本',
    page_number=1,
    position_start=0,
    position_end=10
)
print(f'✅ 批注创建成功: {comment.author} - {comment.text}')

task = Task(
    id='test_task',
    type=TaskType.REWRITE,
    locator=Locator(by=LocatorType.FIND, value='测试'),
    instruction='重写这段内容',
    source_comment_id='test_comment'
)
print(f'✅ 任务创建成功: {task.type.value} - {task.instruction}')

# 测试LLM客户端
print()
print('🤖 测试LLM客户端...')
from autoword.core.llm_client import LLMClient, ModelType

client = LLMClient()
print('✅ LLM客户端创建成功')

# 检查API密钥
api_available = bool(os.getenv('GPT5_KEY') or os.getenv('CLAUDE37_KEY'))
if api_available:
    print('✅ API密钥已配置')
else:
    print('⚠️ API密钥未配置 (演示模式)')

# 测试提示构建器
print()
print('📝 测试提示构建器...')
from autoword.core.prompt_builder import PromptBuilder, PromptContext
from autoword.core.models import Heading, Style

context = PromptContext(
    headings=[Heading(text='测试标题', level=1, page_number=1)],
    styles=[Style(name='正文', type='paragraph')],
    toc_entries=[],
    hyperlinks=[],
    comments=[comment]
)

builder = PromptBuilder()
prompt = builder.build_user_prompt(context)
print(f'✅ 提示词构建成功 (长度: {len(prompt)} 字符)')

# 测试任务规划器
print()
print('🎯 测试任务规划器...')
from autoword.core.planner import FormatProtectionGuard

guard = FormatProtectionGuard()
tasks = [task]
filtered_tasks = guard.filter_unauthorized_tasks(tasks, ['test_comment'])
print(f'✅ 格式保护过滤成功 (任务数: {len(filtered_tasks)})')

# 测试增强执行器
print()
print('🚀 测试增强执行器...')
from autoword.core.enhanced_executor import EnhancedExecutor, WorkflowMode

executor = EnhancedExecutor()
print('✅ 增强执行器创建成功')

print()
print('🎉 所有核心功能测试通过！')
print()
print('💡 系统状态:')
print(f'  📊 数据模型: ✅ 正常')
print(f'  🤖 LLM客户端: ✅ 正常')
print(f'  📝 提示构建器: ✅ 正常')
print(f'  🛡️ 格式保护: ✅ 正常')
print(f'  🚀 执行器: ✅ 正常')
print()
print('🚀 AutoWord 系统就绪，可以开始处理文档！')
"

echo.
echo 📋 测试完成！双击 test_autoword.bat 进行完整测试
echo.
pause