# AutoWord

AutoWord 是一个智能的 Word 文档自动化工具，通过 LLM 技术将文档批注转换为可执行的任务，实现文档的自动化编辑和格式化。

## 🌟 特性

- 🤖 **智能批注解析**: 使用 GPT-5/Claude 3.7 理解文档批注意图
- 📝 **自动任务生成**: 将批注转换为结构化的执行任务
- 🔒 **四重格式保护**: 确保只有授权的格式变更被执行
- ⚡ **Word COM 集成**: 直接操作 Word 文档，保持格式完整性
- 🎯 **精确定位**: 支持书签、范围、标题、文本查找等定位方式
- 📊 **详细报告**: 生成完整的执行日志和变更报告
- 🔄 **智能回滚**: 检测未授权变更并自动回滚
- 🧪 **试运行模式**: 预览变更效果而不实际修改文档

## 🛠️ 系统要求

- **操作系统**: Windows 10 或更高版本
- **Python**: 3.10+
- **Microsoft Word**: 2016 或更高版本
- **内存**: 4GB RAM（推荐 8GB）
- **网络**: 用于 LLM API 调用

## 📦 安装

### 从源码安装

```bash
git clone https://github.com/autoword/autoword.git
cd autoword
pip install -e .
```

### 安装依赖

```bash
# 基础依赖
pip install -r requirements.txt

# 开发依赖
pip install -e .[dev]
```

## 🚀 快速开始

### 1. 配置 API 密钥

设置环境变量：

```bash
# Windows CMD
set GPT5_KEY=your_gpt5_api_key
set CLAUDE37_KEY=your_claude37_api_key

# Windows PowerShell
$env:GPT5_KEY="your_gpt5_api_key"
$env:CLAUDE37_KEY="your_claude37_api_key"
```

### 2. 检查环境

```bash
python autoword/cli/main.py check
```

### 3. 处理文档

#### 命令行方式

```bash
# 基本处理
python autoword/cli/main.py process document.docx

# 试运行（不实际修改文档）
python autoword/cli/main.py process document.docx --dry-run

# 使用 Claude 模型
python autoword/cli/main.py process document.docx --model claude37

# 详细输出
python autoword/cli/main.py process document.docx --verbose

# 检查文档结构
python autoword/cli/main.py inspect document.docx
```

#### Python API 方式

```python
from autoword.core.pipeline import process_document_simple
from autoword.core.llm_client import ModelType

# 简单处理
result = process_document_simple("document.docx")

if result.success:
    print(f"✅ 处理成功!")
    print(f"   完成任务: {result.execution_result.completed_tasks}/{result.execution_result.total_tasks}")
    print(f"   总耗时: {result.total_time:.2f}s")
else:
    print(f"❌ 处理失败: {result.error_message}")

# 试运行
result = process_document_simple("document.docx", dry_run=True)

# 使用不同模型
result = process_document_simple("document.docx", model=ModelType.CLAUDE37)
```

## 📖 核心概念

### 批注驱动的自动化

AutoWord 通过分析 Word 文档中的批注来理解用户意图：

```
批注示例：
- "重写这段文字，使其更加简洁明了"
- "将此标题改为2级标题"
- "在这里插入一个关于技术架构的段落"
- "删除这个过时的信息"
- "更新目录页码"
```

### 四重格式保护防线

确保文档格式安全的多层保护机制：

1. **🎯 提示词硬约束**: LLM 系统提示明确禁止未授权格式变更
2. **🔍 规划期过滤**: 过滤无批注来源的格式类任务
3. **🚫 执行期拦截**: 执行前再次校验批注授权
4. **🔄 事后校验回滚**: 检测未授权变更并自动回滚

### 支持的任务类型

#### 内容任务（无需批注授权）
- `rewrite`: 重写文本内容
- `insert`: 插入新内容
- `delete`: 删除指定内容

#### 格式任务（需要批注授权）
- `set_paragraph_style`: 设置段落样式
- `set_heading_level`: 设置标题级别
- `apply_template`: 应用文档模板

#### 结构任务
- `rebuild_toc`: 重建目录
- `update_toc_levels`: 更新目录级别
- `refresh_toc_numbers`: 刷新目录页码
- `replace_hyperlink`: 替换超链接

## 🔧 高级用法

### 自定义配置

```python
from autoword.core.pipeline import DocumentProcessor, PipelineConfig
from autoword.core.llm_client import ModelType
from autoword.core.word_executor import ExecutionMode

# 创建自定义配置
config = PipelineConfig(
    model=ModelType.CLAUDE37,           # 使用 Claude 模型
    execution_mode=ExecutionMode.SAFE,  # 安全模式
    create_backup=True,                 # 创建备份
    enable_validation=True,             # 启用验证
    export_results=True,                # 导出结果
    output_dir="custom_output",         # 自定义输出目录
    visible_word=False,                 # 隐藏 Word 窗口
    max_retries=3                       # 最大重试次数
)

# 使用自定义配置
processor = DocumentProcessor(config)

# 添加进度回调
def progress_callback(progress):
    print(f"[{progress.stage.value}] {progress.progress:.1%} - {progress.message}")

processor.add_progress_callback(progress_callback)

# 处理文档
result = processor.process_document("document.docx")
```

### 批量处理

```python
import os
from pathlib import Path
from autoword.core.pipeline import process_document_simple

def batch_process(directory, pattern="*.docx"):
    """批量处理文档"""
    documents = Path(directory).glob(pattern)
    results = []
    
    for doc_path in documents:
        print(f"处理: {doc_path.name}")
        
        try:
            result = process_document_simple(str(doc_path))
            results.append((doc_path.name, result.success, result.error_message))
            
            if result.success:
                print(f"  ✅ 成功 - {result.execution_result.completed_tasks} 个任务")
            else:
                print(f"  ❌ 失败 - {result.error_message}")
                
        except Exception as e:
            print(f"  💥 异常 - {e}")
            results.append((doc_path.name, False, str(e)))
    
    return results

# 使用示例
results = batch_process("documents/")
```

## 🧪 开发

### 环境设置

```bash
# 克隆仓库
git clone https://github.com/autoword/autoword.git
cd autoword

# 创建虚拟环境
python -m venv venv
venv\Scripts\activate

# 安装开发依赖
pip install -e .[dev]
```

### 运行测试

```bash
# 运行所有测试
python -m pytest tests/ -v

# 运行特定测试
python -m pytest tests/test_pipeline.py -v

# 生成覆盖率报告
python -m pytest tests/ --cov=autoword --cov-report=html
```

### 构建和打包

```bash
# 验证环境
python scripts/build.py --validate

# 运行测试
python scripts/build.py --test

# 构建 Python 包
python scripts/build.py --package

# 构建可执行文件
python scripts/build.py --exe

# 完整构建流程
python scripts/build.py --all
```

## 📋 项目结构

```
autoword/
├── autoword/
│   ├── core/                 # 核心功能模块
│   │   ├── models.py         # 数据模型
│   │   ├── doc_loader.py     # 文档加载器
│   │   ├── doc_inspector.py  # 文档检查器
│   │   ├── llm_client.py     # LLM 客户端
│   │   ├── prompt_builder.py # 提示词构建器
│   │   ├── planner.py        # 任务规划器
│   │   ├── word_executor.py  # Word 执行器
│   │   ├── pipeline.py       # 管道编排器
│   │   ├── exporter.py       # 结果导出器
│   │   └── error_handler.py  # 错误处理器
│   ├── cli/                  # 命令行界面
│   └── gui/                  # 图形界面（规划中）
├── tests/                    # 测试文件
├── examples/                 # 示例代码
├── schemas/                  # JSON Schema
├── scripts/                  # 构建脚本
└── docs/                     # 文档
```

## 🤝 贡献

我们欢迎各种形式的贡献！

### 报告问题

- 使用 GitHub Issues 报告 bug
- 提供详细的错误信息和重现步骤
- 包含系统环境信息

### 提交代码

1. Fork 项目
2. 创建特性分支 (`git checkout -b feature/amazing-feature`)
3. 提交更改 (`git commit -m 'Add amazing feature'`)
4. 推送到分支 (`git push origin feature/amazing-feature`)
5. 创建 Pull Request

### 开发指南

- 遵循 PEP 8 代码风格
- 添加适当的测试
- 更新相关文档
- 确保所有测试通过

## 📄 许可证

本项目采用 MIT 许可证。

## 🙏 致谢

- Microsoft Word COM API
- OpenAI GPT 和 Anthropic Claude
- Python 开源社区

---

**AutoWord** - 让 Word 文档编辑更智能、更高效！ 🚀