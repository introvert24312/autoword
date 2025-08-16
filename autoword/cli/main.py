#!/usr/bin/env python3
"""
AutoWord CLI 主入口
"""

import os
import sys
import click
import logging
from pathlib import Path

# 添加项目根目录到路径
sys.path.insert(0, str(Path(__file__).parent.parent.parent))

from autoword.core.pipeline import process_document_simple
from autoword.core.llm_client import ModelType
from autoword.core.error_handler import handle_error
from autoword.core.exceptions import AutoWordError


# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)


@click.group()
@click.version_option(version="1.0.0", prog_name="AutoWord")
def cli():
    """AutoWord - 智能 Word 文档自动化工具"""
    pass


@cli.command()
@click.argument('document_path', type=click.Path(exists=True))
@click.option('--model', type=click.Choice(['gpt5', 'claude37']), default='gpt5', help='LLM 模型')
@click.option('--dry-run', is_flag=True, help='试运行模式')
@click.option('--output', default='output', help='输出目录')
@click.option('--verbose', '-v', is_flag=True, help='详细输出')
def process(document_path, model, dry_run, output, verbose):
    """处理 Word 文档"""
    
    if verbose:
        logging.getLogger().setLevel(logging.DEBUG)
    
    try:
        click.echo(f"开始处理文档: {document_path}")
        
        # 转换模型名称
        model_type = ModelType.GPT5 if model == 'gpt5' else ModelType.CLAUDE37
        
        # 处理文档
        result = process_document_simple(
            document_path=document_path,
            model=model_type,
            dry_run=dry_run,
            output_dir=output
        )
        
        if result.success:
            click.echo(f"✅ 处理成功!")
            click.echo(f"   总耗时: {result.total_time:.2f}s")
            click.echo(f"   完成阶段: {len(result.stages_completed)}")
            
            if result.task_plan:
                click.echo(f"   任务数量: {len(result.task_plan.tasks)}")
            
            if result.execution_result:
                click.echo(f"   执行结果: {result.execution_result.completed_tasks}/{result.execution_result.total_tasks}")
            
            if result.exported_files:
                click.echo(f"   导出文件: {len(result.exported_files)} 个")
                for name, path in result.exported_files.items():
                    click.echo(f"     {name}: {path}")
        else:
            click.echo(f"❌ 处理失败: {result.error_message}")
            sys.exit(1)
            
    except Exception as e:
        error_info = handle_error(e, "CLI处理")
        click.echo(f"❌ {error_info.user_message}")
        
        if verbose:
            click.echo(f"技术详情: {error_info.message}")
            if error_info.suggestions:
                click.echo("建议:")
                for suggestion in error_info.suggestions:
                    click.echo(f"  - {suggestion}")
        
        sys.exit(1)


@cli.command()
def check():
    """检查系统环境"""
    
    click.echo("检查 AutoWord 运行环境...")
    
    # 检查 Python 版本
    python_version = sys.version_info
    if python_version >= (3, 10):
        click.echo(f"✅ Python 版本: {python_version.major}.{python_version.minor}.{python_version.micro}")
    else:
        click.echo(f"❌ Python 版本过低: {python_version.major}.{python_version.minor}.{python_version.micro} (需要 3.10+)")
    
    # 检查 Word COM
    try:
        import win32com.client
        word = win32com.client.Dispatch("Word.Application")
        word_version = word.Version
        word.Quit()
        click.echo(f"✅ Microsoft Word: {word_version}")
    except Exception as e:
        click.echo(f"❌ Microsoft Word 不可用: {e}")
    
    # 检查 API 密钥
    gpt5_key = os.getenv("GPT5_KEY")
    claude_key = os.getenv("CLAUDE37_KEY")
    
    if gpt5_key:
        click.echo("✅ GPT5_KEY 已配置")
    else:
        click.echo("⚠️ GPT5_KEY 未配置")
    
    if claude_key:
        click.echo("✅ CLAUDE37_KEY 已配置")
    else:
        click.echo("⚠️ CLAUDE37_KEY 未配置")
    
    if not gpt5_key and not claude_key:
        click.echo("❌ 未配置任何 LLM API 密钥")
        click.echo("请设置环境变量 GPT5_KEY 或 CLAUDE37_KEY")
    
    click.echo("环境检查完成")


@cli.command()
@click.argument('document_path', type=click.Path(exists=True))
def inspect(document_path):
    """检查文档结构和批注"""
    
    try:
        from autoword.core.doc_loader import WordSession, DocLoader
        from autoword.core.doc_inspector import DocInspector
        
        click.echo(f"检查文档: {document_path}")
        
        with WordSession() as word_app:
            loader = DocLoader()
            inspector = DocInspector()
            
            # 加载文档
            document = loader.load_document(word_app, document_path)
            
            try:
                # 提取信息
                comments = inspector.extract_comments(document)
                structure = inspector.extract_structure(document)
                
                # 显示结果
                click.echo(f"📄 文档信息:")
                click.echo(f"   页数: {structure.page_count}")
                click.echo(f"   字数: {structure.word_count}")
                click.echo(f"   标题: {len(structure.headings)} 个")
                click.echo(f"   样式: {len([s for s in structure.styles if s.in_use])} 个")
                click.echo(f"   批注: {len(comments)} 个")
                click.echo(f"   超链接: {len(structure.hyperlinks)} 个")
                
                if comments:
                    click.echo(f"\n💬 批注详情:")
                    for i, comment in enumerate(comments[:5], 1):
                        click.echo(f"   {i}. {comment.author} (第{comment.page}页): {comment.comment_text[:50]}...")
                    
                    if len(comments) > 5:
                        click.echo(f"   ... 还有 {len(comments) - 5} 个批注")
                
            finally:
                document.Close()
        
    except Exception as e:
        error_info = handle_error(e, "文档检查")
        click.echo(f"❌ {error_info.user_message}")
        sys.exit(1)


def main():
    """主入口函数"""
    cli()


if __name__ == "__main__":
    main()