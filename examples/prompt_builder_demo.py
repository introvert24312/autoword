"""
AutoWord Prompt Builder Demo
提示词构建器使用示例
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import json
from autoword.core.prompt_builder import PromptBuilder, PromptContext
from autoword.core.models import (
    Comment, DocumentStructure, Heading, Style, TocEntry, Hyperlink, Reference
)


def demo_basic_usage():
    """演示基本用法"""
    print("=== AutoWord Prompt Builder Demo ===\n")
    
    # 创建提示词构建器
    builder = PromptBuilder()
    
    # 创建示例文档结构
    headings = [
        Heading(level=1, text="第一章 项目概述", style="标题 1", range_start=0, range_end=20),
        Heading(level=2, text="1.1 项目背景", style="标题 2", range_start=100, range_end=120),
        Heading(level=2, text="1.2 项目目标", style="标题 2", range_start=200, range_end=220),
        Heading(level=1, text="第二章 技术方案", style="标题 1", range_start=500, range_end=520),
        Heading(level=2, text="2.1 架构设计", style="标题 2", range_start=600, range_end=620),
    ]
    
    styles = [
        Style(name="标题 1", type="paragraph", built_in=True, in_use=True),
        Style(name="标题 2", type="paragraph", built_in=True, in_use=True),
        Style(name="正文", type="paragraph", built_in=True, in_use=True),
        Style(name="代码", type="character", built_in=False, in_use=True),
    ]
    
    toc_entries = [
        TocEntry(level=1, text="第一章 项目概述", page_number=1, range_start=0, range_end=20),
        TocEntry(level=2, text="1.1 项目背景", page_number=2, range_start=100, range_end=120),
        TocEntry(level=2, text="1.2 项目目标", page_number=3, range_start=200, range_end=220),
        TocEntry(level=1, text="第二章 技术方案", page_number=5, range_start=500, range_end=520),
    ]
    
    hyperlinks = [
        Hyperlink(text="GitHub", address="https://github.com", type="web", range_start=300, range_end=310),
        Hyperlink(text="内部链接", address="bookmark1", type="internal", range_start=400, range_end=410),
        Hyperlink(text="联系邮箱", address="mailto:test@example.com", type="email", range_start=450, range_end=460),
    ]
    
    references = [
        Reference(type="bookmark", text="重要章节", target="chapter1", range_start=50, range_end=60),
        Reference(type="field_reference", text="图 1", target="REF figure1", range_start=250, range_end=260),
    ]
    
    document_structure = DocumentStructure(
        headings=headings,
        styles=styles,
        toc_entries=toc_entries,
        hyperlinks=hyperlinks,
        references=references,
        page_count=10,
        word_count=2500
    )
    
    # 创建示例批注
    comments = [
        Comment(
            id="comment_1",
            author="张三",
            page=1,
            anchor_text="项目背景介绍部分需要更详细的说明",
            comment_text="这里需要补充更多关于项目起源和市场需求的内容，建议增加2-3段详细描述",
            range_start=110,
            range_end=150
        ),
        Comment(
            id="comment_2",
            author="李四",
            page=2,
            anchor_text="技术选型说明",
            comment_text="建议在这里插入一个技术对比表格，展示不同方案的优缺点",
            range_start=610,
            range_end=650
        ),
        Comment(
            id="comment_3",
            author="王五",
            page=3,
            anchor_text="架构图",
            comment_text="这个架构图需要重新绘制，当前版本不够清晰，建议使用更专业的绘图工具",
            range_start=700,
            range_end=720
        ),
        Comment(
            id="comment_4",
            author="张三",
            page=5,
            anchor_text="代码示例",
            comment_text="删除这个过时的代码示例，替换为最新版本的实现",
            range_start=800,
            range_end=850
        ),
    ]
    
    # 创建提示词上下文
    context = PromptContext(
        document_structure=document_structure,
        comments=comments,
        document_path="项目报告.docx"
    )
    
    print("✅ 创建了示例文档结构和批注")
    print(f"📊 文档统计: {len(headings)} 个标题, {len(comments)} 个批注")
    
    # 演示系统提示词
    print("\n=== 系统提示词 ===")
    system_prompt = builder.build_system_prompt()
    print(system_prompt)
    
    # 演示文档结构摘要
    print("\n=== 文档结构摘要 ===")
    structure_summary = builder.build_document_summary(document_structure)
    print(structure_summary)
    
    # 演示批注摘要
    print("\n=== 批注摘要 ===")
    comments_summary = builder.build_comments_summary(comments)
    print(comments_summary)
    
    # 演示完整用户提示词
    print("\n=== 完整用户提示词 ===")
    user_prompt = builder.build_user_prompt(context)
    print(f"用户提示词长度: {len(user_prompt)} 字符")
    print("前500字符预览:")
    print(user_prompt[:500] + "...")
    
    # 演示上下文长度检查
    print("\n=== 上下文长度检查 ===")
    length_check = builder.check_context_length(context)
    print(f"系统提示词 tokens: {length_check['system_tokens']}")
    print(f"用户提示词 tokens: {length_check['user_tokens']}")
    print(f"总计 tokens: {length_check['total_tokens']}")
    print(f"是否在限制内: {length_check['is_within_limit']}")
    print(f"最大限制: {length_check['max_tokens']}")
    
    if length_check['overflow_tokens'] > 0:
        print(f"超出 tokens: {length_check['overflow_tokens']}")


def demo_context_overflow():
    """演示上下文溢出处理"""
    print("\n=== 上下文溢出处理演示 ===\n")
    
    builder = PromptBuilder()
    
    # 创建大量内容来模拟溢出
    large_headings = []
    large_comments = []
    
    # 创建多个章节
    for chapter in range(1, 6):  # 5个章节
        # 章节标题
        chapter_heading = Heading(
            level=1, 
            text=f"第{chapter}章 章节标题{chapter}",
            style="标题 1",
            range_start=chapter * 1000,
            range_end=chapter * 1000 + 50
        )
        large_headings.append(chapter_heading)
        
        # 每章节的子标题
        for section in range(1, 4):  # 每章3个小节
            section_heading = Heading(
                level=2,
                text=f"{chapter}.{section} 小节标题{chapter}.{section}",
                style="标题 2", 
                range_start=chapter * 1000 + section * 100,
                range_end=chapter * 1000 + section * 100 + 30
            )
            large_headings.append(section_heading)
            
            # 每小节的批注
            for comment_num in range(1, 3):  # 每小节2个批注
                comment = Comment(
                    id=f"comment_{chapter}_{section}_{comment_num}",
                    author=f"审阅者{comment_num}",
                    page=chapter * 2 + section,
                    anchor_text=f"第{chapter}章第{section}节的内容需要修改" * 3,  # 长锚点文本
                    comment_text=f"这是第{chapter}章第{section}节的详细批注内容，需要进行大幅度的修改和完善。" * 5,  # 长批注内容
                    range_start=chapter * 1000 + section * 100 + comment_num * 10,
                    range_end=chapter * 1000 + section * 100 + comment_num * 10 + 20
                )
                large_comments.append(comment)
    
    large_structure = DocumentStructure(
        headings=large_headings,
        styles=[],
        toc_entries=[],
        hyperlinks=[],
        references=[],
        page_count=20,
        word_count=10000
    )
    
    # 设置较小的上下文限制来触发溢出
    context = PromptContext(
        document_structure=large_structure,
        comments=large_comments,
        max_context_length=5000  # 较小的限制
    )
    
    print(f"📊 创建了大型文档: {len(large_headings)} 个标题, {len(large_comments)} 个批注")
    
    # 检查是否溢出
    length_check = builder.check_context_length(context)
    print(f"总计 tokens: {length_check['total_tokens']}")
    print(f"最大限制: {length_check['max_tokens']}")
    print(f"是否溢出: {not length_check['is_within_limit']}")
    
    if not length_check['is_within_limit']:
        print(f"超出 tokens: {length_check['overflow_tokens']}")
        
        # 处理溢出
        print("\n🔄 处理上下文溢出...")
        chunks = builder.handle_context_overflow(context)
        
        print(f"✅ 分割成 {len(chunks)} 个块:")
        for i, chunk in enumerate(chunks, 1):
            chunk_check = builder.check_context_length(chunk)
            print(f"  块 {i}: {len(chunk.document_structure.headings)} 个标题, "
                  f"{len(chunk.comments)} 个批注, "
                  f"{chunk_check['total_tokens']} tokens")


def demo_json_schema():
    """演示 JSON Schema 功能"""
    print("\n=== JSON Schema 演示 ===\n")
    
    builder = PromptBuilder()
    
    # 获取 Schema
    schema = builder.get_schema()
    
    print("📋 JSON Schema 结构:")
    print(json.dumps(schema, ensure_ascii=False, indent=2)[:1000] + "...")
    
    # 验证 Schema 的关键部分
    print(f"\n✅ Schema 验证:")
    print(f"- 根类型: {schema.get('type')}")
    print(f"- 必需字段: {schema.get('required', [])}")
    
    if 'properties' in schema and 'tasks' in schema['properties']:
        tasks_schema = schema['properties']['tasks']
        print(f"- 任务数组类型: {tasks_schema.get('type')}")
        
        if 'items' in tasks_schema:
            task_schema = tasks_schema['items']
            print(f"- 任务对象必需字段: {task_schema.get('required', [])}")
            
            if 'properties' in task_schema and 'type' in task_schema['properties']:
                type_enum = task_schema['properties']['type'].get('enum', [])
                print(f"- 支持的任务类型: {len(type_enum)} 种")
                print(f"  {', '.join(type_enum[:5])}{'...' if len(type_enum) > 5 else ''}")


if __name__ == "__main__":
    demo_basic_usage()
    demo_context_overflow()
    demo_json_schema()
    
    print("\n=== 演示完成 ===")
    print("💡 提示: 提示词构建器已准备好与 LLM 客户端集成使用")