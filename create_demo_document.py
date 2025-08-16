#!/usr/bin/env python3
"""
创建演示用的Word文档
包含批注和各种格式，用于测试AutoWord功能
"""

import os
import sys
from datetime import datetime

def create_demo_document():
    """创建演示文档"""
    try:
        import win32com.client
    except ImportError:
        print("❌ 需要安装 pywin32: pip install pywin32")
        return False
    
    try:
        # 启动Word应用
        print("🚀 启动Microsoft Word...")
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = True
        
        # 创建新文档
        doc = word.Documents.Add()
        
        # 设置文档标题
        print("📝 创建文档内容...")
        doc.Range().Text = ""
        
        # 添加标题
        title_range = doc.Range()
        title_range.Text = "AutoWord 演示文档\n\n"
        title_range.Style = "标题 1"
        title_range.ParagraphFormat.Alignment = 1  # 居中
        
        # 添加项目背景部分
        doc.Range().InsertAfter("项目背景\n")
        heading1 = doc.Paragraphs(doc.Paragraphs.Count).Range
        heading1.Style = "标题 2"
        
        background_text = """本项目旨在开发一个基于人工智能的文档自动化处理系统。当前市场上缺乏智能化的文档编辑工具，大多数解决方案仍然依赖人工操作，效率低下且容易出错。

我们的解决方案将利用最新的大语言模型技术，实现文档的智能化处理和编辑。"""
        
        doc.Range().InsertAfter(background_text + "\n\n")
        
        # 添加技术方案部分
        doc.Range().InsertAfter("技术方案\n")
        heading2 = doc.Paragraphs(doc.Paragraphs.Count).Range
        heading2.Style = "标题 2"
        
        # 添加技术架构图子标题
        doc.Range().InsertAfter("技术架构图\n")
        subheading = doc.Paragraphs(doc.Paragraphs.Count).Range
        subheading.Style = "标题 3"
        
        tech_text = """系统采用模块化设计，主要包括以下组件：
- LLM客户端模块
- 文档解析模块  
- 任务规划模块
- 执行引擎模块
- 安全保护模块

旧技术栈说明：
早期版本使用的是传统的规则引擎和模板匹配技术，但这种方法灵活性不足，无法处理复杂的文档结构和语义理解需求。因此我们决定采用基于大语言模型的新架构。"""
        
        doc.Range().InsertAfter(tech_text + "\n\n")
        
        # 添加项目结论部分
        doc.Range().InsertAfter("项目结论\n")
        heading3 = doc.Paragraphs(doc.Paragraphs.Count).Range
        heading3.Style = "标题 2"
        
        conclusion_text = """通过本项目的实施，我们将能够提供一个完整的文档自动化解决方案，大幅提升文档处理效率。

预期效果：
- 处理效率提升80%以上
- 错误率降低90%以上  
- 用户满意度达到95%以上"""
        
        doc.Range().InsertAfter(conclusion_text + "\n\n")
        
        # 添加批注
        print("💬 添加批注...")
        
        # 在项目背景部分添加批注
        background_para = None
        for para in doc.Paragraphs:
            if "本项目旨在开发" in para.Range.Text:
                background_para = para
                break
        
        if background_para:
            comment_range = background_para.Range
            comment_range.Start = comment_range.Start + comment_range.Text.find("项目背景")
            comment_range.End = comment_range.Start + len("项目背景")
            
            comment1 = doc.Comments.Add(comment_range)
            comment1.Range.Text = "重写项目背景部分，增加市场分析和竞争对手分析，使内容更加详细和专业"
            comment1.Author = "张三"
        
        # 在技术架构图标题添加批注
        for para in doc.Paragraphs:
            if para.Range.Text.strip() == "技术架构图":
                comment_range = para.Range
                comment2 = doc.Comments.Add(comment_range)
                comment2.Range.Text = "将此标题设置为2级标题"
                comment2.Author = "李四"
                break
        
        # 在旧技术栈部分添加批注
        for para in doc.Paragraphs:
            if "旧技术栈说明" in para.Range.Text:
                comment_range = para.Range
                start_pos = comment_range.Text.find("旧技术栈说明")
                comment_range.Start = comment_range.Start + start_pos
                comment_range.End = comment_range.Start + len("旧技术栈")
                
                comment3 = doc.Comments.Add(comment_range)
                comment3.Range.Text = "删除过时的技术栈说明段落"
                comment3.Author = "王五"
                break
        
        # 在项目结论部分添加批注
        for para in doc.Paragraphs:
            if para.Range.Text.strip() == "项目结论":
                comment_range = para.Range
                comment4 = doc.Comments.Add(comment_range)
                comment4.Range.Text = "在结论部分插入项目时间线表格"
                comment4.Author = "赵六"
                break
        
        # 创建目录
        print("📑 创建目录...")
        doc.Range(0, 0).InsertBreak(7)  # 插入分页符
        toc_range = doc.Range(0, 0)
        toc_range.InsertAfter("目录\n\n")
        
        # 插入目录
        toc = doc.TablesOfContents.Add(
            Range=doc.Range(toc_range.End, toc_range.End),
            RightAlignPageNumbers=True,
            UseHeadingStyles=True,
            IncludePageNumbers=True
        )
        
        # 添加超链接
        print("🔗 添加超链接...")
        link_text = "AutoWord官网"
        link_range = doc.Range()
        link_range.Start = doc.Range().End
        link_range.Text = f"\n\n参考链接：{link_text}"
        
        # 创建超链接
        link_start = link_range.Start + link_range.Text.find(link_text)
        link_end = link_start + len(link_text)
        hyperlink_range = doc.Range(link_start, link_end)
        
        doc.Hyperlinks.Add(
            Anchor=hyperlink_range,
            Address="https://github.com/your-repo/autoword",
            TextToDisplay=link_text
        )
        
        # 保存文档
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"AutoWord_演示文档_{timestamp}.docx"
        filepath = os.path.join(os.getcwd(), filename)
        
        print(f"💾 保存文档: {filename}")
        doc.SaveAs2(filepath)
        
        print("✅ 演示文档创建完成！")
        print()
        print("📋 文档包含以下测试内容:")
        print("  📝 多级标题结构")
        print("  💬 4个测试批注")
        print("  📑 自动生成的目录")
        print("  🔗 外部超链接")
        print("  📊 不同的段落样式")
        print()
        print("🎯 可以测试的AutoWord功能:")
        print("  ✅ 批注驱动的内容重写")
        print("  ✅ 标题级别调整")
        print("  ✅ 内容删除操作")
        print("  ✅ 表格插入功能")
        print("  ✅ 格式保护机制")
        print("  ✅ 目录更新功能")
        print("  ✅ 超链接管理")
        print()
        print(f"📁 文档路径: {filepath}")
        
        return True
        
    except Exception as e:
        print(f"❌ 创建文档失败: {str(e)}")
        return False
    
    finally:
        try:
            # 不关闭Word，让用户可以查看文档
            pass
        except:
            pass

if __name__ == "__main__":
    print("=== AutoWord 演示文档生成器 ===")
    print()
    
    if create_demo_document():
        print("🎉 演示文档生成成功！")
        print()
        print("📋 下一步操作:")
        print("1. 查看生成的Word文档")
        print("2. 运行 test_autoword.bat 测试系统")
        print("3. 使用AutoWord处理这个文档")
        print()
        print("💡 提示: 文档中的批注将指导AutoWord进行自动化处理")
    else:
        print("❌ 演示文档生成失败")
    
    input("\n按回车键退出...")