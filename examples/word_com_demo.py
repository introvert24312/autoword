"""
AutoWord Word COM Demo
Word COM 功能演示
"""

import os
import sys
from pathlib import Path

# 添加项目根目录到路径
sys.path.insert(0, str(Path(__file__).parent.parent))

from autoword.core.doc_loader import DocLoader
from autoword.core.doc_inspector import DocInspector
from autoword.core.exceptions import DocumentError, COMError


def demo_word_availability():
    """演示 Word 可用性检查"""
    print("=== Word COM 可用性检查 ===\n")
    
    loader = DocLoader()
    
    # 检查 Word 是否可用
    if loader.check_word_availability():
        version = loader.get_word_version()
        print(f"✅ Word COM 可用")
        print(f"📋 Word 版本: {version}")
    else:
        print("❌ Word COM 不可用")
        print("请确保已安装 Microsoft Word")
        return False
    
    return True


def demo_document_operations():
    """演示文档操作（需要真实的 Word 文档）"""
    print("\n=== 文档操作演示 ===\n")
    
    # 查找示例文档
    sample_docs = [
        "sample.docx",
        "test.docx", 
        "demo.docx",
        "example.docx"
    ]
    
    doc_path = None
    for doc in sample_docs:
        if Path(doc).exists():
            doc_path = doc
            break
    
    if not doc_path:
        print("⚠️ 未找到示例文档")
        print("请在当前目录放置一个 .docx 文件进行测试")
        print(f"支持的文件名: {', '.join(sample_docs)}")
        return
    
    print(f"📄 找到文档: {doc_path}")
    
    try:
        loader = DocLoader(visible=False)  # 不显示 Word 窗口
        inspector = DocInspector()
        
        print("🔄 正在加载文档...")
        
        # 使用上下文管理器打开文档
        with loader.open_document(doc_path, create_backup=True) as (word_app, document):
            print("✅ 文档加载成功")
            
            # 提取文档结构
            print("🔍 正在分析文档结构...")
            structure = inspector.extract_structure(document)
            
            print(f"📊 文档统计:")
            print(f"   - 页数: {structure.page_count}")
            print(f"   - 字数: {structure.word_count}")
            print(f"   - 标题数: {len(structure.headings)}")
            print(f"   - 样式数: {len(structure.styles)}")
            print(f"   - 超链接数: {len(structure.hyperlinks)}")
            print(f"   - 引用数: {len(structure.references)}")
            
            # 显示标题信息
            if structure.headings:
                print(f"\n📋 标题列表:")
                for i, heading in enumerate(structure.headings[:5], 1):  # 只显示前5个
                    print(f"   {i}. [{heading.level}级] {heading.text[:50]}...")
                if len(structure.headings) > 5:
                    print(f"   ... 还有 {len(structure.headings) - 5} 个标题")
            
            # 提取批注
            print("\n💬 正在提取批注...")
            comments = inspector.extract_comments(document)
            
            if comments:
                print(f"📝 找到 {len(comments)} 个批注:")
                for i, comment in enumerate(comments[:3], 1):  # 只显示前3个
                    print(f"   {i}. 作者: {comment.author}")
                    print(f"      页码: {comment.page}")
                    print(f"      内容: {comment.comment_text[:100]}...")
                    print(f"      锚点: {comment.anchor_text[:50]}...")
                    print()
                if len(comments) > 3:
                    print(f"   ... 还有 {len(comments) - 3} 个批注")
            else:
                print("📝 文档中没有批注")
            
            # 创建文档快照
            print("\n📸 正在创建文档快照...")
            snapshot = inspector.create_snapshot(document, doc_path)
            print(f"✅ 快照创建成功")
            print(f"   - 时间戳: {snapshot.timestamp}")
            print(f"   - 校验和: {snapshot.checksum[:16]}...")
        
        print("\n✅ 文档操作演示完成")
        
    except COMError as e:
        print(f"❌ COM 错误: {e}")
    except DocumentError as e:
        print(f"❌ 文档错误: {e}")
    except Exception as e:
        print(f"❌ 未知错误: {e}")


def demo_backup_operations():
    """演示备份操作"""
    print("\n=== 备份操作演示 ===\n")
    
    # 创建一个临时测试文件
    test_file = "temp_test.txt"
    try:
        with open(test_file, 'w', encoding='utf-8') as f:
            f.write("这是一个测试文件\n用于演示备份功能")
        
        print(f"📄 创建测试文件: {test_file}")
        
        loader = DocLoader()
        
        # 注意：这里会因为文件扩展名不是 .docx 而失败，这是预期的
        try:
            backup_path = loader.create_backup(test_file)
            print(f"✅ 备份创建成功: {backup_path}")
        except Exception as e:
            print(f"⚠️ 备份失败（预期的）: {e}")
            print("💡 备份功能只支持 .docx 文件")
        
    finally:
        # 清理测试文件
        if Path(test_file).exists():
            os.remove(test_file)
            print(f"🗑️ 清理测试文件: {test_file}")


def main():
    """主函数"""
    print("🚀 AutoWord Word COM 功能演示\n")
    
    # 检查 Word 可用性
    if not demo_word_availability():
        return
    
    # 演示备份操作
    demo_backup_operations()
    
    # 演示文档操作
    demo_document_operations()
    
    print("\n🎉 演示完成！")
    print("\n💡 提示:")
    print("   - 将 .docx 文件放在当前目录可以测试完整功能")
    print("   - 建议使用包含批注和标题的文档进行测试")
    print("   - 所有操作都会自动创建备份文件")


if __name__ == "__main__":
    main()