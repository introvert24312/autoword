#!/usr/bin/env python3
"""
测试批注提取功能
"""

import sys
import os
import json
from pathlib import Path

# 添加项目路径
sys.path.insert(0, str(Path(__file__).parent))

def test_comment_extraction():
    """测试批注提取"""
    print("🔍 测试批注提取功能")
    print("=" * 40)
    
    try:
        import win32com.client
        
        # 查找带批注的文档
        test_files = list(Path(".").glob("*批注*.docx"))
        if not test_files:
            test_files = list(Path("C:/Users/y/Desktop").glob("*批注*.docx"))
        
        if not test_files:
            print("❌ 未找到带批注的测试文档")
            return False
        
        docx_path = str(test_files[0])
        print(f"📄 测试文档: {docx_path}")
        
        # 打开Word应用
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        
        try:
            # 打开文档
            doc = word.Documents.Open(docx_path)
            
            # 提取批注
            comments = []
            print(f"📝 文档中共有 {doc.Comments.Count} 条批注")
            
            for i, comment in enumerate(doc.Comments):
                if not comment.Done:  # 只处理未解决的批注
                    comment_data = {
                        "comment_id": f"comment_{i+1}",
                        "author": comment.Author,
                        "created_time": str(comment.Date),
                        "text": comment.Range.Text.strip(),
                        "resolved": comment.Done,
                        "anchor": {
                            "paragraph_start": comment.Scope.Start,
                            "paragraph_end": comment.Scope.End,
                            "char_start": comment.Scope.Start,
                            "char_end": comment.Scope.End
                        }
                    }
                    comments.append(comment_data)
                    
                    print(f"批注 {i+1}: {comment.Author} - {comment.Range.Text[:50]}...")
            
            # 关闭文档
            doc.Close(False)
            
            # 创建批注JSON
            comments_json = {
                "schema_version": "comments.v1",
                "document_path": docx_path,
                "extraction_time": "2025-01-18 13:30:00",
                "total_comments": len(comments),
                "comments": comments
            }
            
            # 保存JSON文件
            docx_file = Path(docx_path)
            json_filename = docx_file.parent / f"{docx_file.stem}_comments.json"
            
            with open(json_filename, 'w', encoding='utf-8') as f:
                json.dump(comments_json, f, indent=2, ensure_ascii=False)
            
            print(f"✅ 批注JSON已保存: {json_filename}")
            print(f"📊 提取了 {len(comments)} 条未解决的批注")
            
            # 显示JSON内容预览
            print("\n📋 JSON内容预览:")
            print("-" * 30)
            print(json.dumps(comments_json, indent=2, ensure_ascii=False)[:500] + "...")
            
            return True
            
        finally:
            word.Quit()
            
    except Exception as e:
        print(f"❌ 批注提取失败: {e}")
        return False

def main():
    """主函数"""
    success = test_comment_extraction()
    
    if success:
        print("\n🎉 批注提取测试成功！")
    else:
        print("\n❌ 批注提取测试失败！")
    
    input("\n按回车键退出...")

if __name__ == "__main__":
    main()