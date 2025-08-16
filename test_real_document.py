#!/usr/bin/env python3
"""
测试真实文档处理
"""

import sys
sys.path.append('.')

from autoword.core.enhanced_executor import EnhancedExecutor, WorkflowMode
from autoword.core.llm_client import ModelType

def test_real_document():
    print("测试真实文档处理...")
    
    document_path = r"C:\Users\y\Desktop\郭宇睿-现代汉语言文学在社会发展中的作用与影响研究2稿批注.docx"
    
    # 创建执行器
    executor = EnhancedExecutor(
        llm_model=ModelType.GPT5,
        visible=False
    )
    
    print("开始处理文档...")
    
    # 先测试文档加载
    try:
        document_info = executor._load_and_inspect_document(document_path)
        print(f"✅ 文档加载成功")
        print(f"📄 标题: {document_info.title}")
        print(f"📊 页数: {document_info.page_count}")
        print(f"📝 字数: {document_info.word_count}")
        print(f"💬 批注数: {len(document_info.comments)}")
        
        print("\n批注详情:")
        for i, comment in enumerate(document_info.comments, 1):
            print(f"  {i}. ID: {comment.id}")
            print(f"     作者: {comment.author}")
            print(f"     内容: {comment.comment_text[:100]}...")
            print(f"     位置: {comment.anchor_text[:50]}...")
            print()
        
        if document_info.comments:
            print("测试任务规划...")
            try:
                context = executor.prompt_builder.build_context_from_document(document_info)
                user_prompt = executor.prompt_builder.build_user_prompt(context)
                system_prompt = executor.prompt_builder.build_system_prompt()
                
                print(f"系统提示词长度: {len(system_prompt)}")
                print(f"用户提示词长度: {len(user_prompt)}")
                
                print("\n系统提示词:")
                print(system_prompt)
                
                print("\n用户提示词:")
                print(user_prompt[:1000] + "..." if len(user_prompt) > 1000 else user_prompt)
                
                # 直接调用LLM
                response = executor.llm_client.call_gpt5(system_prompt, user_prompt)
                
                print(f"\nLLM响应:")
                print(f"成功: {response.success}")
                print(f"内容: {response.content}")
                print(f"错误: {response.error}")
                
            except Exception as e:
                print(f"❌ 任务规划失败: {e}")
                import traceback
                traceback.print_exc()
        else:
            print("⚠️ 文档中没有批注")
            
    except Exception as e:
        print(f"❌ 文档加载失败: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_real_document()