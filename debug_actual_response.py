#!/usr/bin/env python3
"""
调试实际的LLM响应
"""

import sys
import json
from pathlib import Path

# 添加项目根目录到Python路径
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

def debug_actual_response():
    """调试实际的LLM响应"""
    print("========================================")
    print("      调试实际LLM响应")
    print("========================================")
    print()
    
    # 测试文档路径
    test_doc = r"C:\Users\y\Desktop\郭宇睿-现代汉语言文学在社会发展中的作用与影响研究2稿批注.docx"
    
    try:
        # 1. 加载文档和提取批注
        print("[1/3] 加载文档和提取批注...")
        from autoword.core.doc_loader import DocLoader
        from autoword.core.doc_inspector import DocInspector
        
        doc_loader = DocLoader()
        doc_inspector = DocInspector()
        
        word_app, document = doc_loader.load_document(test_doc, create_backup=False)
        
        try:
            comments = doc_inspector.extract_comments(document)
            structure = doc_inspector.extract_structure(document)
            print(f"✅ 提取到 {len(comments)} 个批注")
            
        finally:
            document.Close()
            word_app.Quit()
        
        # 2. 构建提示词
        print("[2/3] 构建提示词...")
        from autoword.core.prompt_builder import PromptBuilder, PromptContext
        
        prompt_builder = PromptBuilder()
        context = PromptContext(
            document_structure=structure,
            comments=comments,
            document_path=test_doc
        )
        
        system_prompt = prompt_builder.build_system_prompt()
        user_prompt = prompt_builder.build_user_prompt(context)
        
        print(f"系统提示词长度: {len(system_prompt)}")
        print(f"用户提示词长度: {len(user_prompt)}")
        print()
        
        print("系统提示词:")
        print("=" * 50)
        print(system_prompt)
        print("=" * 50)
        print()
        
        print("用户提示词前500字符:")
        print("=" * 50)
        print(user_prompt[:500])
        print("=" * 50)
        print()
        
        # 3. 直接调用LLM
        print("[3/3] 直接调用LLM...")
        from autoword.core.llm_client import LLMClient, ModelType
        
        api_keys = {
            "claude": "sk-3w1JFbWUq7tKjpLlopdkISQ9F6fpLhHx5viD0frh43ESE9Io",
            "gpt": "sk-NhjnJtqlZMx4PGTqvkGlH4POT82HHBrBnBbWOat99Bs5VZXi"
        }
        
        llm_client = LLMClient(api_keys=api_keys)
        
        # 直接调用，不使用重试机制
        response = llm_client.call_model(ModelType.CLAUDE37, system_prompt, user_prompt)
        
        if response.success:
            print("原始LLM响应:")
            print("=" * 50)
            print(repr(response.content))
            print("=" * 50)
            print()
            
            print("原始LLM响应内容:")
            print("=" * 50)
            print(response.content)
            print("=" * 50)
            print()
            
            # 尝试解析
            try:
                parsed = json.loads(response.content)
                print("✅ 原始JSON解析成功!")
                return True
            except json.JSONDecodeError as e:
                print(f"❌ 原始JSON解析失败: {e}")
                print(f"错误位置: 第{e.lineno}行，第{e.colno}列")
                
                # 显示错误位置
                lines = response.content.split('\n')
                if e.lineno <= len(lines):
                    error_line = lines[e.lineno - 1]
                    print(f"错误行: {repr(error_line)}")
                    if e.colno <= len(error_line):
                        print(f"错误字符: {repr(error_line[max(0, e.colno-10):e.colno+10])}")
                        print(f"错误位置: {' ' * max(0, min(10, e.colno-1))}^")
                print()
                
                # 尝试修复
                print("尝试修复...")
                fixed = llm_client._aggressive_json_fix(response.content)
                print("修复后内容:")
                print("=" * 50)
                print(fixed)
                print("=" * 50)
                print()
                
                try:
                    parsed = json.loads(fixed)
                    print("✅ 修复后JSON解析成功!")
                    return True
                except json.JSONDecodeError as e2:
                    print(f"❌ 修复后仍然失败: {e2}")
                    return False
        else:
            print(f"❌ LLM调用失败: {response.error}")
            return False
            
    except Exception as e:
        print(f"❌ 调试过程中发生异常: {e}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == "__main__":
    try:
        success = debug_actual_response()
        if success:
            print()
            print("🎉 调试成功!")
            sys.exit(0)
        else:
            print()
            print("❌ 调试失败。")
            sys.exit(1)
    except KeyboardInterrupt:
        print("\n调试被用户中断")
        sys.exit(1)
    except Exception as e:
        print(f"调试过程中发生异常: {e}")
        sys.exit(1)