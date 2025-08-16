#!/usr/bin/env python3
"""
测试真实的任务规划场景
"""

import sys
from pathlib import Path

# 添加项目根目录到Python路径
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

def test_real_planning():
    """测试真实的任务规划场景"""
    print("========================================")
    print("      测试真实任务规划场景")
    print("========================================")
    print()
    
    # 测试文档路径
    test_doc = r"C:\Users\y\Desktop\郭宇睿-现代汉语言文学在社会发展中的作用与影响研究2稿批注.docx"
    
    # 检查文档是否存在
    if not Path(test_doc).exists():
        print(f"❌ 测试文档不存在: {test_doc}")
        return False
    
    print(f"✅ 找到测试文档: {Path(test_doc).name}")
    print()
    
    try:
        # 1. 加载文档和提取批注
        print("[1/4] 加载文档和提取批注...")
        from autoword.core.doc_loader import DocLoader
        from autoword.core.doc_inspector import DocInspector
        
        doc_loader = DocLoader()
        doc_inspector = DocInspector()
        
        word_app, document = doc_loader.load_document(test_doc, create_backup=False)
        
        try:
            comments = doc_inspector.extract_comments(document)
            structure = doc_inspector.extract_structure(document)
            
            print(f"✅ 提取到 {len(comments)} 个批注")
            print(f"✅ 文档结构: {structure.page_count} 页, {structure.word_count} 字")
            
            # 显示批注内容
            for i, comment in enumerate(comments, 1):
                print(f"  批注 {i}: {comment.comment_text[:50]}...")
            print()
            
        finally:
            document.Close()
            word_app.Quit()
        
        # 2. 测试任务规划器
        print("[2/4] 测试任务规划器...")
        from autoword.core.planner import TaskPlanner
        from autoword.core.llm_client import ModelType
        
        # 配置API密钥
        api_keys = {
            "claude": "sk-3w1JFbWUq7tKjpLlopdkISQ9F6fpLhHx5viD0frh43ESE9Io",
            "gpt": "sk-NhjnJtqlZMx4PGTqvkGlH4POT82HHBrBnBbWOat99Bs5VZXi"
        }
        
        task_planner = TaskPlanner(api_keys=api_keys)
        
        # 3. 生成任务计划
        print("[3/4] 生成任务计划...")
        
        try:
            planning_result = task_planner.generate_plan(
                structure, comments, test_doc, ModelType.CLAUDE37
            )
            
            if planning_result.success:
                print(f"✅ 任务规划成功!")
                print(f"   生成任务数量: {len(planning_result.task_plan.tasks)}")
                print(f"   规划耗时: {planning_result.planning_time:.2f}秒")
                
                # 显示生成的任务
                for i, task in enumerate(planning_result.task_plan.tasks, 1):
                    print(f"  任务 {i}: {task.type.value} - {task.instruction[:50]}...")
                
                return True
                
            else:
                print(f"❌ 任务规划失败: {planning_result.error_message}")
                return False
                
        except Exception as e:
            print(f"❌ 任务规划异常: {e}")
            import traceback
            traceback.print_exc()
            return False
        
        # 4. 测试完成
        print("[4/4] 测试完成")
        
    except Exception as e:
        print(f"❌ 测试过程中发生异常: {e}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == "__main__":
    try:
        success = test_real_planning()
        if success:
            print()
            print("🎉 真实任务规划测试成功!")
            print("   JSON解析问题已解决，可以正常使用GUI。")
            sys.exit(0)
        else:
            print()
            print("❌ 真实任务规划测试失败。")
            sys.exit(1)
    except KeyboardInterrupt:
        print("\n测试被用户中断")
        sys.exit(1)
    except Exception as e:
        print(f"测试过程中发生异常: {e}")
        sys.exit(1)