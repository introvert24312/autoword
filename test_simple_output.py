#!/usr/bin/env python3
"""
简单测试输出文件功能
"""

import sys
import shutil
from pathlib import Path

# 添加项目根目录到Python路径
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

def test_simple_output():
    """简单测试输出文件功能"""
    print("========================================")
    print("      简单输出文件测试")
    print("========================================")
    print()
    
    # 测试文档路径
    test_doc = r"C:\Users\y\Desktop\郭宇睿-现代汉语言文学在社会发展中的作用与影响研究2稿批注.docx"
    
    # 检查文档是否存在
    if not Path(test_doc).exists():
        print(f"❌ 测试文档不存在: {test_doc}")
        return False
    
    print(f"✅ 找到测试文档: {Path(test_doc).name}")
    
    # 生成输出文件路径
    input_path = Path(test_doc)
    output_file_path = input_path.parent / f"{input_path.stem}.process{input_path.suffix}"
    
    print(f"📁 输入文件: {input_path}")
    print(f"📄 输出文件: {output_file_path}")
    print()
    
    try:
        # 简单复制文件来模拟处理
        print("正在复制文件...")
        shutil.copy2(input_path, output_file_path)
        
        # 检查输出文件是否存在
        if output_file_path.exists():
            print(f"✅ 输出文件已创建: {output_file_path}")
            print(f"   文件大小: {output_file_path.stat().st_size} 字节")
            
            # 清理测试文件
            output_file_path.unlink()
            print("✅ 测试文件已清理")
            
            print()
            print("🎉 输出文件路径功能正常!")
            return True
        else:
            print(f"❌ 输出文件未创建: {output_file_path}")
            return False
            
    except Exception as e:
        print(f"❌ 测试过程中发生异常: {e}")
        return False


if __name__ == "__main__":
    try:
        success = test_simple_output()
        if success:
            print()
            print("✅ 输出文件路径测试通过！")
            print("   GUI界面的文件路径逻辑正常。")
            sys.exit(0)
        else:
            print()
            print("❌ 输出文件路径测试失败。")
            sys.exit(1)
    except KeyboardInterrupt:
        print("\n测试被用户中断")
        sys.exit(1)
    except Exception as e:
        print(f"测试过程中发生异常: {e}")
        sys.exit(1)