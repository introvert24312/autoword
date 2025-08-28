#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
测试GUI集成封面保护功能
"""

import sys
import os
from pathlib import Path

# 添加项目根目录到Python路径
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

def test_gui_integration():
    """测试GUI集成"""
    print("🧪 测试GUI集成封面保护功能")
    print("=" * 50)
    
    try:
        # 测试导入
        print("1. 测试模块导入...")
        from 简化版GUI import AutoWordGUI
        print("   ✅ GUI模块导入成功")
        
        from autoword.vnext.simple_pipeline import SimplePipeline
        from autoword.vnext.core import VNextConfig
        print("   ✅ SimplePipeline导入成功")
        
        # 测试配置创建
        print("\n2. 测试配置创建...")
        config = VNextConfig()
        pipeline = SimplePipeline(config)
        print("   ✅ SimplePipeline实例创建成功")
        
        # 测试GUI创建（不显示）
        print("\n3. 测试GUI创建...")
        import tkinter as tk
        root = tk.Tk()
        root.withdraw()  # 隐藏窗口
        
        # 创建GUI实例
        app = AutoWordGUI()
        app.root.withdraw()  # 隐藏GUI窗口
        
        print("   ✅ GUI实例创建成功")
        
        # 测试封面保护选项
        print("\n4. 测试封面保护选项...")
        if hasattr(app, 'cover_protection'):
            print(f"   ✅ 封面保护选项存在: {app.cover_protection.get()}")
        else:
            print("   ❌ 封面保护选项不存在")
        
        if hasattr(app, 'auto_page_break'):
            print(f"   ✅ 自动分页符选项存在: {app.auto_page_break.get()}")
        else:
            print("   ❌ 自动分页符选项不存在")
        
        if hasattr(app, 'cover_status'):
            print(f"   ✅ 封面状态显示存在: {app.cover_status.get()}")
        else:
            print("   ❌ 封面状态显示不存在")
        
        # 测试方法存在性
        print("\n5. 测试方法存在性...")
        methods_to_check = [
            'toggle_cover_protection',
            'test_cover_protection'
        ]
        
        for method_name in methods_to_check:
            if hasattr(app, method_name):
                print(f"   ✅ 方法 {method_name} 存在")
            else:
                print(f"   ❌ 方法 {method_name} 不存在")
        
        # 清理
        app.root.destroy()
        root.destroy()
        
        print("\n🎉 GUI集成测试完成！")
        print("现在可以运行 'python 简化版GUI.py' 来使用带封面保护功能的GUI")
        
    except Exception as e:
        print(f"\n❌ 测试失败: {e}")
        import traceback
        traceback.print_exc()

def show_usage_instructions():
    """显示使用说明"""
    print("\n" + "=" * 60)
    print("📖 使用说明")
    print("=" * 60)
    print("1. 运行GUI: python 简化版GUI.py")
    print("2. 在'批注处理'区域可以看到新增的选项:")
    print("   - 🛡️ 保护封面和目录格式 (默认启用)")
    print("   - 📄 自动插入分页符")
    print("3. 新增的'🛡️ 测试封面保护'按钮可以快速测试功能")
    print("4. 处理文档时会自动保护封面和目录区域不被修改")
    print("\n💡 功能特点:")
    print("- 智能识别封面和目录区域（无页码区域）")
    print("- 基于分页符位置判断正文开始位置")
    print("- 只修改正文区域的样式，保护封面格式")
    print("- 支持自动插入分页符分隔封面和正文")

if __name__ == "__main__":
    test_gui_integration()
    show_usage_instructions()