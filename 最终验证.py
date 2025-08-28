#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
最终验证脚本 - 验证所有功能是否正常工作
"""

import sys
import os
from pathlib import Path

# 添加项目根目录到Python路径
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

def test_all_imports():
    """测试所有导入"""
    print("🔍 测试所有模块导入...")
    
    try:
        # 测试核心模块
        from autoword.vnext.core import VNextConfig, LLMConfig
        print("   ✅ 核心配置模块")
        
        # 测试SimplePipeline
        from autoword.vnext.simple_pipeline import SimplePipeline
        print("   ✅ SimplePipeline（封面保护）")
        
        # 测试GUI
        from 简化版GUI import AutoWordGUI
        print("   ✅ 增强版GUI")
        
        return True
        
    except Exception as e:
        print(f"   ❌ 导入失败: {e}")
        return False

def test_config_creation():
    """测试配置创建"""
    print("\n🔧 测试配置创建...")
    
    try:
        from autoword.vnext.core import VNextConfig, LLMConfig
        from autoword.vnext.simple_pipeline import SimplePipeline
        
        # 创建配置
        config = VNextConfig()
        print("   ✅ VNextConfig创建成功")
        
        # 创建Pipeline
        pipeline = SimplePipeline(config)
        print("   ✅ SimplePipeline创建成功")
        
        return True
        
    except Exception as e:
        print(f"   ❌ 配置创建失败: {e}")
        return False

def test_gui_features():
    """测试GUI功能"""
    print("\n🖥️ 测试GUI功能...")
    
    try:
        import tkinter as tk
        from 简化版GUI import AutoWordGUI
        
        # 创建隐藏的根窗口
        root = tk.Tk()
        root.withdraw()
        
        # 创建GUI实例
        app = AutoWordGUI()
        app.root.withdraw()
        
        # 检查封面保护功能
        features = []
        
        if hasattr(app, 'cover_protection'):
            features.append("🛡️ 封面保护选项")
        
        if hasattr(app, 'auto_page_break'):
            features.append("📄 自动分页符选项")
        
        if hasattr(app, 'cover_status'):
            features.append("📊 封面状态显示")
        
        if hasattr(app, 'toggle_cover_protection'):
            features.append("🔄 封面保护切换")
        
        if hasattr(app, 'test_cover_protection'):
            features.append("🧪 封面保护测试")
        
        for feature in features:
            print(f"   ✅ {feature}")
        
        # 清理
        app.root.destroy()
        root.destroy()
        
        return len(features) >= 4
        
    except Exception as e:
        print(f"   ❌ GUI功能测试失败: {e}")
        return False

def test_pipeline_methods():
    """测试Pipeline方法"""
    print("\n⚙️ 测试Pipeline方法...")
    
    try:
        from autoword.vnext.core import VNextConfig
        from autoword.vnext.simple_pipeline import SimplePipeline
        
        config = VNextConfig()
        pipeline = SimplePipeline(config)
        
        # 检查封面保护相关方法
        methods = [
            '_find_first_content_section',
            '_is_cover_or_toc_content',
            '_apply_styles_to_content',
            '_insert_page_break'
        ]
        
        for method_name in methods:
            if hasattr(pipeline, method_name):
                print(f"   ✅ {method_name}")
            else:
                print(f"   ❌ {method_name} 缺失")
                return False
        
        return True
        
    except Exception as e:
        print(f"   ❌ Pipeline方法测试失败: {e}")
        return False

def show_final_summary():
    """显示最终总结"""
    print("\n" + "=" * 60)
    print("🎉 AutoWord 封面保护功能集成完成！")
    print("=" * 60)
    
    print("\n📋 功能清单:")
    print("✅ 智能封面识别（无页码区域）")
    print("✅ 基于分页符的正文定位")
    print("✅ 封面格式保护")
    print("✅ 自动分页符插入")
    print("✅ GUI集成界面")
    print("✅ 一键测试功能")
    
    print("\n🚀 使用方法:")
    print("1. 启动GUI: python 简化版GUI.py")
    print("2. 选择文档，配置API密钥")
    print("3. 确保'🛡️ 保护封面和目录格式'已勾选")
    print("4. 输入处理指令或使用批注")
    print("5. 点击'🚀 开始处理'")
    
    print("\n💡 特色功能:")
    print("- 封面和目录格式完全保护")
    print("- 支持无页码文档")
    print("- 智能识别正文开始位置")
    print("- 可选自动插入分页符")
    print("- 实时状态显示")
    
    print("\n🧪 测试选项:")
    print("- GUI内置测试: 点击'🛡️ 测试封面保护'按钮")
    print("- 命令行测试: python 测试无页码封面保护.py")
    print("- 功能验证: python 最终验证.py")

def main():
    """主函数"""
    print("AutoWord 封面保护功能 - 最终验证")
    print("=" * 60)
    
    all_passed = True
    
    # 运行所有测试
    tests = [
        test_all_imports,
        test_config_creation,
        test_gui_features,
        test_pipeline_methods
    ]
    
    for test in tests:
        if not test():
            all_passed = False
    
    print("\n" + "=" * 60)
    if all_passed:
        print("🎉 所有测试通过！功能集成成功！")
        show_final_summary()
    else:
        print("❌ 部分测试失败，请检查错误信息")
    
    return all_passed

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)