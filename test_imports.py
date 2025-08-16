#!/usr/bin/env python3
"""
测试AutoWord模块导入
"""

import sys
sys.path.append('.')

def test_imports():
    print("Testing AutoWord imports...")
    
    try:
        from autoword.core.enhanced_executor import EnhancedExecutor, WorkflowMode
        print("✅ EnhancedExecutor imported successfully")
        
        executor = EnhancedExecutor()
        print("✅ EnhancedExecutor created successfully")
        
        print("✅ All imports working correctly")
        return True
        
    except Exception as e:
        print(f"❌ Import error: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = test_imports()
    sys.exit(0 if success else 1)