#!/usr/bin/env python3
"""
AutoWord GUI 测试脚本
测试GUI组件的基本功能
"""

import sys
import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

# 添加项目根目录到Python路径
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

# 设置Qt应用程序测试环境
os.environ['QT_QPA_PLATFORM'] = 'offscreen'

try:
    from PySide6.QtWidgets import QApplication
    from PySide6.QtTest import QTest
    from PySide6.QtCore import Qt
    
    # 创建QApplication实例（测试需要）
    if not QApplication.instance():
        app = QApplication(sys.argv)
    
    from autoword.gui.config_manager import ConfigurationManager, AppConfig, SecureStorage
    from autoword.gui.error_handler import ErrorHandler
    from autoword.gui.processor_controller import DocumentProcessorController
    
    GUI_AVAILABLE = True
except ImportError as e:
    print(f"GUI测试跳过，缺少依赖: {e}")
    GUI_AVAILABLE = False


class TestSecureStorage(unittest.TestCase):
    """测试安全存储功能"""
    
    def setUp(self):
        """设置测试环境"""
        if not GUI_AVAILABLE:
            self.skipTest("GUI依赖不可用")
        
        # 使用临时目录
        self.temp_dir = tempfile.mkdtemp()
        self.storage = SecureStorage()
        # 重定向密钥文件到临时目录
        self.storage.key_file = Path(self.temp_dir) / ".test_key"
        self.storage.key = self.storage._get_or_create_key()
        self.storage.cipher = self.storage.__class__().__dict__['__init__'](self.storage)
    
    def test_encrypt_decrypt(self):
        """测试加密解密功能"""
        test_data = "test_api_key_12345"
        
        # 加密
        encrypted = self.storage.encrypt(test_data)
        self.assertNotEqual(encrypted, test_data)
        self.assertTrue(len(encrypted) > 0)
        
        # 解密
        decrypted = self.storage.decrypt(encrypted)
        self.assertEqual(decrypted, test_data)
    
    def test_empty_data(self):
        """测试空数据处理"""
        self.assertEqual(self.storage.encrypt(""), "")
        self.assertEqual(self.storage.decrypt(""), "")
    
    def tearDown(self):
        """清理测试环境"""
        import shutil
        if hasattr(self, 'temp_dir'):
            shutil.rmtree(self.temp_dir, ignore_errors=True)


class TestConfigurationManager(unittest.TestCase):
    """测试配置管理器"""
    
    def setUp(self):
        """设置测试环境"""
        if not GUI_AVAILABLE:
            self.skipTest("GUI依赖不可用")
        
        # 使用临时目录
        self.temp_dir = tempfile.mkdtemp()
        
        # 创建配置管理器并重定向配置目录
        self.config_manager = ConfigurationManager()
        self.config_manager.config_dir = Path(self.temp_dir)
        self.config_manager.config_file = self.config_manager.config_dir / "test_config.json"
        
        # 重定向安全存储
        self.config_manager.secure_storage.key_file = Path(self.temp_dir) / ".test_key"
    
    def test_default_config(self):
        """测试默认配置"""
        config = self.config_manager.config
        self.assertIsInstance(config, AppConfig)
        self.assertEqual(config.selected_model, "claude")
        self.assertEqual(config.auto_save_settings, True)
    
    def test_api_key_management(self):
        """测试API密钥管理"""
        test_key = "test_api_key_12345"
        
        # 设置API密钥
        self.config_manager.set_api_key("claude", test_key)
        
        # 获取API密钥
        retrieved_key = self.config_manager.get_api_key("claude")
        self.assertEqual(retrieved_key, test_key)
    
    def test_model_selection(self):
        """测试模型选择"""
        # 设置模型
        self.config_manager.set_selected_model("gpt")
        self.assertEqual(self.config_manager.get_selected_model(), "gpt")
        
        # 获取可用模型
        models = self.config_manager.get_available_models()
        self.assertIn("claude", models)
        self.assertIn("gpt", models)
    
    def test_path_management(self):
        """测试路径管理"""
        test_input = "/test/input.docx"
        test_output = "/test/output/"
        
        # 设置路径
        self.config_manager.set_last_paths(test_input, test_output)
        
        # 获取路径
        paths = self.config_manager.get_last_paths()
        self.assertEqual(paths['input'], test_input)
        self.assertEqual(paths['output'], test_output)
    
    def test_config_validation(self):
        """测试配置验证"""
        # 无效配置
        errors = self.config_manager.validate_config()
        self.assertIn('api_key', errors)
        
        # 有效配置
        self.config_manager.set_api_key("claude", "valid_key_12345")
        errors = self.config_manager.validate_config()
        self.assertNotIn('api_key', errors)
    
    def tearDown(self):
        """清理测试环境"""
        import shutil
        if hasattr(self, 'temp_dir'):
            shutil.rmtree(self.temp_dir, ignore_errors=True)


class TestErrorHandler(unittest.TestCase):
    """测试错误处理器"""
    
    def setUp(self):
        """设置测试环境"""
        if not GUI_AVAILABLE:
            self.skipTest("GUI依赖不可用")
        
        self.error_handler = ErrorHandler()
    
    def test_error_logging(self):
        """测试错误记录"""
        error_type = "TestError"
        error_message = "This is a test error"
        
        # 记录错误
        self.error_handler.log_error_info(error_type, error_message)
        
        # 检查错误历史
        self.assertEqual(len(self.error_handler.error_history), 1)
        self.assertEqual(self.error_handler.error_history[0]['type'], error_type)
        self.assertEqual(self.error_handler.error_history[0]['message'], error_message)
    
    def test_exception_logging(self):
        """测试异常记录"""
        try:
            raise ValueError("Test exception")
        except Exception as e:
            self.error_handler.log_error(e, "Test context")
        
        # 检查错误历史
        self.assertEqual(len(self.error_handler.error_history), 1)
        self.assertEqual(self.error_handler.error_history[0]['type'], "ValueError")
        self.assertIn("Test context", self.error_handler.error_history[0]['message'])
    
    def test_error_statistics(self):
        """测试错误统计"""
        # 添加多个错误
        self.error_handler.log_error_info("Error1", "Message 1")
        self.error_handler.log_error_info("Error2", "Message 2")
        self.error_handler.log_error_info("Error1", "Message 3")
        
        # 获取统计信息
        stats = self.error_handler.get_error_statistics()
        
        self.assertEqual(stats['total'], 3)
        self.assertEqual(stats['by_type']['Error1'], 2)
        self.assertEqual(stats['by_type']['Error2'], 1)
        self.assertEqual(len(stats['recent']), 3)
    
    def test_clear_history(self):
        """测试清空历史"""
        # 添加错误
        self.error_handler.log_error_info("TestError", "Test message")
        self.assertEqual(len(self.error_handler.error_history), 1)
        
        # 清空历史
        self.error_handler.clear_error_history()
        self.assertEqual(len(self.error_handler.error_history), 0)


class TestDocumentProcessorController(unittest.TestCase):
    """测试文档处理控制器"""
    
    def setUp(self):
        """设置测试环境"""
        if not GUI_AVAILABLE:
            self.skipTest("GUI依赖不可用")
        
        # 创建模拟配置管理器
        self.config_manager = Mock()
        self.controller = DocumentProcessorController(self.config_manager)
    
    def test_parameter_validation(self):
        """测试参数验证"""
        # 无效参数
        errors = self.controller._validate_parameters("", "", "invalid", "")
        self.assertTrue(len(errors) > 0)
        
        # 创建临时文件进行测试
        with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as temp_file:
            temp_path = temp_file.name
        
        try:
            # 有效参数
            errors = self.controller._validate_parameters(
                temp_path, 
                tempfile.gettempdir(), 
                "claude", 
                "valid_api_key_12345"
            )
            # 应该只有Word格式检查错误（因为是空文件）
            self.assertTrue(any("Word文档格式" in error for error in errors) or len(errors) == 0)
        finally:
            os.unlink(temp_path)
    
    def test_processing_status(self):
        """测试处理状态"""
        status = self.controller.get_processing_status()
        
        self.assertIn('is_running', status)
        self.assertIn('current_stage', status)
        self.assertIn('progress', status)
        self.assertEqual(status['is_running'], False)


def run_gui_tests():
    """运行GUI测试"""
    if not GUI_AVAILABLE:
        print("跳过GUI测试 - 缺少必要依赖")
        return True
    
    # 创建测试套件
    test_suite = unittest.TestSuite()
    
    # 添加测试用例
    test_classes = [
        TestSecureStorage,
        TestConfigurationManager,
        TestErrorHandler,
        TestDocumentProcessorController,
    ]
    
    for test_class in test_classes:
        tests = unittest.TestLoader().loadTestsFromTestCase(test_class)
        test_suite.addTests(tests)
    
    # 运行测试
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(test_suite)
    
    return result.wasSuccessful()


def test_imports():
    """测试模块导入"""
    print("测试模块导入...")
    
    try:
        from autoword.gui import MainWindow, ConfigurationManager, DocumentProcessorController, ErrorHandler
        print("✓ GUI模块导入成功")
        return True
    except ImportError as e:
        print(f"✗ GUI模块导入失败: {e}")
        return False


def main():
    """主函数"""
    print("AutoWord GUI 测试")
    print("=" * 30)
    
    success = True
    
    # 测试导入
    if not test_imports():
        success = False
    
    # 运行GUI测试
    if GUI_AVAILABLE:
        print("\n运行GUI功能测试...")
        if not run_gui_tests():
            success = False
    else:
        print("\n跳过GUI功能测试 - 缺少依赖")
    
    print("\n" + "=" * 30)
    if success:
        print("✓ 所有测试通过")
        return 0
    else:
        print("✗ 部分测试失败")
        return 1


if __name__ == "__main__":
    sys.exit(main())