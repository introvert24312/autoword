"""
Test AutoWord Document Loader
测试文档加载器
"""

import pytest
from unittest.mock import Mock, patch, MagicMock, PropertyMock
from pathlib import Path

from autoword.core.doc_loader import DocLoader, WordSession
from autoword.core.exceptions import DocumentError, COMError


class TestWordSession:
    """测试 WordSession 类"""
    
    @patch('autoword.core.doc_loader.pythoncom')
    @patch('autoword.core.doc_loader.win32')
    def test_word_session_context_manager(self, mock_win32, mock_pythoncom):
        """测试 Word 会话上下文管理器"""
        # 模拟 Word 应用程序
        mock_word_app = Mock()
        mock_win32.gencache.EnsureDispatch.return_value = mock_word_app
        
        # 测试上下文管理器
        with WordSession(visible=False) as word_app:
            assert word_app == mock_word_app
            assert word_app.Visible is False
            assert word_app.DisplayAlerts == 0
            assert word_app.ScreenUpdating is False
        
        # 验证清理调用
        mock_word_app.Quit.assert_called_once_with(SaveChanges=0)
        mock_pythoncom.CoInitialize.assert_called_once()
        mock_pythoncom.CoUninitialize.assert_called_once()
    
    @patch('autoword.core.doc_loader.pythoncom')
    @patch('autoword.core.doc_loader.win32')
    def test_word_session_com_error(self, mock_win32, mock_pythoncom):
        """测试 COM 初始化错误"""
        mock_pythoncom.CoInitialize.side_effect = Exception("COM Error")
        
        with pytest.raises(COMError, match="Failed to initialize Word COM"):
            with WordSession():
                pass
    
    @patch('autoword.core.doc_loader.pythoncom')
    @patch('autoword.core.doc_loader.win32')
    def test_word_session_cleanup_on_error(self, mock_win32, mock_pythoncom):
        """测试错误时的清理"""
        mock_word_app = Mock()
        mock_win32.gencache.EnsureDispatch.return_value = mock_word_app
        # 设置 DisplayAlerts 属性赋值时抛出异常
        type(mock_word_app).DisplayAlerts = PropertyMock(side_effect=Exception("Word Error"))
        
        with pytest.raises(COMError):
            with WordSession():
                pass
        
        # 验证清理仍然被调用
        mock_pythoncom.CoUninitialize.assert_called_once()


class TestDocLoader:
    """测试 DocLoader 类"""
    
    def setup_method(self):
        """测试前设置"""
        self.loader = DocLoader(visible=False)
    
    @patch('autoword.core.doc_loader.shutil.copy2')
    @patch('autoword.core.doc_loader.validate_document_path')
    @patch('autoword.core.doc_loader.generate_backup_path')
    def test_create_backup_success(self, mock_generate_path, mock_validate, mock_copy):
        """测试成功创建备份"""
        # 设置模拟
        mock_validate.return_value = Path("test.docx")
        mock_generate_path.return_value = "test_backup_20240101_120000.docx"
        
        # 执行测试
        backup_path = self.loader.create_backup("test.docx")
        
        # 验证结果
        assert backup_path == "test_backup_20240101_120000.docx"
        mock_validate.assert_called_once_with("test.docx")
        mock_generate_path.assert_called_once()
        mock_copy.assert_called_once()
    
    @patch('autoword.core.doc_loader.shutil.copy2')
    @patch('autoword.core.doc_loader.validate_document_path')
    def test_create_backup_failure(self, mock_validate, mock_copy):
        """测试备份创建失败"""
        mock_validate.return_value = Path("test.docx")
        mock_copy.side_effect = Exception("Copy failed")
        
        with pytest.raises(DocumentError, match="Failed to create backup"):
            self.loader.create_backup("test.docx")
    
    @patch.object(DocLoader, 'create_backup')
    @patch('autoword.core.doc_loader.WordSession')
    @patch('autoword.core.doc_loader.validate_document_path')
    def test_load_document_success(self, mock_validate, mock_word_session, mock_backup):
        """测试成功加载文档"""
        # 设置模拟
        mock_validate.return_value = Path("test.docx")
        mock_backup.return_value = "backup.docx"
        
        mock_word_app = Mock()
        mock_document = Mock()
        mock_document.TrackRevisions = True
        
        mock_session_instance = MagicMock()
        mock_session_instance.__enter__.return_value = mock_word_app
        mock_session_instance.__exit__.return_value = None
        mock_word_session.return_value = mock_session_instance
        
        mock_word_app.Documents.Open.return_value = mock_document
        
        # 执行测试
        word_app, document = self.loader.load_document("test.docx")
        
        # 验证结果
        assert word_app == mock_word_app
        assert document == mock_document
        assert document.TrackRevisions is False
        assert document.ShowGrammaticalErrors is False
        assert document.ShowSpellingErrors is False
        
        mock_backup.assert_called_once_with("test.docx")
        mock_word_app.Documents.Open.assert_called_once()
    
    @patch.object(DocLoader, 'create_backup')
    @patch('autoword.core.doc_loader.WordSession')
    @patch('autoword.core.doc_loader.validate_document_path')
    def test_load_document_no_backup(self, mock_validate, mock_word_session, mock_backup):
        """测试不创建备份的文档加载"""
        mock_validate.return_value = Path("test.docx")
        
        mock_word_app = Mock()
        mock_document = Mock()
        mock_document.TrackRevisions = False
        
        mock_session_instance = MagicMock()
        mock_session_instance.__enter__.return_value = mock_word_app
        mock_session_instance.__exit__.return_value = None
        mock_word_session.return_value = mock_session_instance
        
        mock_word_app.Documents.Open.return_value = mock_document
        
        # 执行测试（不创建备份）
        word_app, document = self.loader.load_document("test.docx", create_backup=False)
        
        # 验证结果
        assert word_app == mock_word_app
        assert document == mock_document
        mock_backup.assert_not_called()
    
    @patch('autoword.core.doc_loader.WordSession')
    @patch('autoword.core.doc_loader.validate_document_path')
    def test_load_document_open_failure(self, mock_validate, mock_word_session):
        """测试文档打开失败"""
        mock_validate.return_value = Path("test.docx")
        
        mock_word_app = Mock()
        mock_word_app.Documents.Open.side_effect = Exception("Open failed")
        
        mock_session_instance = MagicMock()
        mock_session_instance.__enter__.return_value = mock_word_app
        mock_session_instance.__exit__.return_value = None
        mock_word_session.return_value = mock_session_instance
        
        with pytest.raises(DocumentError, match="Failed to open document"):
            self.loader.load_document("test.docx", create_backup=False)
    
    @patch.object(DocLoader, 'create_backup')
    @patch('autoword.core.doc_loader.WordSession')
    @patch('autoword.core.doc_loader.validate_document_path')
    def test_open_document_context_manager(self, mock_validate, mock_word_session, mock_backup):
        """测试文档上下文管理器"""
        mock_validate.return_value = Path("test.docx")
        mock_backup.return_value = "backup.docx"
        
        mock_word_app = Mock()
        mock_document = Mock()
        mock_document.TrackRevisions = True
        
        mock_session_instance = MagicMock()
        mock_session_instance.__enter__.return_value = mock_word_app
        mock_session_instance.__exit__.return_value = None
        mock_word_session.return_value = mock_session_instance
        
        mock_word_app.Documents.Open.return_value = mock_document
        
        # 测试上下文管理器
        with self.loader.open_document("test.docx") as (word_app, document):
            assert word_app == mock_word_app
            assert document == mock_document
            assert document.TrackRevisions is False
        
        # 验证清理被调用
        mock_session_instance.__exit__.assert_called_once()
    
    def test_save_document_default_location(self):
        """测试保存文档到默认位置"""
        mock_document = Mock()
        mock_document.FullName = "test.docx"
        
        result = self.loader.save_document(mock_document)
        
        assert result == "test.docx"
        mock_document.Save.assert_called_once()
    
    def test_save_document_custom_location(self):
        """测试保存文档到指定位置"""
        mock_document = Mock()
        
        result = self.loader.save_document(mock_document, "new_path.docx")
        
        assert result == "new_path.docx"
        mock_document.SaveAs2.assert_called_once_with(FileName="new_path.docx")
    
    def test_save_document_failure(self):
        """测试保存文档失败"""
        mock_document = Mock()
        mock_document.Save.side_effect = Exception("Save failed")
        
        with pytest.raises(DocumentError, match="Failed to save document"):
            self.loader.save_document(mock_document)
    
    @patch('autoword.core.doc_loader.win32_constants')
    def test_close_document(self, mock_constants):
        """测试关闭文档"""
        mock_constants.wdSaveChanges = 1
        mock_constants.wdDoNotSaveChanges = 0
        
        mock_document = Mock()
        
        # 测试保存更改
        self.loader.close_document(mock_document, save_changes=True)
        mock_document.Close.assert_called_with(SaveChanges=1)
        
        # 测试不保存更改
        self.loader.close_document(mock_document, save_changes=False)
        mock_document.Close.assert_called_with(SaveChanges=0)
    
    @patch('autoword.core.doc_loader.WordSession')
    def test_check_word_availability_success(self, mock_word_session):
        """测试 Word 可用性检查成功"""
        mock_word_app = Mock()
        mock_word_app.Version = "16.0"
        
        mock_session_instance = MagicMock()
        mock_session_instance.__enter__.return_value = mock_word_app
        mock_session_instance.__exit__.return_value = None
        mock_word_session.return_value = mock_session_instance
        
        result = self.loader.check_word_availability()
        
        assert result is True
    
    @patch('autoword.core.doc_loader.WordSession')
    def test_check_word_availability_failure(self, mock_word_session):
        """测试 Word 可用性检查失败"""
        mock_word_session.side_effect = Exception("Word not available")
        
        result = self.loader.check_word_availability()
        
        assert result is False
    
    @patch('autoword.core.doc_loader.WordSession')
    def test_get_word_version_success(self, mock_word_session):
        """测试获取 Word 版本成功"""
        mock_word_app = Mock()
        mock_word_app.Version = "16.0"
        
        mock_session_instance = MagicMock()
        mock_session_instance.__enter__.return_value = mock_word_app
        mock_session_instance.__exit__.return_value = None
        mock_word_session.return_value = mock_session_instance
        
        version = self.loader.get_word_version()
        
        assert version == "16.0"
    
    @patch('autoword.core.doc_loader.WordSession')
    def test_get_word_version_failure(self, mock_word_session):
        """测试获取 Word 版本失败"""
        mock_word_session.side_effect = Exception("Word not available")
        
        version = self.loader.get_word_version()
        
        assert version is None


if __name__ == "__main__":
    pytest.main([__file__])