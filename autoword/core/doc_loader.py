"""
AutoWord Document Loader
Word COM 文档加载器
"""

import os
import shutil
import logging
from typing import Optional, Any
from pathlib import Path
from contextlib import contextmanager

import pythoncom
import win32com.client as win32
from win32com.client import constants as win32_constants

from .models import DocumentSnapshot
from .utils import generate_backup_path, validate_document_path, calculate_file_checksum
from .exceptions import DocumentError, COMError


logger = logging.getLogger(__name__)


class WordSession:
    """Word COM 会话管理器"""
    
    def __init__(self, visible: bool = False):
        """
        初始化 Word 会话
        
        Args:
            visible: 是否显示 Word 窗口
        """
        self.visible = visible
        self.word_app: Optional[Any] = None
        self.com_initialized = False
    
    def __enter__(self):
        """进入上下文管理器"""
        try:
            # 初始化 COM
            pythoncom.CoInitialize()
            self.com_initialized = True
            logger.debug("COM initialized successfully")
            
            # 创建 Word 应用程序实例
            self.word_app = win32.gencache.EnsureDispatch('Word.Application')
            self.word_app.Visible = self.visible
            
            # 禁用一些可能干扰的功能
            self.word_app.DisplayAlerts = 0  # 禁用警告对话框
            self.word_app.ScreenUpdating = False  # 禁用屏幕更新以提高性能
            
            logger.info(f"Word application started (visible={self.visible})")
            return self.word_app
            
        except Exception as e:
            self._cleanup()
            raise COMError(f"Failed to initialize Word COM: {e}")
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """退出上下文管理器"""
        self._cleanup()
    
    def _cleanup(self):
        """清理资源"""
        try:
            if self.word_app:
                # 恢复设置
                self.word_app.ScreenUpdating = True
                self.word_app.DisplayAlerts = -1  # 恢复默认警告
                
                # 退出 Word 应用程序
                self.word_app.Quit(SaveChanges=0)  # 不保存更改
                self.word_app = None
                logger.info("Word application closed")
                
        except Exception as e:
            logger.warning(f"Error during Word cleanup: {e}")
        
        finally:
            if self.com_initialized:
                try:
                    pythoncom.CoUninitialize()
                    self.com_initialized = False
                    logger.debug("COM uninitialized")
                except Exception as e:
                    logger.warning(f"Error during COM cleanup: {e}")


class DocLoader:
    """Word 文档加载器"""
    
    def __init__(self, visible: bool = False):
        """
        初始化文档加载器
        
        Args:
            visible: 是否显示 Word 窗口
        """
        self.visible = visible
    
    def create_backup(self, document_path: str) -> str:
        """
        创建文档备份
        
        Args:
            document_path: 原始文档路径
            
        Returns:
            备份文件路径
            
        Raises:
            DocumentError: 备份创建失败时抛出
        """
        try:
            # 验证原始文档
            original_path = validate_document_path(document_path)
            
            # 生成备份路径
            backup_path = generate_backup_path(str(original_path))
            
            # 复制文件
            shutil.copy2(original_path, backup_path)
            
            logger.info(f"Backup created: {backup_path}")
            return backup_path
            
        except Exception as e:
            raise DocumentError(f"Failed to create backup: {e}")
    
    def load_document(self, document_path: str, create_backup: bool = True) -> tuple[Any, Any]:
        """
        加载 Word 文档
        
        Args:
            document_path: 文档路径
            create_backup: 是否创建备份
            
        Returns:
            (word_app, document) 元组
            
        Raises:
            DocumentError: 文档加载失败时抛出
        """
        try:
            # 验证文档路径
            doc_path = validate_document_path(document_path)
            
            # 创建备份
            backup_path = None
            if create_backup:
                backup_path = self.create_backup(str(doc_path))
            
            # 初始化 Word COM
            word_session = WordSession(visible=self.visible)
            word_app = word_session.__enter__()
            
            try:
                # 打开文档
                document = word_app.Documents.Open(
                    FileName=str(doc_path),
                    ReadOnly=False,
                    AddToRecentFiles=False,
                    Visible=self.visible
                )
                
                # 禁用修订跟踪
                if document.TrackRevisions:
                    document.TrackRevisions = False
                    logger.info("Track changes disabled")
                
                # 禁用拼写和语法检查以提高性能
                document.ShowGrammaticalErrors = False
                document.ShowSpellingErrors = False
                
                logger.info(f"Document loaded: {doc_path}")
                
                return word_app, document
                
            except Exception as e:
                word_session.__exit__(None, None, None)
                raise DocumentError(f"Failed to open document: {e}")
                
        except Exception as e:
            if isinstance(e, DocumentError):
                raise
            else:
                raise DocumentError(f"Failed to load document: {e}")
    
    @contextmanager
    def open_document(self, document_path: str, create_backup: bool = True):
        """
        上下文管理器方式打开文档
        
        Args:
            document_path: 文档路径
            create_backup: 是否创建备份
            
        Yields:
            (word_app, document) 元组
        """
        word_session = None
        try:
            # 验证文档路径
            doc_path = validate_document_path(document_path)
            
            # 创建备份
            if create_backup:
                self.create_backup(str(doc_path))
            
            # 创建 Word 会话
            word_session = WordSession(visible=self.visible)
            word_app = word_session.__enter__()
            
            # 打开文档
            document = word_app.Documents.Open(
                FileName=str(doc_path),
                ReadOnly=False,
                AddToRecentFiles=False,
                Visible=self.visible
            )
            
            # 配置文档
            if document.TrackRevisions:
                document.TrackRevisions = False
            
            document.ShowGrammaticalErrors = False
            document.ShowSpellingErrors = False
            
            logger.info(f"Document opened in context: {doc_path}")
            
            yield word_app, document
            
        except Exception as e:
            raise DocumentError(f"Failed to open document in context: {e}")
        
        finally:
            if word_session:
                word_session.__exit__(None, None, None)
    
    def save_document(self, document: Any, save_path: Optional[str] = None) -> str:
        """
        保存文档
        
        Args:
            document: Word 文档对象
            save_path: 保存路径（可选，默认保存到原位置）
            
        Returns:
            保存的文件路径
            
        Raises:
            DocumentError: 保存失败时抛出
        """
        try:
            if save_path:
                # 另存为
                document.SaveAs2(FileName=save_path)
                logger.info(f"Document saved as: {save_path}")
                return save_path
            else:
                # 保存到原位置
                document.Save()
                saved_path = document.FullName
                logger.info(f"Document saved: {saved_path}")
                return saved_path
                
        except Exception as e:
            raise DocumentError(f"Failed to save document: {e}")
    
    def close_document(self, document: Any, save_changes: bool = True):
        """
        关闭文档
        
        Args:
            document: Word 文档对象
            save_changes: 是否保存更改
        """
        try:
            if document:
                save_option = win32_constants.wdSaveChanges if save_changes else win32_constants.wdDoNotSaveChanges
                document.Close(SaveChanges=save_option)
                logger.info("Document closed")
                
        except Exception as e:
            logger.warning(f"Error closing document: {e}")
    
    def check_word_availability(self) -> bool:
        """
        检查 Word COM 是否可用
        
        Returns:
            True 如果 Word 可用，否则 False
        """
        try:
            with WordSession(visible=False) as word_app:
                version = word_app.Version
                logger.info(f"Word COM available, version: {version}")
                return True
                
        except Exception as e:
            logger.error(f"Word COM not available: {e}")
            return False
    
    def get_word_version(self) -> Optional[str]:
        """
        获取 Word 版本信息
        
        Returns:
            Word 版本字符串，如果不可用则返回 None
        """
        try:
            with WordSession(visible=False) as word_app:
                return word_app.Version
                
        except Exception:
            return None