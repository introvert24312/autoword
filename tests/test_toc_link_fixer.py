"""
Test AutoWord TOC and Link Fixer
测试 TOC 和超链接管理器
"""

import pytest
from unittest.mock import Mock, patch, MagicMock

from autoword.core.toc_link_fixer import (
    TocManager, HyperlinkManager, TocAndLinkFixer,
    TocConfig, TocStyle, LinkType, LinkValidationResult
)
from autoword.core.models import Hyperlink, ValidationResult
from autoword.core.exceptions import TaskExecutionError


class TestTocConfig:
    """测试目录配置"""
    
    def test_toc_config_defaults(self):
        """测试目录配置默认值"""
        config = TocConfig()
        
        assert config.use_heading_styles is True
        assert config.upper_heading_level == 1
        assert config.lower_heading_level == 3
        assert config.use_hyperlinks is True
        assert config.right_align_page_numbers is True
        assert config.include_page_numbers is True
        assert config.style == TocStyle.CLASSIC
    
    def test_toc_config_custom(self):
        """测试自定义目录配置"""
        config = TocConfig(
            upper_heading_level=2,
            lower_heading_level=4,
            style=TocStyle.MODERN,
            use_hyperlinks=False
        )
        
        assert config.upper_heading_level == 2
        assert config.lower_heading_level == 4
        assert config.style == TocStyle.MODERN
        assert config.use_hyperlinks is False


class TestTocManager:
    """测试目录管理器"""
    
    def setup_method(self):
        """测试前设置"""
        self.mock_word_app = Mock()
        self.mock_document = Mock()
        self.manager = TocManager(self.mock_word_app, self.mock_document)
    
    def test_create_toc_with_default_config(self):
        """测试使用默认配置创建目录"""
        mock_range = Mock()
        mock_toc = Mock()
        
        self.mock_document.TablesOfContents.Add.return_value = mock_toc
        
        result = self.manager.create_toc(mock_range)
        
        assert result == mock_toc
        self.mock_document.TablesOfContents.Add.assert_called_once()
        
        # 验证调用参数
        call_args = self.mock_document.TablesOfContents.Add.call_args
        assert call_args.kwargs['Range'] == mock_range
        assert call_args.kwargs['UseHeadingStyles'] is True
        assert call_args.kwargs['UpperHeadingLevel'] == 1
        assert call_args.kwargs['LowerHeadingLevel'] == 3
    
    def test_create_toc_with_custom_config(self):
        """测试使用自定义配置创建目录"""
        mock_range = Mock()
        mock_toc = Mock()
        
        self.mock_document.TablesOfContents.Add.return_value = mock_toc
        
        config = TocConfig(
            upper_heading_level=2,
            lower_heading_level=4,
            use_hyperlinks=False
        )
        
        result = self.manager.create_toc(mock_range, config)
        
        assert result == mock_toc
        
        # 验证调用参数
        call_args = self.mock_document.TablesOfContents.Add.call_args
        assert call_args.kwargs['UpperHeadingLevel'] == 2
        assert call_args.kwargs['LowerHeadingLevel'] == 4
        assert call_args.kwargs['UseHyperlinks'] is False
    
    def test_update_toc_success(self):
        """测试成功更新目录"""
        mock_toc = Mock()
        
        result = self.manager.update_toc(mock_toc)
        
        assert result is True
        mock_toc.Update.assert_called_once()
    
    def test_update_toc_failure(self):
        """测试更新目录失败"""
        mock_toc = Mock()
        mock_toc.Update.side_effect = Exception("Update failed")
        
        result = self.manager.update_toc(mock_toc)
        
        assert result is False
    
    def test_update_toc_page_numbers_success(self):
        """测试成功更新目录页码"""
        mock_toc = Mock()
        
        result = self.manager.update_toc_page_numbers(mock_toc)
        
        assert result is True
        mock_toc.UpdatePageNumbers.assert_called_once()
    
    def test_rebuild_toc(self):
        """测试重建目录"""
        mock_toc = Mock()
        mock_toc.Range = Mock()
        mock_new_toc = Mock()
        
        self.mock_document.TablesOfContents.Add.return_value = mock_new_toc
        
        result = self.manager.rebuild_toc(mock_toc)
        
        assert result == mock_new_toc
        mock_toc.Delete.assert_called_once()
        self.mock_document.TablesOfContents.Add.assert_called_once()
    
    def test_get_all_tocs(self):
        """测试获取所有目录"""
        mock_toc1 = Mock()
        mock_toc2 = Mock()
        
        self.mock_document.TablesOfContents = [mock_toc1, mock_toc2]
        
        tocs = self.manager.get_all_tocs()
        
        assert len(tocs) == 2
        assert mock_toc1 in tocs
        assert mock_toc2 in tocs
    
    def test_validate_toc_structure_success(self):
        """测试目录结构验证成功"""
        mock_toc = Mock()
        mock_toc.Range.Text = "目录内容"
        mock_toc.UpperHeadingLevel = 1
        mock_toc.LowerHeadingLevel = 3
        
        self.mock_document.TablesOfContents = [mock_toc]
        
        result = self.manager.validate_toc_structure()
        
        assert result.is_valid is True
        assert len(result.errors) == 0
        assert result.details["toc_count"] == 1
    
    def test_validate_toc_structure_with_errors(self):
        """测试目录结构验证有错误"""
        mock_toc = Mock()
        mock_toc.Range.Text = ""  # 空目录
        mock_toc.UpperHeadingLevel = 3
        mock_toc.LowerHeadingLevel = 1  # 级别设置错误
        
        self.mock_document.TablesOfContents = [mock_toc]
        
        result = self.manager.validate_toc_structure()
        
        assert result.is_valid is False
        assert len(result.errors) == 2  # 空目录 + 级别错误


class TestHyperlinkManager:
    """测试超链接管理器"""
    
    def setup_method(self):
        """测试前设置"""
        self.mock_word_app = Mock()
        self.mock_document = Mock()
        self.manager = HyperlinkManager(self.mock_word_app, self.mock_document)
    
    def test_detect_link_type_external(self):
        """测试检测外部链接类型"""
        assert self.manager._detect_link_type("https://example.com") == LinkType.EXTERNAL
        assert self.manager._detect_link_type("http://example.com") == LinkType.EXTERNAL
    
    def test_detect_link_type_email(self):
        """测试检测邮箱链接类型"""
        assert self.manager._detect_link_type("mailto:test@example.com") == LinkType.EMAIL
    
    def test_detect_link_type_file(self):
        """测试检测文件链接类型"""
        assert self.manager._detect_link_type("file://C:/path/to/file.txt") == LinkType.FILE
    
    def test_detect_link_type_internal(self):
        """测试检测内部链接类型"""
        assert self.manager._detect_link_type("bookmark1") == LinkType.INTERNAL
        assert self.manager._detect_link_type("") == LinkType.INTERNAL
    
    def test_create_external_hyperlink(self):
        """测试创建外部超链接"""
        mock_range = Mock()
        mock_hyperlink = Mock()
        
        self.mock_document.Hyperlinks.Add.return_value = mock_hyperlink
        
        result = self.manager.create_hyperlink(
            range_obj=mock_range,
            address="https://example.com",
            text="示例网站"
        )
        
        assert result == mock_hyperlink
        self.mock_document.Hyperlinks.Add.assert_called_once()
        
        # 验证调用参数
        call_args = self.mock_document.Hyperlinks.Add.call_args
        assert call_args.kwargs['Anchor'] == mock_range
        assert call_args.kwargs['Address'] == "https://example.com"
        assert call_args.kwargs['TextToDisplay'] == "示例网站"
    
    def test_create_internal_hyperlink(self):
        """测试创建内部超链接"""
        mock_range = Mock()
        mock_hyperlink = Mock()
        
        self.mock_document.Hyperlinks.Add.return_value = mock_hyperlink
        
        result = self.manager.create_hyperlink(
            range_obj=mock_range,
            address="bookmark1",
            text="内部链接",
            link_type=LinkType.INTERNAL
        )
        
        assert result == mock_hyperlink
        
        # 验证调用参数
        call_args = self.mock_document.Hyperlinks.Add.call_args
        assert call_args.kwargs['Address'] == ""
        assert call_args.kwargs['SubAddress'] == "bookmark1"
    
    def test_update_hyperlink_address(self):
        """测试更新超链接地址"""
        mock_hyperlink = Mock()
        
        result = self.manager.update_hyperlink(
            hyperlink=mock_hyperlink,
            new_address="https://newsite.com"
        )
        
        assert result is True
        assert mock_hyperlink.Address == "https://newsite.com"
        assert mock_hyperlink.SubAddress == ""
    
    def test_update_hyperlink_text(self):
        """测试更新超链接文本"""
        mock_hyperlink = Mock()
        
        result = self.manager.update_hyperlink(
            hyperlink=mock_hyperlink,
            new_text="新链接文本"
        )
        
        assert result is True
        assert mock_hyperlink.TextToDisplay == "新链接文本"
    
    def test_validate_external_link_valid(self):
        """测试验证有效的外部链接"""
        link = Hyperlink(
            text="示例网站",
            address="https://example.com",
            type="external",
            range_start=0,
            range_end=10
        )
        
        result = self.manager._validate_external_link(link)
        
        assert result.is_valid is True
        assert result.error_message is None
    
    def test_validate_external_link_invalid(self):
        """测试验证无效的外部链接"""
        link = Hyperlink(
            text="无效链接",
            address="not-a-valid-url",
            type="external",
            range_start=0,
            range_end=10
        )
        
        result = self.manager._validate_external_link(link)
        
        assert result.is_valid is False
        assert "URL 格式无效" in result.error_message
    
    def test_validate_email_link_valid(self):
        """测试验证有效的邮箱链接"""
        link = Hyperlink(
            text="联系邮箱",
            address="mailto:test@example.com",
            type="email",
            range_start=0,
            range_end=10
        )
        
        result = self.manager._validate_email_link(link)
        
        assert result.is_valid is True
    
    def test_validate_email_link_invalid(self):
        """测试验证无效的邮箱链接"""
        link = Hyperlink(
            text="无效邮箱",
            address="mailto:invalid-email",
            type="email",
            range_start=0,
            range_end=10
        )
        
        result = self.manager._validate_email_link(link)
        
        assert result.is_valid is False
        assert "邮箱格式无效" in result.error_message


class TestTocAndLinkFixer:
    """测试 TOC 和链接修复器主类"""
    
    def setup_method(self):
        """测试前设置"""
        self.mock_word_app = Mock()
        self.mock_document = Mock()
        self.fixer = TocAndLinkFixer(self.mock_word_app, self.mock_document)
    
    @patch.object(TocAndLinkFixer, '_fix_all_tocs')
    @patch.object(TocAndLinkFixer, '_fix_all_links')
    def test_fix_all_tocs_and_links_success(self, mock_fix_links, mock_fix_tocs):
        """测试成功修复所有目录和链接"""
        # 设置模拟返回值
        mock_fix_tocs.return_value = {
            "success": True,
            "updated_count": 2,
            "rebuilt_count": 1,
            "errors": []
        }
        
        mock_fix_links.return_value = {
            "success": True,
            "total_links": 5,
            "valid_links": 4,
            "fixed_count": 1,
            "errors": []
        }
        
        # 执行测试
        report = self.fixer.fix_all_tocs_and_links()
        
        # 验证结果
        assert report["summary"]["overall_success"] is True
        assert report["summary"]["tocs_updated"] == 2
        assert report["summary"]["tocs_rebuilt"] == 1
        assert report["summary"]["links_validated"] == 5
        assert report["summary"]["links_fixed"] == 1
    
    @patch.object(TocManager, 'get_all_tocs')
    @patch.object(TocManager, 'update_toc')
    def test_fix_all_tocs_update_success(self, mock_update, mock_get_tocs):
        """测试成功更新所有目录"""
        mock_toc1 = Mock()
        mock_toc2 = Mock()
        mock_get_tocs.return_value = [mock_toc1, mock_toc2]
        mock_update.return_value = True
        
        result = self.fixer._fix_all_tocs()
        
        assert result["success"] is True
        assert result["updated_count"] == 2
        assert result["rebuilt_count"] == 0
        assert len(result["errors"]) == 0
    
    @patch.object(TocManager, 'get_all_tocs')
    @patch.object(TocManager, 'update_toc')
    @patch.object(TocManager, 'rebuild_toc')
    def test_fix_all_tocs_with_rebuild(self, mock_rebuild, mock_update, mock_get_tocs):
        """测试更新失败时重建目录"""
        mock_toc = Mock()
        mock_get_tocs.return_value = [mock_toc]
        mock_update.return_value = False  # 更新失败
        mock_rebuild.return_value = Mock()  # 重建成功
        
        result = self.fixer._fix_all_tocs()
        
        assert result["success"] is True
        assert result["updated_count"] == 0
        assert result["rebuilt_count"] == 1
        mock_rebuild.assert_called_once_with(mock_toc)
    
    @patch.object(HyperlinkManager, 'validate_hyperlinks')
    @patch.object(HyperlinkManager, 'fix_broken_links')
    def test_fix_all_links_success(self, mock_fix_links, mock_validate):
        """测试成功修复所有链接"""
        # 模拟验证结果
        valid_result = LinkValidationResult(
            link=Mock(),
            is_valid=True
        )
        
        invalid_result = LinkValidationResult(
            link=Mock(),
            is_valid=False,
            error_message="链接损坏",
            suggested_fix="https://fixed.com"
        )
        
        mock_validate.return_value = [valid_result, invalid_result]
        mock_fix_links.return_value = 1  # 修复了1个链接
        
        result = self.fixer._fix_all_links()
        
        assert result["success"] is True
        assert result["total_links"] == 2
        assert result["valid_links"] == 1
        assert result["fixed_count"] == 1


class TestLinkValidationResult:
    """测试链接验证结果"""
    
    def test_link_validation_result_valid(self):
        """测试有效链接的验证结果"""
        link = Hyperlink(
            text="测试链接",
            address="https://example.com",
            type="external",
            range_start=0,
            range_end=10
        )
        
        result = LinkValidationResult(
            link=link,
            is_valid=True
        )
        
        assert result.is_valid is True
        assert result.error_message is None
        assert result.suggested_fix is None
    
    def test_link_validation_result_invalid(self):
        """测试无效链接的验证结果"""
        link = Hyperlink(
            text="损坏链接",
            address="broken-link",
            type="external",
            range_start=0,
            range_end=10
        )
        
        result = LinkValidationResult(
            link=link,
            is_valid=False,
            error_message="链接格式错误",
            suggested_fix="https://corrected.com"
        )
        
        assert result.is_valid is False
        assert result.error_message == "链接格式错误"
        assert result.suggested_fix == "https://corrected.com"


if __name__ == "__main__":
    pytest.main([__file__])