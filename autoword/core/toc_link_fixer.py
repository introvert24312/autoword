"""
AutoWord TOC and Link Fixer
TOC 和超链接管理器，负责目录和链接的创建、更新和修复
"""

import logging
from typing import List, Dict, Any, Optional, Tuple
from dataclasses import dataclass
from enum import Enum

import win32com.client as win32
from win32com.client import constants as win32_constants

from .models import TocEntry, Hyperlink, Reference, ValidationResult
from .exceptions import COMError, TaskExecutionError
from .utils import truncate_text


logger = logging.getLogger(__name__)


class TocStyle(str, Enum):
    """目录样式枚举"""
    CLASSIC = "classic"          # 经典样式
    MODERN = "modern"            # 现代样式
    SIMPLE = "simple"            # 简洁样式
    FORMAL = "formal"            # 正式样式


class LinkType(str, Enum):
    """链接类型枚举"""
    INTERNAL = "internal"        # 内部链接（书签、标题）
    EXTERNAL = "external"        # 外部链接（网址）
    EMAIL = "email"              # 邮箱链接
    FILE = "file"                # 文件链接


@dataclass
class TocConfig:
    """目录配置"""
    use_heading_styles: bool = True
    upper_heading_level: int = 1
    lower_heading_level: int = 3
    use_hyperlinks: bool = True
    right_align_page_numbers: bool = True
    include_page_numbers: bool = True
    use_tab_leader: bool = True
    tab_leader: str = "."
    style: TocStyle = TocStyle.CLASSIC


@dataclass
class LinkValidationResult:
    """链接验证结果"""
    link: Hyperlink
    is_valid: bool
    error_message: Optional[str] = None
    suggested_fix: Optional[str] = None


class TocManager:
    """目录管理器"""
    
    def __init__(self, word_app: Any, document: Any):
        """
        初始化目录管理器
        
        Args:
            word_app: Word 应用程序对象
            document: Word 文档对象
        """
        self.word_app = word_app
        self.document = document
    
    def create_toc(self, 
                   range_obj: Any,
                   config: TocConfig = None) -> Any:
        """
        创建目录
        
        Args:
            range_obj: 插入目录的位置
            config: 目录配置
            
        Returns:
            创建的目录对象
        """
        if config is None:
            config = TocConfig()
        
        try:
            logger.info(f"创建目录: 级别 {config.upper_heading_level}-{config.lower_heading_level}")
            
            # 创建目录
            toc = self.document.TablesOfContents.Add(
                Range=range_obj,
                UseHeadingStyles=config.use_heading_styles,
                UpperHeadingLevel=config.upper_heading_level,
                LowerHeadingLevel=config.lower_heading_level,
                UseFields=False,
                TableID=None,
                RightAlignPageNumbers=config.right_align_page_numbers,
                IncludePageNumbers=config.include_page_numbers,
                AddedStyles=None,
                UseHyperlinks=config.use_hyperlinks,
                HidePageNumbersInWeb=True
            )
            
            # 应用样式
            self._apply_toc_style(toc, config.style)
            
            # 设置制表符引导符
            if config.use_tab_leader:
                self._set_tab_leader(toc, config.tab_leader)
            
            logger.info("目录创建成功")
            return toc
            
        except Exception as e:
            raise TaskExecutionError(f"创建目录失败: {e}")
    
    def update_toc(self, toc: Any) -> bool:
        """
        更新目录
        
        Args:
            toc: 目录对象
            
        Returns:
            是否更新成功
        """
        try:
            logger.info("更新目录内容")
            toc.Update()
            logger.info("目录更新成功")
            return True
        except Exception as e:
            logger.error(f"更新目录失败: {e}")
            return False
    
    def update_toc_page_numbers(self, toc: Any) -> bool:
        """
        更新目录页码
        
        Args:
            toc: 目录对象
            
        Returns:
            是否更新成功
        """
        try:
            logger.info("更新目录页码")
            toc.UpdatePageNumbers()
            logger.info("目录页码更新成功")
            return True
        except Exception as e:
            logger.error(f"更新目录页码失败: {e}")
            return False
    
    def rebuild_toc(self, 
                    toc: Any,
                    config: TocConfig = None) -> Any:
        """
        重建目录
        
        Args:
            toc: 现有目录对象
            config: 新的目录配置
            
        Returns:
            重建的目录对象
        """
        try:
            logger.info("重建目录")
            
            # 获取目录位置
            toc_range = toc.Range
            
            # 删除现有目录
            toc.Delete()
            
            # 创建新目录
            new_toc = self.create_toc(toc_range, config)
            
            logger.info("目录重建成功")
            return new_toc
            
        except Exception as e:
            raise TaskExecutionError(f"重建目录失败: {e}")
    
    def get_all_tocs(self) -> List[Any]:
        """
        获取文档中的所有目录
        
        Returns:
            目录对象列表
        """
        try:
            tocs = []
            for toc in self.document.TablesOfContents:
                tocs.append(toc)
            return tocs
        except Exception as e:
            logger.error(f"获取目录列表失败: {e}")
            return []
    
    def validate_toc_structure(self) -> ValidationResult:
        """
        验证目录结构
        
        Returns:
            验证结果
        """
        errors = []
        warnings = []
        
        try:
            tocs = self.get_all_tocs()
            
            if not tocs:
                warnings.append("文档中没有目录")
            
            for i, toc in enumerate(tocs, 1):
                try:
                    # 检查目录是否为空
                    if not toc.Range.Text.strip():
                        errors.append(f"目录 {i} 为空")
                    
                    # 检查级别设置
                    if toc.UpperHeadingLevel > toc.LowerHeadingLevel:
                        errors.append(f"目录 {i} 级别设置错误: 上级 {toc.UpperHeadingLevel} > 下级 {toc.LowerHeadingLevel}")
                    
                    # 检查级别范围
                    if toc.UpperHeadingLevel < 1 or toc.LowerHeadingLevel > 9:
                        warnings.append(f"目录 {i} 级别范围异常: {toc.UpperHeadingLevel}-{toc.LowerHeadingLevel}")
                    
                except Exception as e:
                    errors.append(f"检查目录 {i} 时出错: {e}")
            
            return ValidationResult(
                is_valid=len(errors) == 0,
                errors=errors,
                warnings=warnings,
                details={"toc_count": len(tocs)}
            )
            
        except Exception as e:
            return ValidationResult(
                is_valid=False,
                errors=[f"目录结构验证失败: {e}"],
                warnings=[],
                details={}
            )
    
    def _apply_toc_style(self, toc: Any, style: TocStyle):
        """应用目录样式"""
        try:
            # 根据样式类型设置不同的格式
            if style == TocStyle.MODERN:
                # 现代样式：无引导符，左对齐页码
                toc.RightAlignPageNumbers = False
            elif style == TocStyle.SIMPLE:
                # 简洁样式：无页码
                toc.IncludePageNumbers = False
            elif style == TocStyle.FORMAL:
                # 正式样式：使用超链接，右对齐页码
                toc.UseHyperlinks = True
                toc.RightAlignPageNumbers = True
            # CLASSIC 样式使用默认设置
            
        except Exception as e:
            logger.warning(f"应用目录样式失败: {e}")
    
    def _set_tab_leader(self, toc: Any, leader: str):
        """设置制表符引导符"""
        try:
            # 这里可以设置制表符引导符的样式
            # Word COM 中制表符引导符的设置比较复杂，简化实现
            logger.debug(f"设置制表符引导符: {leader}")
        except Exception as e:
            logger.warning(f"设置制表符引导符失败: {e}")


class HyperlinkManager:
    """超链接管理器"""
    
    def __init__(self, word_app: Any, document: Any):
        """
        初始化超链接管理器
        
        Args:
            word_app: Word 应用程序对象
            document: Word 文档对象
        """
        self.word_app = word_app
        self.document = document
    
    def create_hyperlink(self,
                        range_obj: Any,
                        address: str,
                        text: str = None,
                        link_type: LinkType = None) -> Any:
        """
        创建超链接
        
        Args:
            range_obj: 链接位置
            address: 链接地址
            text: 链接文本（可选）
            link_type: 链接类型（可选，自动检测）
            
        Returns:
            创建的超链接对象
        """
        try:
            if link_type is None:
                link_type = self._detect_link_type(address)
            
            logger.info(f"创建超链接: {link_type.value} - {address}")
            
            # 根据链接类型设置参数
            if link_type == LinkType.INTERNAL:
                # 内部链接使用 SubAddress
                hyperlink = self.document.Hyperlinks.Add(
                    Anchor=range_obj,
                    Address="",
                    SubAddress=address,
                    TextToDisplay=text
                )
            else:
                # 外部链接使用 Address
                hyperlink = self.document.Hyperlinks.Add(
                    Anchor=range_obj,
                    Address=address,
                    SubAddress="",
                    TextToDisplay=text
                )
            
            logger.info("超链接创建成功")
            return hyperlink
            
        except Exception as e:
            raise TaskExecutionError(f"创建超链接失败: {e}")
    
    def update_hyperlink(self,
                        hyperlink: Any,
                        new_address: str = None,
                        new_text: str = None) -> bool:
        """
        更新超链接
        
        Args:
            hyperlink: 超链接对象
            new_address: 新地址（可选）
            new_text: 新文本（可选）
            
        Returns:
            是否更新成功
        """
        try:
            logger.info("更新超链接")
            
            if new_address:
                link_type = self._detect_link_type(new_address)
                if link_type == LinkType.INTERNAL:
                    hyperlink.Address = ""
                    hyperlink.SubAddress = new_address
                else:
                    hyperlink.Address = new_address
                    hyperlink.SubAddress = ""
            
            if new_text:
                hyperlink.TextToDisplay = new_text
            
            logger.info("超链接更新成功")
            return True
            
        except Exception as e:
            logger.error(f"更新超链接失败: {e}")
            return False
    
    def validate_hyperlinks(self) -> List[LinkValidationResult]:
        """
        验证所有超链接
        
        Returns:
            链接验证结果列表
        """
        results = []
        
        try:
            logger.info("开始验证超链接")
            
            for hyperlink in self.document.Hyperlinks:
                try:
                    # 提取链接信息
                    text = hyperlink.TextToDisplay or ""
                    address = hyperlink.Address or ""
                    sub_address = hyperlink.SubAddress or ""
                    
                    # 创建 Hyperlink 模型对象
                    link_obj = Hyperlink(
                        text=text,
                        address=address or sub_address,
                        type=self._detect_link_type(address or sub_address).value,
                        range_start=hyperlink.Range.Start,
                        range_end=hyperlink.Range.End
                    )
                    
                    # 验证链接
                    validation_result = self._validate_single_link(link_obj, hyperlink)
                    results.append(validation_result)
                    
                except Exception as e:
                    # 单个链接验证失败不影响其他链接
                    error_result = LinkValidationResult(
                        link=Hyperlink(
                            text="未知链接",
                            address="",
                            type="unknown",
                            range_start=0,
                            range_end=0
                        ),
                        is_valid=False,
                        error_message=f"验证链接时出错: {e}"
                    )
                    results.append(error_result)
            
            valid_count = sum(1 for r in results if r.is_valid)
            logger.info(f"超链接验证完成: {valid_count}/{len(results)} 个有效")
            
        except Exception as e:
            logger.error(f"验证超链接失败: {e}")
        
        return results
    
    def fix_broken_links(self, validation_results: List[LinkValidationResult]) -> int:
        """
        修复损坏的链接
        
        Args:
            validation_results: 验证结果列表
            
        Returns:
            修复的链接数量
        """
        fixed_count = 0
        
        try:
            logger.info("开始修复损坏的链接")
            
            for result in validation_results:
                if not result.is_valid and result.suggested_fix:
                    try:
                        # 查找对应的超链接对象
                        hyperlink = self._find_hyperlink_by_range(
                            result.link.range_start,
                            result.link.range_end
                        )
                        
                        if hyperlink:
                            # 应用建议的修复
                            if self.update_hyperlink(hyperlink, new_address=result.suggested_fix):
                                fixed_count += 1
                                logger.info(f"修复链接: {result.link.address} -> {result.suggested_fix}")
                    
                    except Exception as e:
                        logger.error(f"修复链接失败: {e}")
            
            logger.info(f"链接修复完成: {fixed_count} 个")
            
        except Exception as e:
            logger.error(f"修复链接过程失败: {e}")
        
        return fixed_count
    
    def get_all_hyperlinks(self) -> List[Hyperlink]:
        """
        获取所有超链接
        
        Returns:
            超链接列表
        """
        hyperlinks = []
        
        try:
            for hyperlink in self.document.Hyperlinks:
                try:
                    text = hyperlink.TextToDisplay or ""
                    address = hyperlink.Address or ""
                    sub_address = hyperlink.SubAddress or ""
                    
                    link_obj = Hyperlink(
                        text=text,
                        address=address or sub_address,
                        type=self._detect_link_type(address or sub_address).value,
                        range_start=hyperlink.Range.Start,
                        range_end=hyperlink.Range.End
                    )
                    
                    hyperlinks.append(link_obj)
                    
                except Exception as e:
                    logger.warning(f"读取超链接失败: {e}")
                    continue
        
        except Exception as e:
            logger.error(f"获取超链接列表失败: {e}")
        
        return hyperlinks
    
    def _detect_link_type(self, address: str) -> LinkType:
        """检测链接类型"""
        if not address:
            return LinkType.INTERNAL
        
        address_lower = address.lower()
        
        if address_lower.startswith(('http://', 'https://')):
            return LinkType.EXTERNAL
        elif address_lower.startswith('mailto:'):
            return LinkType.EMAIL
        elif address_lower.startswith('file://'):
            return LinkType.FILE
        else:
            return LinkType.INTERNAL
    
    def _validate_single_link(self, link: Hyperlink, hyperlink_obj: Any) -> LinkValidationResult:
        """验证单个链接"""
        try:
            link_type = LinkType(link.type)
            
            if link_type == LinkType.INTERNAL:
                return self._validate_internal_link(link, hyperlink_obj)
            elif link_type == LinkType.EXTERNAL:
                return self._validate_external_link(link)
            elif link_type == LinkType.EMAIL:
                return self._validate_email_link(link)
            elif link_type == LinkType.FILE:
                return self._validate_file_link(link)
            else:
                return LinkValidationResult(
                    link=link,
                    is_valid=False,
                    error_message="未知链接类型"
                )
        
        except Exception as e:
            return LinkValidationResult(
                link=link,
                is_valid=False,
                error_message=f"验证链接时出错: {e}"
            )
    
    def _validate_internal_link(self, link: Hyperlink, hyperlink_obj: Any) -> LinkValidationResult:
        """验证内部链接"""
        try:
            # 检查书签是否存在
            bookmark_name = link.address
            
            if bookmark_name:
                # 检查书签是否存在
                bookmark_exists = False
                try:
                    for bookmark in self.document.Bookmarks:
                        if bookmark.Name == bookmark_name:
                            bookmark_exists = True
                            break
                except:
                    pass
                
                if not bookmark_exists:
                    return LinkValidationResult(
                        link=link,
                        is_valid=False,
                        error_message=f"书签 '{bookmark_name}' 不存在",
                        suggested_fix=self._suggest_bookmark_fix(bookmark_name)
                    )
            
            return LinkValidationResult(link=link, is_valid=True)
            
        except Exception as e:
            return LinkValidationResult(
                link=link,
                is_valid=False,
                error_message=f"验证内部链接失败: {e}"
            )
    
    def _validate_external_link(self, link: Hyperlink) -> LinkValidationResult:
        """验证外部链接"""
        try:
            import re
            
            # 简单的 URL 格式验证
            url_pattern = r'^https?://[^\s/$.?#].[^\s]*$'
            
            if re.match(url_pattern, link.address):
                return LinkValidationResult(link=link, is_valid=True)
            else:
                return LinkValidationResult(
                    link=link,
                    is_valid=False,
                    error_message="URL 格式无效"
                )
        
        except Exception as e:
            return LinkValidationResult(
                link=link,
                is_valid=False,
                error_message=f"验证外部链接失败: {e}"
            )
    
    def _validate_email_link(self, link: Hyperlink) -> LinkValidationResult:
        """验证邮箱链接"""
        try:
            import re
            
            # 提取邮箱地址
            email = link.address.replace('mailto:', '')
            
            # 简单的邮箱格式验证
            email_pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
            
            if re.match(email_pattern, email):
                return LinkValidationResult(link=link, is_valid=True)
            else:
                return LinkValidationResult(
                    link=link,
                    is_valid=False,
                    error_message="邮箱格式无效"
                )
        
        except Exception as e:
            return LinkValidationResult(
                link=link,
                is_valid=False,
                error_message=f"验证邮箱链接失败: {e}"
            )
    
    def _validate_file_link(self, link: Hyperlink) -> LinkValidationResult:
        """验证文件链接"""
        try:
            import os
            
            # 提取文件路径
            file_path = link.address.replace('file://', '')
            
            if os.path.exists(file_path):
                return LinkValidationResult(link=link, is_valid=True)
            else:
                return LinkValidationResult(
                    link=link,
                    is_valid=False,
                    error_message=f"文件不存在: {file_path}"
                )
        
        except Exception as e:
            return LinkValidationResult(
                link=link,
                is_valid=False,
                error_message=f"验证文件链接失败: {e}"
            )
    
    def _suggest_bookmark_fix(self, bookmark_name: str) -> Optional[str]:
        """建议书签修复方案"""
        try:
            # 查找相似的书签名称
            existing_bookmarks = []
            for bookmark in self.document.Bookmarks:
                existing_bookmarks.append(bookmark.Name)
            
            # 简单的相似度匹配
            for existing in existing_bookmarks:
                if bookmark_name.lower() in existing.lower() or existing.lower() in bookmark_name.lower():
                    return existing
            
            return None
            
        except Exception:
            return None
    
    def _find_hyperlink_by_range(self, start: int, end: int) -> Optional[Any]:
        """根据范围查找超链接对象"""
        try:
            for hyperlink in self.document.Hyperlinks:
                if hyperlink.Range.Start == start and hyperlink.Range.End == end:
                    return hyperlink
            return None
        except Exception:
            return None


class TocAndLinkFixer:
    """TOC 和超链接修复器主类"""
    
    def __init__(self, word_app: Any, document: Any):
        """
        初始化修复器
        
        Args:
            word_app: Word 应用程序对象
            document: Word 文档对象
        """
        self.word_app = word_app
        self.document = document
        self.toc_manager = TocManager(word_app, document)
        self.link_manager = HyperlinkManager(word_app, document)
    
    def fix_all_tocs_and_links(self) -> Dict[str, Any]:
        """
        修复所有目录和链接
        
        Returns:
            修复结果报告
        """
        report = {
            "toc_results": {},
            "link_results": {},
            "summary": {}
        }
        
        try:
            logger.info("开始修复所有目录和链接")
            
            # 1. 修复目录
            toc_results = self._fix_all_tocs()
            report["toc_results"] = toc_results
            
            # 2. 修复链接
            link_results = self._fix_all_links()
            report["link_results"] = link_results
            
            # 3. 生成摘要
            report["summary"] = {
                "tocs_updated": toc_results.get("updated_count", 0),
                "tocs_rebuilt": toc_results.get("rebuilt_count", 0),
                "links_validated": link_results.get("total_links", 0),
                "links_fixed": link_results.get("fixed_count", 0),
                "overall_success": toc_results.get("success", False) and link_results.get("success", False)
            }
            
            logger.info("目录和链接修复完成")
            
        except Exception as e:
            logger.error(f"修复过程失败: {e}")
            report["summary"]["error"] = str(e)
        
        return report
    
    def _fix_all_tocs(self) -> Dict[str, Any]:
        """修复所有目录"""
        results = {
            "success": False,
            "updated_count": 0,
            "rebuilt_count": 0,
            "errors": []
        }
        
        try:
            tocs = self.toc_manager.get_all_tocs()
            
            for i, toc in enumerate(tocs, 1):
                try:
                    # 尝试更新目录
                    if self.toc_manager.update_toc(toc):
                        results["updated_count"] += 1
                        logger.info(f"目录 {i} 更新成功")
                    else:
                        # 更新失败，尝试重建
                        logger.warning(f"目录 {i} 更新失败，尝试重建")
                        self.toc_manager.rebuild_toc(toc)
                        results["rebuilt_count"] += 1
                        logger.info(f"目录 {i} 重建成功")
                
                except Exception as e:
                    error_msg = f"处理目录 {i} 失败: {e}"
                    results["errors"].append(error_msg)
                    logger.error(error_msg)
            
            results["success"] = len(results["errors"]) == 0
            
        except Exception as e:
            results["errors"].append(f"修复目录过程失败: {e}")
        
        return results
    
    def _fix_all_links(self) -> Dict[str, Any]:
        """修复所有链接"""
        results = {
            "success": False,
            "total_links": 0,
            "valid_links": 0,
            "fixed_count": 0,
            "errors": []
        }
        
        try:
            # 验证所有链接
            validation_results = self.link_manager.validate_hyperlinks()
            results["total_links"] = len(validation_results)
            
            # 统计有效链接
            valid_results = [r for r in validation_results if r.is_valid]
            results["valid_links"] = len(valid_results)
            
            # 修复损坏的链接
            results["fixed_count"] = self.link_manager.fix_broken_links(validation_results)
            
            # 记录错误
            for result in validation_results:
                if not result.is_valid and not result.suggested_fix:
                    results["errors"].append(result.error_message)
            
            results["success"] = True
            
        except Exception as e:
            results["errors"].append(f"修复链接过程失败: {e}")
        
        return results