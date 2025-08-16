#!/usr/bin/env python3
"""
文档验证和回滚系统
提供文档状态比较、未授权变更检测和自动回滚功能
"""

import os
import json
import shutil
from typing import Dict, List, Optional, Any, Tuple
from dataclasses import dataclass, asdict
from datetime import datetime
from pathlib import Path

from .models import ValidationRule, ValidationResult, DocumentSnapshot
from .exceptions import ValidationError, DocumentError
from .doc_loader import WordSession
from .utils import safe_filename


@dataclass
class DocumentState:
    """文档状态快照"""
    headings: List[Dict[str, Any]]
    styles: List[Dict[str, Any]]
    toc_entries: List[Dict[str, Any]]
    hyperlinks: List[Dict[str, Any]]
    paragraph_count: int
    word_count: int
    page_count: int
    timestamp: str


class DocumentValidator:
    """文档验证器"""
    
    def __init__(self):
        self.validation_rules = self._load_default_rules()
    
    def _load_default_rules(self) -> List[ValidationRule]:
        """加载默认验证规则"""
        return [
            ValidationRule(
                name="标题级别检查",
                description="检查标题级别是否合理(1-6级)",
                rule_type="heading_level",
                parameters={"min_level": 1, "max_level": 6}
            ),
            ValidationRule(
                name="样式一致性检查",
                description="检查段落样式是否一致",
                rule_type="style_consistency",
                parameters={"allowed_styles": ["正文", "标题 1", "标题 2", "标题 3"]}
            ),
            ValidationRule(
                name="超链接有效性检查",
                description="检查超链接是否有效",
                rule_type="hyperlink_validity",
                parameters={"check_external": True}
            ),
            ValidationRule(
                name="目录完整性检查",
                description="检查目录是否完整和准确",
                rule_type="toc_integrity",
                parameters={"max_depth": 4}
            )
        ]
    
    def validate_document(self, document_path: str, 
                         custom_rules: Optional[List[ValidationRule]] = None) -> List[ValidationResult]:
        """验证文档"""
        rules = custom_rules or self.validation_rules
        results = []
        
        try:
            with WordSession() as word_app:
                document = word_app.Documents.Open(document_path)
                
                try:
                    document_data = self._extract_document_data(document)
                    
                    for rule in rules:
                        result = self._apply_rule(rule, document_data)
                        results.append(result)
                        
                finally:
                    document.Close(SaveChanges=False)
                    
        except Exception as e:
            results.append(ValidationResult(
                rule_name="文档访问",
                passed=False,
                message=f"无法访问文档: {str(e)}",
                details={"error": str(e)}
            ))
        
        return results
    
    def _extract_document_data(self, document) -> Dict[str, Any]:
        """提取文档数据用于验证"""
        try:
            # 提取标题
            headings = []
            for paragraph in document.Paragraphs:
                if paragraph.Style.NameLocal.startswith("标题"):
                    level_str = paragraph.Style.NameLocal.split()[-1]
                    try:
                        level = int(level_str)
                        headings.append({
                            "text": paragraph.Range.Text.strip(),
                            "level": level,
                            "style": paragraph.Style.NameLocal
                        })
                    except ValueError:
                        pass
            
            # 提取样式
            styles = []
            used_styles = set()
            for paragraph in document.Paragraphs:
                style_name = paragraph.Style.NameLocal
                if style_name not in used_styles:
                    styles.append({
                        "name": style_name,
                        "type": "paragraph"
                    })
                    used_styles.add(style_name)
            
            # 提取超链接
            hyperlinks = []
            for hyperlink in document.Hyperlinks:
                hyperlinks.append({
                    "text": hyperlink.TextToDisplay,
                    "address": hyperlink.Address or "",
                    "type": "external" if hyperlink.Address else "internal"
                })
            
            # 提取目录信息
            toc_entries = []
            try:
                if document.TablesOfContents.Count > 0:
                    toc = document.TablesOfContents(1)
                    for i in range(1, min(toc.Range.Paragraphs.Count + 1, 100)):
                        try:
                            para = toc.Range.Paragraphs(i)
                            text = para.Range.Text.strip()
                            if text and not text.isspace():
                                toc_entries.append({
                                    "text": text,
                                    "level": 1,  # 简化处理
                                    "page_number": 1
                                })
                        except:
                            break
            except:
                pass
            
            return {
                "headings": headings,
                "styles": styles,
                "hyperlinks": hyperlinks,
                "toc_entries": toc_entries,
                "paragraph_count": document.Paragraphs.Count,
                "word_count": document.Words.Count,
                "page_count": document.Range.Information(4)  # wdNumberOfPagesInDocument
            }
            
        except Exception as e:
            raise ValidationError(f"提取文档数据失败: {str(e)}")
    
    def _apply_rule(self, rule: ValidationRule, document_data: Dict[str, Any]) -> ValidationResult:
        """应用验证规则"""
        try:
            if rule.rule_type == "heading_level":
                return self._validate_heading_levels(rule, document_data)
            elif rule.rule_type == "style_consistency":
                return self._validate_style_consistency(rule, document_data)
            elif rule.rule_type == "hyperlink_validity":
                return self._validate_hyperlinks(rule, document_data)
            elif rule.rule_type == "toc_integrity":
                return self._validate_toc_integrity(rule, document_data)
            else:
                return ValidationResult(
                    rule_name=rule.name,
                    passed=False,
                    message=f"未知的验证规则类型: {rule.rule_type}",
                    details={}
                )
        except Exception as e:
            return ValidationResult(
                rule_name=rule.name,
                passed=False,
                message=f"验证规则执行失败: {str(e)}",
                details={"error": str(e)}
            )
    
    def _validate_heading_levels(self, rule: ValidationRule, document_data: Dict[str, Any]) -> ValidationResult:
        """验证标题级别"""
        headings = document_data.get("headings", [])
        min_level = rule.parameters.get("min_level", 1)
        max_level = rule.parameters.get("max_level", 6)
        
        invalid_headings = []
        for heading in headings:
            level = heading.get("level", 0)
            if level < min_level or level > max_level:
                invalid_headings.append(heading)
        
        if invalid_headings:
            return ValidationResult(
                rule_name=rule.name,
                passed=False,
                message=f"发现 {len(invalid_headings)} 个标题级别不合理",
                details={
                    "invalid_headings": invalid_headings,
                    "valid_range": f"{min_level}-{max_level}"
                }
            )
        else:
            return ValidationResult(
                rule_name=rule.name,
                passed=True,
                message=f"所有 {len(headings)} 个标题级别都合理",
                details={"heading_count": len(headings)}
            )
    
    def _validate_style_consistency(self, rule: ValidationRule, document_data: Dict[str, Any]) -> ValidationResult:
        """验证样式一致性"""
        styles = document_data.get("styles", [])
        allowed_styles = rule.parameters.get("allowed_styles", [])
        
        if not allowed_styles:
            return ValidationResult(
                rule_name=rule.name,
                passed=True,
                message="未指定允许的样式列表，跳过检查",
                details={}
            )
        
        invalid_styles = []
        for style in styles:
            if style["name"] not in allowed_styles:
                invalid_styles.append(style["name"])
        
        if invalid_styles:
            return ValidationResult(
                rule_name=rule.name,
                passed=False,
                message=f"发现 {len(invalid_styles)} 个不允许的样式",
                details={
                    "invalid_styles": invalid_styles,
                    "allowed_styles": allowed_styles
                }
            )
        else:
            return ValidationResult(
                rule_name=rule.name,
                passed=True,
                message="所有样式都在允许列表中",
                details={"style_count": len(styles)}
            )
    
    def _validate_hyperlinks(self, rule: ValidationRule, document_data: Dict[str, Any]) -> ValidationResult:
        """验证超链接有效性"""
        hyperlinks = document_data.get("hyperlinks", [])
        
        if not hyperlinks:
            return ValidationResult(
                rule_name=rule.name,
                passed=True,
                message="文档中没有超链接",
                details={}
            )
        
        invalid_links = []
        for link in hyperlinks:
            address = link.get("address", "")
            if not address or address.strip() == "":
                invalid_links.append(link)
        
        if invalid_links:
            return ValidationResult(
                rule_name=rule.name,
                passed=False,
                message=f"发现 {len(invalid_links)} 个无效超链接",
                details={
                    "invalid_links": invalid_links,
                    "total_links": len(hyperlinks)
                }
            )
        else:
            return ValidationResult(
                rule_name=rule.name,
                passed=True,
                message=f"所有 {len(hyperlinks)} 个超链接都有效",
                details={"link_count": len(hyperlinks)}
            )
    
    def _validate_toc_integrity(self, rule: ValidationRule, document_data: Dict[str, Any]) -> ValidationResult:
        """验证目录完整性"""
        toc_entries = document_data.get("toc_entries", [])
        headings = document_data.get("headings", [])
        
        if not toc_entries:
            if headings:
                return ValidationResult(
                    rule_name=rule.name,
                    passed=False,
                    message="文档有标题但没有目录",
                    details={
                        "heading_count": len(headings),
                        "toc_count": 0
                    }
                )
            else:
                return ValidationResult(
                    rule_name=rule.name,
                    passed=True,
                    message="文档没有标题和目录",
                    details={}
                )
        
        # 简化的目录完整性检查
        return ValidationResult(
            rule_name=rule.name,
            passed=True,
            message=f"目录包含 {len(toc_entries)} 个条目",
            details={
                "toc_count": len(toc_entries),
                "heading_count": len(headings)
            }
        )


class DocumentRollback:
    """文档回滚系统"""
    
    def __init__(self, backup_dir: str = "backups"):
        self.backup_dir = Path(backup_dir)
        self.backup_dir.mkdir(exist_ok=True)
    
    def create_snapshot(self, document_path: str) -> DocumentSnapshot:
        """创建文档快照"""
        try:
            # 创建备份文件
            backup_path = self._create_backup_file(document_path)
            
            # 提取文档状态
            state = self._extract_document_state(document_path)
            
            snapshot = DocumentSnapshot(
                original_path=document_path,
                backup_path=str(backup_path),
                state=state,
                timestamp=datetime.now().isoformat()
            )
            
            # 保存快照信息
            self._save_snapshot_info(snapshot)
            
            return snapshot
            
        except Exception as e:
            raise DocumentError(f"创建文档快照失败: {str(e)}")
    
    def _create_backup_file(self, document_path: str) -> Path:
        """创建备份文件"""
        original_path = Path(document_path)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_name = f"{original_path.stem}_backup_{timestamp}{original_path.suffix}"
        backup_path = self.backup_dir / backup_name
        
        shutil.copy2(document_path, backup_path)
        return backup_path
    
    def _extract_document_state(self, document_path: str) -> DocumentState:
        """提取文档状态"""
        try:
            with WordSession() as word_app:
                document = word_app.Documents.Open(document_path)
                
                try:
                    validator = DocumentValidator()
                    document_data = validator._extract_document_data(document)
                    
                    return DocumentState(
                        headings=document_data["headings"],
                        styles=document_data["styles"],
                        toc_entries=document_data["toc_entries"],
                        hyperlinks=document_data["hyperlinks"],
                        paragraph_count=document_data["paragraph_count"],
                        word_count=document_data["word_count"],
                        page_count=document_data["page_count"],
                        timestamp=datetime.now().isoformat()
                    )
                    
                finally:
                    document.Close(SaveChanges=False)
                    
        except Exception as e:
            raise DocumentError(f"提取文档状态失败: {str(e)}")
    
    def _save_snapshot_info(self, snapshot: DocumentSnapshot):
        """保存快照信息"""
        snapshot_file = self.backup_dir / f"snapshot_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        
        with open(snapshot_file, 'w', encoding='utf-8') as f:
            json.dump(asdict(snapshot), f, ensure_ascii=False, indent=2)
    
    def detect_unauthorized_changes(self, before_snapshot: DocumentSnapshot, 
                                  after_document_path: str,
                                  authorized_comment_ids: List[str]) -> List[Dict[str, Any]]:
        """检测未授权的变更"""
        try:
            after_state = self._extract_document_state(after_document_path)
            changes = []
            
            # 检查标题变更
            before_headings = {h["text"]: h for h in before_snapshot.state.headings}
            after_headings = {h["text"]: h for h in after_state.headings}
            
            for text, after_heading in after_headings.items():
                before_heading = before_headings.get(text)
                if before_heading and before_heading["level"] != after_heading["level"]:
                    changes.append({
                        "type": "heading_level_change",
                        "text": text,
                        "before_level": before_heading["level"],
                        "after_level": after_heading["level"],
                        "authorized": False  # 需要进一步检查授权
                    })
            
            # 检查样式变更
            before_styles = set(s["name"] for s in before_snapshot.state.styles)
            after_styles = set(s["name"] for s in after_state.styles)
            
            new_styles = after_styles - before_styles
            removed_styles = before_styles - after_styles
            
            if new_styles:
                changes.append({
                    "type": "styles_added",
                    "styles": list(new_styles),
                    "authorized": False
                })
            
            if removed_styles:
                changes.append({
                    "type": "styles_removed", 
                    "styles": list(removed_styles),
                    "authorized": False
                })
            
            return changes
            
        except Exception as e:
            raise ValidationError(f"检测变更失败: {str(e)}")
    
    def rollback_document(self, snapshot: DocumentSnapshot) -> bool:
        """回滚文档到快照状态"""
        try:
            if not os.path.exists(snapshot.backup_path):
                raise DocumentError(f"备份文件不存在: {snapshot.backup_path}")
            
            # 复制备份文件覆盖原文件
            shutil.copy2(snapshot.backup_path, snapshot.original_path)
            
            return True
            
        except Exception as e:
            raise DocumentError(f"文档回滚失败: {str(e)}")
    
    def cleanup_old_backups(self, days: int = 7):
        """清理旧备份文件"""
        try:
            cutoff_time = datetime.now().timestamp() - (days * 24 * 3600)
            
            for backup_file in self.backup_dir.glob("*_backup_*"):
                if backup_file.stat().st_mtime < cutoff_time:
                    backup_file.unlink()
                    
        except Exception as e:
            print(f"清理备份文件时出错: {str(e)}")


class ValidationSystem:
    """完整的验证系统"""
    
    def __init__(self, backup_dir: str = "backups"):
        self.validator = DocumentValidator()
        self.rollback = DocumentRollback(backup_dir)
    
    def validate_and_protect(self, document_path: str, 
                           authorized_comment_ids: List[str],
                           custom_rules: Optional[List[ValidationRule]] = None) -> Tuple[List[ValidationResult], Optional[DocumentSnapshot]]:
        """验证文档并创建保护快照"""
        # 创建快照
        snapshot = self.rollback.create_snapshot(document_path)
        
        # 验证文档
        validation_results = self.validator.validate_document(document_path, custom_rules)
        
        return validation_results, snapshot
    
    def check_and_rollback_if_needed(self, snapshot: DocumentSnapshot,
                                   document_path: str,
                                   authorized_comment_ids: List[str]) -> Tuple[bool, List[Dict[str, Any]]]:
        """检查变更并在需要时回滚"""
        # 检测未授权变更
        changes = self.rollback.detect_unauthorized_changes(
            snapshot, document_path, authorized_comment_ids
        )
        
        # 过滤出真正未授权的变更
        unauthorized_changes = [c for c in changes if not c.get("authorized", True)]
        
        if unauthorized_changes:
            # 执行回滚
            success = self.rollback.rollback_document(snapshot)
            return success, unauthorized_changes
        
        return True, []


if __name__ == "__main__":
    # 测试验证系统
    system = ValidationSystem()
    
    # 示例：验证文档
    try:
        results, snapshot = system.validate_and_protect("test_document.docx", ["comment_1"])
        
        for result in results:
            status = "✅" if result.passed else "❌"
            print(f"{status} {result.rule_name}: {result.message}")
        
        print(f"\n📸 快照已创建: {snapshot.backup_path}")
        
    except Exception as e:
        print(f"验证失败: {str(e)}")