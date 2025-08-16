"""
AutoWord Constants
常量定义
"""

# 格式化任务类型 - 需要批注授权的任务类型
FORMAT_TYPES = {
    "set_paragraph_style",
    "set_heading_level", 
    "apply_template",
    "replace_hyperlink",
    "rebuild_toc",
    "update_toc_levels"
}

# 允许无批注的任务类型
ALLOWED_WITHOUT_COMMENT = {
    "rewrite",
    "insert", 
    "delete",
    "refresh_toc_numbers"
}

# LLM 系统提示词
SYSTEM_PROMPT = """You are a Word editing planner.
- You will receive the complete context in one turn. Do not request chunking.
- Do not change formatting unless explicitly requested by a comment.
- Formatting includes styles, heading levels, template/theme, figure/table styles, TOC structure, hyperlink text/targets.
- Output only valid JSON matching the schema; no prose.
- TOC/hyperlink updates must be at the end. Allowed without comment: refresh_toc_numbers only."""

# 用户提示词模板
USER_PROMPT_TEMPLATE = """Document Structure Summary:
{snapshot}

Comments (id/author/page/anchor_excerpt/text):
{comments}

Output Schema (JSON):
{schema_json}

Constraints:
- Unless a comment explicitly asks for it, DO NOT change formatting (see types).
- Allowed without comment: rewrite/insert/delete, refresh_toc_numbers.
- Use bookmarks/anchors in 'locator'. Keep citations unchanged."""

# API 配置
API_BASE = "https://globalai.vip/v1/chat/completions"
API_TIMEOUT = 120

# 文件扩展名
SUPPORTED_EXTENSIONS = {".docx"}

# 日志配置
LOG_FORMAT = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
LOG_DATE_FORMAT = "%Y-%m-%d %H:%M:%S"

# 备份配置
BACKUP_SUFFIX = "_backup"
BACKUP_DATE_FORMAT = "%Y%m%d_%H%M%S"

# TOC 配置
TOC_MAX_LEVELS = 3
TOC_MIN_LEVEL = 1

# 错误消息
ERROR_MESSAGES = {
    "missing_api_key": "缺少 API 密钥，请设置环境变量",
    "invalid_document": "无效的文档文件",
    "com_error": "Word COM 操作错误",
    "llm_error": "LLM 服务错误",
    "validation_error": "数据验证错误",
    "format_protection": "格式保护：未授权的格式变更被阻止"
}