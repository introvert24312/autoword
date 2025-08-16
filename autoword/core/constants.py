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
SYSTEM_PROMPT = """你是一个Word文档自动化助手。根据用户提供的批注，生成对应的任务列表。

重要规则：
- 不要更改格式，除非批注明确要求
- 格式包括样式、标题级别、模板/主题等
- 只输出有效的JSON，不要添加说明文字
- 严格按照提供的JSON Schema格式返回

支持的任务类型：
- rewrite: 重写内容（无需批注授权）
- insert: 插入内容（无需批注授权）
- delete: 删除内容（无需批注授权）
- set_heading_level: 设置标题级别（需要批注授权）
- set_paragraph_style: 设置段落样式（需要批注授权）

定位方式：
- find: 通过文本查找
- heading: 通过标题查找
- bookmark: 通过书签查找"""

# 用户提示词模板
USER_PROMPT_TEMPLATE = """文档结构摘要：
{snapshot}

批注列表：
{comments}

请按照以下JSON格式返回任务列表：
{schema_json}

注意：
- 除非批注明确要求，否则不要更改格式
- 每个任务必须有对应的批注ID（source_comment_id）
- 使用准确的定位信息"""

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