"""
LLM数据收集器配置文件
"""

# 服务器配置
SERVER_HOST = "0.0.0.0"
SERVER_PORT = 5000
DEBUG_MODE = False

# 数据存储配置
AUTO_SAVE = False
AUTO_SAVE_PATH = "data/collected_user_data.json"

# 日志配置
LOG_LEVEL = "INFO"
LOG_FORMAT = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"

# API配置
API_VERSION = "v1"
MAX_CONTENT_LENGTH = 16 * 1024 * 1024  # 16MB
