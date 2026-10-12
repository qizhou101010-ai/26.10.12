"""共享配置：目录位置和大模型设置（从 .env 读取）。"""
import os
from pathlib import Path

from dotenv import load_dotenv

CODE_DIR = Path(__file__).resolve().parent
# 工作目录 = 存放资料库、模板、待审核、输出的地方：环境变量 AUDIT_WORKSPACE，否则为当前目录
ROOT = Path(os.getenv("AUDIT_WORKSPACE") or os.getcwd()).resolve()
load_dotenv(ROOT / ".env")
load_dotenv(CODE_DIR / ".env")    # 工作目录没有 .env 时，用代码目录里的（已设置的不会被覆盖）

KB_DIR = ROOT / "资料库"        # 自建资料库：把资料文件放这里，或用 kb.py 导入
TEMPLATE_DIR = ROOT / "模板"    # 空白填报模板
SAMPLE_DIR = ROOT / "样例"      # 填写正确的样例，用于自动总结审核规则
INPUT_DIR = ROOT / "待审核"     # 需要审核的 Excel，可放多个
OUTPUT_DIR = ROOT / "输出"      # 错误报告和修正版 Excel
RULES_FILE = ROOT / "审核规则.xlsx"

LLM_API_KEY = os.getenv("DEEPSEEK_API_KEY", "")
LLM_BASE_URL = os.getenv("LLM_BASE_URL", "https://api.deepseek.com")
LLM_MODEL = os.getenv("LLM_MODEL", "deepseek-chat")
MAX_WORKERS = int(os.getenv("MAX_WORKERS", "4"))
# 大模型使用方式：auto = 程序拿不准时才调用（没填 key 时自动不用）；off = 完全不用，拿不准的交人工
USE_LLM = os.getenv("USE_LLM", "auto")
# 资料全文不超过这么多字时，整篇交给模型读，不做段落检索（检索可能漏掉关键句）
FULL_TEXT_CHARS = int(os.getenv("FULL_TEXT_CHARS", "6000"))
