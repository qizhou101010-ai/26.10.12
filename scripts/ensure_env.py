"""第一次使用时，在 skill 目录下建一个独立的 Python 环境并安装依赖；之后直接复用。

用法：python ensure_env.py      → 最后一行打印该环境的 python 路径，后续脚本都用它运行
需要 Python 3.10 及以上。
"""
import os
import subprocess
import sys
import venv
from pathlib import Path

SKILL_DIR = Path(__file__).resolve().parent.parent
VENV = SKILL_DIR / ".venv"
PY = VENV / ("Scripts/python.exe" if os.name == "nt" else "bin/python")
REQUIREMENTS = SKILL_DIR / "scripts" / "requirements.txt"
CHECK = "import openpyxl, docx, pypdf, jieba, rank_bm25, dotenv, openai"


def ready() -> bool:
    return PY.exists() and subprocess.run([str(PY), "-c", CHECK], capture_output=True).returncode == 0


def main():
    if sys.version_info < (3, 10):
        raise SystemExit(f"需要 Python 3.10 及以上，当前是 {sys.version.split()[0]}")
    if not ready():
        if not PY.exists():
            print(f"正在创建 Python 环境：{VENV}", file=sys.stderr)
            venv.create(VENV, with_pip=True)
        print("正在安装依赖（只需一次）...", file=sys.stderr)
        subprocess.check_call([str(PY), "-m", "pip", "install", "-q", "--prefer-binary", "-r", str(REQUIREMENTS)])
        if not ready():
            raise SystemExit("依赖安装后仍无法导入，请检查网络或 pip 源")
    print(PY)


if __name__ == "__main__":
    main()
