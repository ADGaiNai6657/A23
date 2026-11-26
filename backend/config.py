import os
from pathlib import Path

from dotenv import load_dotenv

BACKEND = Path(__file__).resolve().parent
ROOT = BACKEND.parent
load_dotenv(BACKEND / ".env")

DATA = Path(os.environ.get("A23_DATA_DIR", BACKEND / "data")).resolve()
TRAINING_INPUTS = ROOT / "doc"
ARCHIVE_TEMPLATES = ROOT / "dataset" / "测试集" / "包含模板文件"
DEFAULT_INPUTS = ARCHIVE_TEMPLATES if ARCHIVE_TEMPLATES.is_dir() and not TRAINING_INPUTS.is_dir() else TRAINING_INPUTS
INPUTS = Path(os.environ.get("A23_DOC_DIR", DEFAULT_INPUTS)).resolve()
DIST = ROOT / "frontend" / "dist"


def api_key() -> str:
    key = os.getenv("DEEPSEEK_API_KEY", "").strip()
    if not key:
        raise ValueError("请在 backend/.env 中配置 DEEPSEEK_API_KEY")
    return key
