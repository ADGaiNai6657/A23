import os
from pathlib import Path

from dotenv import load_dotenv

BACKEND = Path(__file__).resolve().parent
ROOT = BACKEND.parent
load_dotenv(BACKEND / ".env")

DATA = Path(os.environ.get("A23_DATA_DIR", BACKEND / "data")).resolve()
INPUTS = Path(os.environ.get("A23_DOC_DIR", ROOT / "doc")).resolve()
DIST = ROOT / "frontend" / "dist"


def api_key() -> str:
    key = os.getenv("DEEPSEEK_API_KEY", "").strip()
    if not key:
        raise ValueError("请在 backend/.env 中配置 DEEPSEEK_API_KEY")
    return key
