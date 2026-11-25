from pathlib import Path

from parser_docx import parse_docx
from parser_text import parse_text
from parser_xlsx import parse_xlsx
from schema import ParsedDoc

PARSERS = {".txt": parse_text, ".md": parse_text, ".docx": parse_docx, ".xlsx": parse_xlsx}


def leaf_name(name: str) -> str:
    if not name or name in {".", ".."} or any(char in name for char in ("/", "\\", "\x00", ":")):
        raise ValueError("请输入文件名，不能包含目录")
    return name


def parse_document(data: bytes, name: str) -> ParsedDoc:
    name = leaf_name(name)
    extension = Path(name).suffix.lower()
    if extension not in PARSERS:
        raise ValueError(f"不支持文件格式 {extension or '（无后缀）'}")
    try:
        blocks = PARSERS[extension](data)
    except (OSError, ValueError, KeyError) as error:
        raise ValueError(f"无法读取 {extension} 文件，请检查内容和格式") from error
    except Exception as error:
        # OOXML readers also raise package and XML exceptions.
        raise ValueError(f"无法解析 {extension} 文件，请检查文件是否损坏") from error
    return ParsedDoc(file=name, blocks=blocks)
