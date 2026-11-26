from io import BytesIO
from posixpath import normpath
from xml.etree import ElementTree
from zipfile import ZIP_DEFLATED, ZipFile

from docx import Document
from docx.document import Document as WordDocument
from docx.table import Table

from schema import Block


def read_document(data: bytes) -> WordDocument:
    try:
        return Document(BytesIO(data))
    except KeyError:
        relationships = "word/_rels/document.xml.rels"
        with ZipFile(BytesIO(data)) as source:
            if relationships not in source.namelist():
                raise
            root = ElementTree.fromstring(source.read(relationships))
            parts = set(source.namelist())
            missing = []
            for link in root:
                role = link.get("Type", "").rsplit("/", 1)[-1]
                target = link.get("Target", "")
                if role not in {"footnotes", "endnotes"} or link.get("TargetMode") == "External":
                    continue
                part = target.lstrip("/") if target.startswith("/") else normpath(f"word/{target}")
                if part not in parts:
                    missing.append(link)
            if not missing:
                raise
            for link in missing:
                root.remove(link)
            repaired = BytesIO()
            with ZipFile(repaired, "w", ZIP_DEFLATED) as package:
                for name in source.namelist():
                    content = ElementTree.tostring(root) if name == relationships else source.read(name)
                    package.writestr(name, content)
        repaired.seek(0)
        return Document(repaired)


def parse_docx(data: bytes) -> list[Block]:
    blocks = []
    paragraph = table = 0
    for item in read_document(data).iter_inner_content():
        if isinstance(item, Table):
            table += 1
            seen = set()
            for row_number, row in enumerate(item.rows, 1):
                for column, cell in enumerate(row.cells, row.grid_cols_before + 1):
                    if cell._tc in seen:
                        continue
                    seen.add(cell._tc)
                    blocks.append(Block(
                        type="table_cell", text=cell.text.strip(),
                        locator={"table": table, "row": row_number, "column": column},
                    ))
        else:
            paragraph += 1
            if item.text.strip():
                blocks.append(Block(
                    type="paragraph", text=item.text.strip(), locator={"paragraph": paragraph},
                ))
    return blocks
