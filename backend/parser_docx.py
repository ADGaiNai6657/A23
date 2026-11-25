from io import BytesIO

from docx import Document
from docx.table import Table

from schema import Block


def parse_docx(data: bytes) -> list[Block]:
    blocks = []
    paragraph = table = 0
    for item in Document(BytesIO(data)).iter_inner_content():
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
