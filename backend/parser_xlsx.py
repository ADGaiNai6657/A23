from collections.abc import Iterator
from io import BytesIO

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from schema import Block


def workbook_rows(data: bytes) -> Iterator[tuple[str, int, dict[str, object]]]:
    workbook = load_workbook(BytesIO(data), read_only=True, data_only=True)
    try:
        for sheet in workbook:
            rows = sheet.iter_rows(values_only=True)
            headers = next(rows, ())
            keys = []
            for index, header in enumerate(headers, 1):
                name = str(header).strip() if header is not None else ""
                keys.append(name if name and name not in keys else get_column_letter(index))
            for number, row in enumerate(rows, 2):
                if any(value is not None and value != "" for value in row):
                    yield sheet.title, number, dict(zip(keys, row))
    finally:
        workbook.close()


def parse_xlsx(data: bytes) -> list[Block]:
    return [
        Block(
            type="spreadsheet_row",
            text="｜".join(f"{name}: {value}" for name, value in values.items() if value is not None),
            locator={"sheet": sheet, "row": row},
        )
        for sheet, row, values in workbook_rows(data)
    ]
