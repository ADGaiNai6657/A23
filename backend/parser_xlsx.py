from collections.abc import Iterator
from io import BytesIO
from itertools import zip_longest
from zipfile import ZIP_DEFLATED, ZipFile

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from schema import Block


STRICT_NAMESPACES = {
    b"http://purl.oclc.org/ooxml/spreadsheetml/main": b"http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships": b"http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    b"http://purl.oclc.org/ooxml/drawingml/main": b"http://schemas.openxmlformats.org/drawingml/2006/main",
}


def compatible_workbook(data: bytes) -> BytesIO:
    with ZipFile(BytesIO(data)) as source:
        workbook_xml = source.read("xl/workbook.xml")
        if not any(namespace in workbook_xml for namespace in STRICT_NAMESPACES):
            return BytesIO(data)
        output = BytesIO()
        with ZipFile(output, "w", ZIP_DEFLATED) as normalized:
            for name in source.namelist():
                content = source.read(name)
                if name.endswith((".xml", ".rels")):
                    for strict, compatible in STRICT_NAMESPACES.items():
                        content = content.replace(strict, compatible)
                normalized.writestr(name, content)
        output.seek(0)
        return output


def header_keys(headers: tuple) -> list[str]:
    names = [str(value).strip() if value is not None else "" for value in headers]
    reserved = set(filter(None, names))
    used = set()
    keys = []
    for index, name in enumerate(names, 1):
        if not name or name in used:
            name = get_column_letter(index)
            while name in reserved or name in used:
                name = f"列 {name}"
        used.add(name)
        keys.append(name)
    return keys


def workbook_rows(data: bytes) -> Iterator[tuple[str, int, dict[str, object]]]:
    workbook = load_workbook(compatible_workbook(data), read_only=True, data_only=True)
    try:
        if not workbook.worksheets:
            raise ValueError("工作簿中没有可读取的工作表")
        for sheet in workbook:
            # Some producers declare A1 as the used range even when XML contains many rows.
            sheet.reset_dimensions()
            rows = sheet.iter_rows(values_only=True)
            headers = next(rows, ())
            keys = header_keys(headers)
            for number, row in enumerate(rows, 2):
                if any(value is not None and value != "" for value in row):
                    if len(row) > len(keys):
                        keys = header_keys(headers + (None,) * (len(row) - len(headers)))
                    yield sheet.title, number, dict(zip_longest(keys, row))
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
