from schema import Block


def parse_text(data: bytes) -> list[Block]:
    try:
        text = data.decode("utf-8-sig")
    except UnicodeDecodeError:
        text = data.decode("gb18030")
    return [
        Block(type="text_line", text=line.strip(), locator={"line": number})
        for number, line in enumerate(text.splitlines(), 1)
        if line.strip()
    ]
